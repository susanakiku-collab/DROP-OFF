// geo
// 位置情報・距離・地図表示・キャスト座標UI

const GOOGLE_GEOCODE_CACHE_KEY = "themis_google_geocode_cache_v1";
const GOOGLE_ROUTE_DISTANCE_CACHE_KEY = "themis_google_route_distance_cache_v1";
let lastCastGeocodeKey = "";
let castGeocodeSeq = 0;
let googleMapsApiPromise = null;

function isValidLatLng(lat, lng) {
  const latNum = Number(lat);
  const lngNum = Number(lng);
  return Number.isFinite(latNum) && Number.isFinite(lngNum) && Math.abs(latNum) <= 90 && Math.abs(lngNum) <= 180;
}

function normalizeGeocodeAddressKey(value = "") {
  return String(value || "")
    .replace(/[　\s]+/g, "")
    .replace(/^日本/, "")
    .replace(/^〒\d{3}-?\d{4}/, "")
    .trim()
    .toLowerCase();
}

function setCastGeoStatus(type = "idle", message = "") {
  if (!els?.castGeoStatus) return;
  const el = els.castGeoStatus;
  el.className = `geo-status ${type}`;
  el.textContent = message || "";
}

function scheduleCastAutoGeocode() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    if (els.castLat) els.castLat.value = "";
    if (els.castLng) els.castLng.value = "";
    if (els.castLatLngText) els.castLatLngText.value = "";
    lastCastGeocodeKey = "";
    castGeocodeSeq++;
    setCastGeoStatus("idle", "住所入力後 Enter で座標取得 / 未取得時は座標貼り付けから手動反映できます");
    return;
  }
  setCastGeoStatus("idle", "住所入力後 Enter で座標取得");
}

async function triggerCastAddressGeocodeNow() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    setCastGeoStatus("idle", "住所を入力してください");
    return null;
  }
  const runSeq = ++castGeocodeSeq;
  const currentKey = normalizeGeocodeAddressKey(address);
  setCastGeoStatus("loading", "取得中…");
  const result = await fillCastLatLngFromAddress({ silent: true, force: currentKey !== lastCastGeocodeKey });
  if (runSeq !== castGeocodeSeq) return result;
  if (result) {
    const sourceText = result.source === "cache" ? "キャッシュから取得" : result.source === "existing" ? "入力済み座標" : result.source === "google" ? "Google Geocoding" : "住所検索";
    setCastGeoStatus("success", `✔ 座標取得済 (${sourceText})`);
  } else {
    setCastGeoStatus("error", "座標を自動取得できませんでした。住所を確認して Enter で再試行するか、座標貼り付けで手動入力してください");
  }
  return result;
}

function loadGeocodeCache() {
  try {
    return JSON.parse(localStorage.getItem(GOOGLE_GEOCODE_CACHE_KEY) || "{}");
  } catch (_) {
    return {};
  }
}

function saveGeocodeCache(cache) {
  try {
    localStorage.setItem(GOOGLE_GEOCODE_CACHE_KEY, JSON.stringify(cache || {}));
  } catch (_) {}
}

function loadRouteDistanceCache() {
  try {
    return JSON.parse(localStorage.getItem(GOOGLE_ROUTE_DISTANCE_CACHE_KEY) || "{}");
  } catch (_) {
    return {};
  }
}

function saveRouteDistanceCache(cache) {
  try {
    localStorage.setItem(GOOGLE_ROUTE_DISTANCE_CACHE_KEY, JSON.stringify(cache || {}));
  } catch (_) {}
}

function ensureCastTravelMinutesUi() {
  if (document.getElementById("castTravelMinutes")) {
    els.castTravelMinutes = document.getElementById("castTravelMinutes");
    els.fetchCastTravelMinutesBtn = document.getElementById("fetchCastTravelMinutesBtn");
    if (els.fetchCastTravelMinutesBtn) {
      els.fetchCastTravelMinutesBtn.style.display = "none";
      els.fetchCastTravelMinutesBtn.disabled = true;
      els.fetchCastTravelMinutesBtn.title = "時間取得は使用しません";
    }
    return;
  }

  const distanceField = els.castDistanceKm?.closest?.(".field");
  if (distanceField?.parentElement) {
    const wrap = document.createElement("div");
    wrap.className = "field";
    wrap.innerHTML = `
      <label for="castTravelMinutes">片道予想時間(分)</label>
      <input id="castTravelMinutes" type="number" min="0" step="1" placeholder="例：35" />
    `;
    distanceField.insertAdjacentElement("afterend", wrap);
  }

  els.castTravelMinutes = document.getElementById("castTravelMinutes");
  els.fetchCastTravelMinutesBtn = document.getElementById("fetchCastTravelMinutesBtn");
  if (els.fetchCastTravelMinutesBtn) {
    els.fetchCastTravelMinutesBtn.style.display = "none";
    els.fetchCastTravelMinutesBtn.disabled = true;
    els.fetchCastTravelMinutesBtn.title = "時間取得は使用しません";
  }
}

function getStoredTravelMinutes(value) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.round(n) : 0;
}

function getCastTravelMinutesValue(castLike) {
  if (!castLike) return 0;
  return getStoredTravelMinutes(castLike.travel_minutes || castLike.travelMinutes);
}

function makeRouteDistanceCacheKey(address, lat, lng) {
  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);
  if (isValidLatLng(latNum, lngNum)) return `latlng:${latNum},${lngNum}`;
  return `addr:${normalizeGeocodeAddressKey(address)}`;
}

async function loadGoogleMapsApi() {
  if (window.google?.maps) return window.google.maps;
  if (googleMapsApiPromise) return googleMapsApiPromise;
  googleMapsApiPromise = Promise.resolve(null);
  return googleMapsApiPromise;
}

async function geocodeAddressGoogle(address) {
  const query = String(address || "").trim();
  if (!query) return null;

  await loadGoogleMapsApi();
  if (window.google?.maps?.Geocoder) {
    try {
      const geocoder = new google.maps.Geocoder();
      const result = await new Promise((resolve, reject) => {
        geocoder.geocode({ address: query, region: "JP" }, (results, status) => {
          if (status === "OK" && results?.length) resolve(results[0]);
          else reject(new Error(status || "GEOCODE_FAILED"));
        });
      });
      const loc = result?.geometry?.location;
      const lat = typeof loc?.lat === "function" ? loc.lat() : Number(loc?.lat);
      const lng = typeof loc?.lng === "function" ? loc.lng() : Number(loc?.lng);
      if (isValidLatLng(lat, lng)) return { lat, lng, source: "google" };
    } catch (_) {}
  }

  try {
    const url = new URL("https://nominatim.openstreetmap.org/search");
    url.searchParams.set("q", query);
    url.searchParams.set("format", "jsonv2");
    url.searchParams.set("limit", "1");
    url.searchParams.set("countrycodes", "jp");
    const res = await fetch(url.toString(), {
      headers: { "Accept": "application/json" }
    });
    if (!res.ok) return null;
    const data = await res.json();
    const row = Array.isArray(data) ? data[0] : null;
    const lat = Number(row?.lat);
    const lng = Number(row?.lon);
    if (isValidLatLng(lat, lng)) return { lat, lng, source: "nominatim" };
  } catch (_) {}
  return null;
}

async function fillCastLatLngFromAddress(options = {}) {
  const silent = Boolean(options?.silent);
  const force = Boolean(options?.force);
  const address = String(els.castAddress?.value || "").trim();
  if (!address) return null;

  const addressKey = normalizeGeocodeAddressKey(address);
  const currentLat = toNullableNumber(els.castLat?.value);
  const currentLng = toNullableNumber(els.castLng?.value);

  if (!force && isValidLatLng(currentLat, currentLng) && addressKey === lastCastGeocodeKey) {
    if (els.castLatLngText) els.castLatLngText.value = `${currentLat},${currentLng}`;
    return { lat: currentLat, lng: currentLng, source: "existing" };
  }

  const cache = loadGeocodeCache();
  const cached = cache[addressKey];
  if (!force && cached && isValidLatLng(cached.lat, cached.lng)) {
    if (els.castLat) els.castLat.value = cached.lat;
    if (els.castLng) els.castLng.value = cached.lng;
    if (els.castLatLngText) els.castLatLngText.value = `${cached.lat},${cached.lng}`;
    lastCastGeocodeKey = addressKey;
    return { lat: Number(cached.lat), lng: Number(cached.lng), source: "cache" };
  }

  const geocoded = await geocodeAddressGoogle(address);
  if (!geocoded || !isValidLatLng(geocoded.lat, geocoded.lng)) {
    if (!silent) setCastGeoStatus("error", "座標を取得できませんでした");
    return null;
  }

  if (els.castLat) els.castLat.value = geocoded.lat;
  if (els.castLng) els.castLng.value = geocoded.lng;
  if (els.castLatLngText) els.castLatLngText.value = `${geocoded.lat},${geocoded.lng}`;
  lastCastGeocodeKey = addressKey;
  cache[addressKey] = { lat: Number(geocoded.lat), lng: Number(geocoded.lng) };
  saveGeocodeCache(cache);

  try {
    const guessedArea = guessArea(Number(geocoded.lat), Number(geocoded.lng), address);
    if (els.castArea && guessedArea) els.castArea.value = normalizeAreaLabel(guessedArea);
    const distance = await resolveDistanceKmFromOrigin(address, geocoded.lat, geocoded.lng);
    if (distance !== null && els.castDistanceKm) els.castDistanceKm.value = String(distance);
  } catch (_) {}

  return { lat: Number(geocoded.lat), lng: Number(geocoded.lng), source: geocoded.source || "google" };
}

function parseLatLngText(text) {
  const raw = String(text || "").trim();
  if (!raw) return null;
  const m = raw.match(/(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)/);
  if (!m) return null;
  const lat = Number(m[1]);
  const lng = Number(m[2]);
  if (!isValidLatLng(lat, lng)) return null;
  return { lat, lng };
}

function haversineKm(lat1, lng1, lat2, lng2) {
  const toRad = d => (Number(d) * Math.PI) / 180;
  const R = 6371;
  const dLat = toRad(Number(lat2) - Number(lat1));
  const dLng = toRad(Number(lng2) - Number(lng1));
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function estimateRoadKmBetweenPoints(lat1, lng1, lat2, lng2) {
  const straight = haversineKm(lat1, lng1, lat2, lng2);
  return Number((straight * 1.25).toFixed(1));
}

function estimateRoadKmFromStation(lat, lng) {
  if (!isValidLatLng(lat, lng)) return 0;
  return estimateRoadKmBetweenPoints(ORIGIN_LAT, ORIGIN_LNG, lat, lng);
}

function getItemLatLng(item) {
  const lat = toNullableNumber(item?.casts?.latitude ?? item?.latitude ?? item?.lat);
  const lng = toNullableNumber(item?.casts?.longitude ?? item?.longitude ?? item?.lng);
  if (!isValidLatLng(lat, lng)) return null;
  return { lat, lng };
}

async function getGoogleDrivingDistanceKmFromOrigin(address, lat, lng) {
  const cacheKey = makeRouteDistanceCacheKey(address, lat, lng);
  if (!cacheKey || cacheKey === "addr:") return null;

  const cache = loadRouteDistanceCache();
  const cached = cache[cacheKey];
  if (Number.isFinite(Number(cached))) return Number(cached);

  await loadGoogleMapsApi();
  if (!window.google?.maps?.DirectionsService) return null;

  const destinationLat = toNullableNumber(lat);
  const destinationLng = toNullableNumber(lng);
  const destination = isValidLatLng(destinationLat, destinationLng)
    ? { lat: destinationLat, lng: destinationLng }
    : String(address || "").trim();

  if (!destination) return null;

  const km = await new Promise(resolve => {
    const service = new google.maps.DirectionsService();
    service.route({
      origin: { lat: ORIGIN_LAT, lng: ORIGIN_LNG },
      destination,
      travelMode: google.maps.TravelMode.DRIVING,
      region: "JP"
    }, (result, status) => {
      if (status === "OK") {
        const leg = result?.routes?.[0]?.legs?.[0];
        const meters = Number(leg?.distance?.value || 0);
        resolve(meters > 0 ? Number((meters / 1000).toFixed(1)) : null);
      } else {
        resolve(null);
      }
    });
  });

  if (Number.isFinite(Number(km))) {
    cache[cacheKey] = Number(km);
    saveRouteDistanceCache(cache);
    return Number(km);
  }
  return null;
}

async function resolveDistanceKmFromOrigin(address, lat, lng) {
  const cacheKey = makeRouteDistanceCacheKey(address, lat, lng);
  const cache = loadRouteDistanceCache();
  const cached = cache[cacheKey];
  if (Number.isFinite(Number(cached))) return Number(cached);

  const googleKm = await getGoogleDrivingDistanceKmFromOrigin(address, lat, lng);
  if (Number.isFinite(Number(googleKm))) return Number(googleKm);

  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);
  if (isValidLatLng(latNum, lngNum)) {
    const estimated = estimateRoadKmFromStation(latNum, lngNum);
    cache[cacheKey] = Number(estimated);
    saveRouteDistanceCache(cache);
    return Number(estimated);
  }
  return null;
}

async function resolveDistanceKmForCastRecord(cast, addressOverride = "") {
  const address = String(addressOverride || cast?.address || "").trim();
  const lat = toNullableNumber(cast?.latitude ?? cast?.lat);
  const lng = toNullableNumber(cast?.longitude ?? cast?.lng);

  const stored = toNullableNumber(cast?.distance_km);
  if (stored !== null) return stored;

  return await resolveDistanceKmFromOrigin(address, lat, lng);
}

function findCastByInputValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const byId = Number(raw);
  if (Number.isFinite(byId) && byId > 0) {
    const castById = (allCastsCache || []).find(c => Number(c.id) === byId);
    if (castById) return castById;
  }
  const lowered = raw.toLowerCase();
  return (allCastsCache || []).find(c => String(c.name || "").trim().toLowerCase() === lowered) || null;
}

function openGoogleMap(address = "") {
  const query = String(address || "").trim();
  if (!query) return;
  const url = `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL || "松戸駅")}&destination=${encodeURIComponent(query)}&travelmode=driving`;
  window.open(url, "_blank", "noopener,noreferrer");
}

function buildCastMapUrl(cast = {}) {
  const address = String(cast?.address || cast?.destination_address || cast?.pickup_address || '').trim();
  const lat = cast?.latitude ?? cast?.lat;
  const lng = cast?.longitude ?? cast?.lng;
  if (address) {
    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL || '松戸駅')}&destination=${encodeURIComponent(address)}&travelmode=driving`;
  }
  const nlat = Number(lat);
  const nlng = Number(lng);
  if (Number.isFinite(nlat) && Number.isFinite(nlng)) {
    return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(ORIGIN_LABEL || '松戸駅')}&destination=${encodeURIComponent(`${nlat},${nlng}`)}&travelmode=driving`;
  }
  return '';
}

function buildMapLinkHtml({ name = '', address = '', lat = null, lng = null, className = 'cast-name-link', suffix = ' 📍' } = {}) {
  const esc = (value) => {
    const str = String(value ?? '');
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  };

  const cast = { name, address, latitude: lat, longitude: lng };
  const url = buildCastMapUrl(cast);
  const label = esc(name || address || '-');
  if (!url) return label;
  return `<a href="${url}" target="_blank" rel="noopener noreferrer" class="${esc(className)}">${label}${esc(suffix)}</a>`;
}
