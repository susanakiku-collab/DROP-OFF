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

function formatCastDistanceForEditor(value) {
  if (typeof window.formatCastDistanceDisplay === 'function') return window.formatCastDistanceDisplay(value);
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? String(numeric) : '';
}

function sanitizeCastComputedDistanceKm(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) return null;
  const rounded = Number(numeric.toFixed(1));
  if (rounded > 3000) return null;
  return rounded;
}
window.sanitizeCastComputedDistanceKm = sanitizeCastComputedDistanceKm;

function castAddressNeedsRemoteGeocode(force = false) {
  const address = String(els.castAddress?.value || '').trim();
  if (!address) return false;
  const addressKey = normalizeGeocodeAddressKey(address);
  const currentLat = toNullableNumber(els.castLat?.value);
  const currentLng = toNullableNumber(els.castLng?.value);
  if (!force && isValidLatLng(currentLat, currentLng) && addressKey === lastCastGeocodeKey) {
    return false;
  }
  const cache = loadGeocodeCache();
  const cached = cache[addressKey];
  if (!force && cached && isValidLatLng(cached.lat, cached.lng)) {
    return false;
  }
  return true;
}

function scheduleCastAutoGeocode() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    if (els.castLat) els.castLat.value = "";
    if (els.castLng) els.castLng.value = "";
    if (els.castLatLngText) els.castLatLngText.value = "";
    lastCastGeocodeKey = "";
    castGeocodeSeq++;
    setCastGeoStatus("idle", "未取得 | 住所入力後に「APIで座標取得」または Enter。未取得時は座標貼り付けで手動反映できます");
    return;
  }
  setCastGeoStatus("idle", "取得待ち | 「APIで座標取得」または Enter で座標取得します");
}

async function triggerCastAddressGeocodeNow() {
  const address = String(els.castAddress?.value || "").trim();
  if (!address) {
    setCastGeoStatus("idle", "未取得 | 住所を入力してください");
    return null;
  }
  const runSeq = ++castGeocodeSeq;
  const currentKey = normalizeGeocodeAddressKey(address);
  const forceLookup = currentKey !== lastCastGeocodeKey;
  const needsRemoteLookup = castAddressNeedsRemoteGeocode(forceLookup);
  if (needsRemoteLookup && typeof ensureGoogleApiCoordinateLookupAccess === 'function') {
    const access = await ensureGoogleApiCoordinateLookupAccess({ actionLabel: 'API座標取得', consume: true });
    if (!access?.allowed) {
      setCastGeoStatus("error", access?.reason || "取得失敗 | 本日のAPI座標取得上限に達しました");
      return null;
    }
  }
  setCastGeoStatus("loading", needsRemoteLookup ? "取得中 | APIで座標取得しています..." : "取得中 | 保存済み座標を確認しています...");
  const result = await fillCastLatLngFromAddress({ silent: true, force: forceLookup });
  if (runSeq !== castGeocodeSeq) return result;
  if (result) {
    const sourceText = result.source === "cache"
      ? "APIで取得済み座標"
      : result.source === "existing"
        ? "入力済み座標"
        : result.source === "google"
          ? "APIで取得した座標"
          : "APIで取得 / 住所検索";
    if (result.metrics_ok === false) {
      setCastGeoStatus("error", `取得済 | ${sourceText} を反映しました / 距離計算は保留です。現在の起点座標を確認してください`);
    } else {
      setCastGeoStatus("success", `取得済 | ${sourceText} を反映しました`);
    }
  } else {
    setCastGeoStatus("error", "取得失敗 | 住所を確認して再試行するか、座標貼り付けで手動反映してください");
  }
  if (typeof refreshCastGoogleApiQuotaUi === 'function') refreshCastGoogleApiQuotaUi();
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
    if (els.castTravelMinutes) {
      try { els.castTravelMinutes.readOnly = true; } catch (_) {}
      try { els.castTravelMinutes.type = "text"; } catch (_) {}
    }
    if (els.fetchCastTravelMinutesBtn) {
      els.fetchCastTravelMinutesBtn.style.display = "";
      els.fetchCastTravelMinutesBtn.title = "住所から座標を取得します";
    }
    return;
  }

  const distanceField = els.castDistanceKm?.closest?.(".field");
  if (distanceField?.parentElement) {
    const wrap = document.createElement("div");
    wrap.className = "field";
    wrap.innerHTML = `
      <label for="castTravelMinutes">片道予想時間</label>
      <input id="castTravelMinutes" type="text" placeholder="座標から自動計算" readonly />
    `;
    distanceField.insertAdjacentElement("afterend", wrap);
  }

  els.castTravelMinutes = document.getElementById("castTravelMinutes");
  els.fetchCastTravelMinutesBtn = document.getElementById("fetchCastTravelMinutesBtn");
  if (els.castTravelMinutes) {
    try { els.castTravelMinutes.readOnly = true; } catch (_) {}
    try { els.castTravelMinutes.type = "text"; } catch (_) {}
  }
  if (els.fetchCastTravelMinutesBtn) {
    els.fetchCastTravelMinutesBtn.style.display = "";
    els.fetchCastTravelMinutesBtn.title = "住所から座標を取得します";
  }
}

function getStoredTravelMinutes(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return 0;

  const direct = Number(raw);
  if (Number.isFinite(direct) && direct > 0) return Math.round(direct);

  const hourMatch = raw.match(/(\d+)\s*時間/);
  const minuteMatch = raw.match(/(\d+)\s*分/);
  const hours = Number(hourMatch?.[1] || 0);
  const minutes = Number(minuteMatch?.[1] || 0);
  const total = (Number.isFinite(hours) ? hours : 0) * 60 + (Number.isFinite(minutes) ? minutes : 0);
  return total > 0 ? total : 0;
}

function getCurrentOriginLatLng() {
  const runtimeOrigin = typeof getCurrentOriginRuntime === "function" ? getCurrentOriginRuntime() : null;
  const lat = Number(runtimeOrigin?.lat ?? ORIGIN_LAT);
  const lng = Number(runtimeOrigin?.lng ?? ORIGIN_LNG);
  if (!isValidLatLng(lat, lng)) return null;
  return {
    lat,
    lng,
    name: String(runtimeOrigin?.name || ORIGIN_LABEL || "起点").trim() || "起点"
  };
}

const DIRECTION_UI_LABELS = [
  "北寄り",
  "北東寄り",
  "東寄り",
  "南東寄り",
  "南寄り",
  "南西寄り",
  "西寄り",
  "北西寄り"
];

function normalizeDirectionUiDegree(deg) {
  const num = Number(deg);
  if (!Number.isFinite(num)) return null;
  return ((num % 360) + 360) % 360;
}

function computeBearingDeg(originLat, originLng, targetLat, targetLng) {
  const lat1 = Number(originLat);
  const lng1 = Number(originLng);
  const lat2 = Number(targetLat);
  const lng2 = Number(targetLng);
  if (!isValidLatLng(lat1, lng1) || !isValidLatLng(lat2, lng2)) return null;

  const toRad = value => (value * Math.PI) / 180;
  const y = Math.sin(toRad(lng2 - lng1)) * Math.cos(toRad(lat2));
  const x =
    Math.cos(toRad(lat1)) * Math.sin(toRad(lat2)) -
    Math.sin(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.cos(toRad(lng2 - lng1));

  return normalizeDirectionUiDegree((Math.atan2(y, x) * 180) / Math.PI);
}

function angularDistanceDeg(a, b) {
  const left = normalizeDirectionUiDegree(a);
  const right = normalizeDirectionUiDegree(b);
  if (left == null || right == null) return Infinity;
  const diff = Math.abs(left - right);
  return Math.min(diff, 360 - diff);
}

function getDirectionUiOriginRuntime(originOverride = null) {
  const lat = toNullableNumber(originOverride?.lat);
  const lng = toNullableNumber(originOverride?.lng);
  if (isValidLatLng(lat, lng)) {
    return {
      lat,
      lng,
      name: String(originOverride?.name || ORIGIN_LABEL || "起点").trim() || "起点"
    };
  }
  return getCurrentOriginLatLng();
}

function getDirectionUiHourValue(item) {
  const raw = item?.hour ?? item?.actual_hour ?? item?.plan_hour ?? item?.dispatch_hour;
  const num = Number(raw);
  if (!Number.isFinite(num)) return null;
  const normalized = Math.trunc(num);
  if (normalized < 0 || normalized > 23) return null;
  return normalized;
}

function getDirectionUiDisplayName(item) {
  const name = item?.person_name || item?.cast_name || item?.name || item?.casts?.name || item?.display_name || "-";
  return String(name || "-").trim() || "-";
}

function directionLabelFromDeg(deg) {
  const normalized = normalizeDirectionUiDegree(deg);
  if (normalized == null) return "不明";
  const index = Math.floor((normalized + 22.5) / 45) % DIRECTION_UI_LABELS.length;
  return DIRECTION_UI_LABELS[index] || "不明";
}

function circularMeanDeg(values = []) {
  const safeValues = values.map(normalizeDirectionUiDegree).filter(v => v != null);
  if (!safeValues.length) return null;
  const toRad = value => (value * Math.PI) / 180;
  let sumSin = 0;
  let sumCos = 0;
  safeValues.forEach(value => {
    sumSin += Math.sin(toRad(value));
    sumCos += Math.cos(toRad(value));
  });
  if (Math.abs(sumSin) < 1e-9 && Math.abs(sumCos) < 1e-9) {
    return safeValues[0];
  }
  return normalizeDirectionUiDegree((Math.atan2(sumSin, sumCos) * 180) / Math.PI);
}

function normalizeDirectionUiSourceItem(item = {}, originOverride = null) {
  const origin = getDirectionUiOriginRuntime(originOverride);
  const lat = toNullableNumber(item?.destination_lat ?? item?.latitude ?? item?.lat ?? item?.casts?.latitude ?? item?.casts?.lat);
  const lng = toNullableNumber(item?.destination_lng ?? item?.longitude ?? item?.lng ?? item?.casts?.longitude ?? item?.casts?.lng);
  const hour = getDirectionUiHourValue(item);
  const distanceKm = toNullableNumber(item?.distance_km ?? item?.distanceKm ?? item?.casts?.distance_km);
  const rawMinutes = item?.travel_minutes ?? item?.travelMinutes ?? item?.casts?.travel_minutes;
  const travelMinutes = Number.isFinite(Number(rawMinutes)) ? Math.round(Number(rawMinutes)) : 0;
  const bearingDeg = origin && isValidLatLng(lat, lng) ? computeBearingDeg(origin.lat, origin.lng, lat, lng) : null;
  const normalizeArea = typeof normalizeAreaLabel === 'function' ? normalizeAreaLabel : (value => String(value || '無し'));
  return {
    __directionUiNormalized: true,
    source: item,
    id: item?.id != null ? String(item.id) : "",
    castId: item?.cast_id != null ? String(item.cast_id) : "",
    name: getDirectionUiDisplayName(item),
    hour,
    status: String(item?.status || "pending").trim() || "pending",
    distanceKm: distanceKm != null ? Number(distanceKm) : null,
    travelMinutes,
    lat,
    lng,
    bearingDeg,
    plannedArea: normalizeArea(item?.planned_area || item?.destination_area || item?.area || item?.casts?.area || '無し')
  };
}

function extractActiveHours(items = []) {
  const hourSet = new Set();
  (Array.isArray(items) ? items : []).forEach(item => {
    const hour = getDirectionUiHourValue(item);
    if (hour != null) hourSet.add(hour);
  });
  return Array.from(hourSet).sort((a, b) => a - b);
}

function sortDirectionUiItemsByBearing(items = []) {
  return [...items].sort((a, b) => {
    const left = normalizeDirectionUiDegree(a?.bearingDeg);
    const right = normalizeDirectionUiDegree(b?.bearingDeg);
    if (left == null && right == null) return 0;
    if (left == null) return 1;
    if (right == null) return -1;
    return left - right;
  });
}

function makeDirectionCluster(clusterItems = []) {
  const items = sortDirectionUiItemsByBearing(clusterItems);
  const centerDeg = circularMeanDeg(items.map(item => item?.bearingDeg));
  return {
    centerDeg,
    label: directionLabelFromDeg(centerDeg),
    items
  };
}

function mergeDirectionClusters(leftCluster, rightCluster) {
  return makeDirectionCluster([...(leftCluster?.items || []), ...(rightCluster?.items || [])]);
}

function buildDirectionClusters(items = [], originOverride = null, options = {}) {
  const origin = getDirectionUiOriginRuntime(originOverride);
  if (!origin) return [];

  const splitThresholdDeg = Number(options?.splitThresholdDeg || 35);
  const maxDirections = Math.max(1, Math.trunc(Number(options?.maxDirections || 6)));
  const normalizedItems = (Array.isArray(items) ? items : [])
    .map(item => item?.__directionUiNormalized ? item : normalizeDirectionUiSourceItem(item, origin))
    .filter(item => item && item.hour != null && isValidLatLng(item.lat, item.lng) && item.bearingDeg != null);

  if (!normalizedItems.length) return [];

  const sortedItems = sortDirectionUiItemsByBearing(normalizedItems);
  let clusters = [];

  sortedItems.forEach(item => {
    const lastCluster = clusters[clusters.length - 1];
    const lastItem = lastCluster?.items?.[lastCluster.items.length - 1];
    if (!lastCluster || angularDistanceDeg(lastItem?.bearingDeg, item?.bearingDeg) > splitThresholdDeg) {
      clusters.push(makeDirectionCluster([item]));
      return;
    }
    lastCluster.items.push(item);
    clusters[clusters.length - 1] = makeDirectionCluster(lastCluster.items);
  });

  if (clusters.length > 1) {
    const firstCluster = clusters[0];
    const lastCluster = clusters[clusters.length - 1];
    const firstItem = firstCluster?.items?.[0];
    const lastItem = lastCluster?.items?.[lastCluster.items.length - 1];
    if (angularDistanceDeg(firstItem?.bearingDeg, lastItem?.bearingDeg) <= splitThresholdDeg) {
      const merged = mergeDirectionClusters(lastCluster, firstCluster);
      clusters = [merged, ...clusters.slice(1, -1)];
    }
  }

  while (clusters.length > maxDirections) {
    let bestIndex = 0;
    let bestScore = Infinity;

    for (let i = 0; i < clusters.length; i += 1) {
      const nextIndex = (i + 1) % clusters.length;
      const score = angularDistanceDeg(clusters[i]?.centerDeg, clusters[nextIndex]?.centerDeg);
      if (score < bestScore) {
        bestScore = score;
        bestIndex = i;
      }
    }

    const nextIndex = (bestIndex + 1) % clusters.length;
    const merged = mergeDirectionClusters(clusters[bestIndex], clusters[nextIndex]);
    if (nextIndex === 0) {
      clusters = [merged, ...clusters.slice(1, bestIndex)];
    } else {
      clusters = [
        ...clusters.slice(0, bestIndex),
        merged,
        ...clusters.slice(nextIndex + 1)
      ];
    }
  }

  return clusters
    .map(cluster => makeDirectionCluster(cluster?.items || []))
    .sort((a, b) => {
      const left = normalizeDirectionUiDegree(a?.centerDeg);
      const right = normalizeDirectionUiDegree(b?.centerDeg);
      if (left == null && right == null) return 0;
      if (left == null) return 1;
      if (right == null) return -1;
      return left - right;
    })
    .map((cluster, index) => ({
      key: `dir_${index + 1}`,
      centerDeg: normalizeDirectionUiDegree(cluster.centerDeg),
      label: cluster.label,
      items: [...(cluster.items || [])],
      count: Number(cluster.items?.length || 0)
    }));
}

function buildTimeDirectionMatrix(items = [], originOverride = null, options = {}) {
  const origin = getDirectionUiOriginRuntime(originOverride);
  if (!origin) {
    return { origin: null, hours: [], directions: [], cells: {}, items: [] };
  }

  const normalizedItems = (Array.isArray(items) ? items : [])
    .map(item => item?.__directionUiNormalized ? item : normalizeDirectionUiSourceItem(item, origin))
    .filter(item => item && item.hour != null && isValidLatLng(item.lat, item.lng) && item.bearingDeg != null);

  const hours = extractActiveHours(normalizedItems);
  const clusters = buildDirectionClusters(normalizedItems, origin, options);
  const cells = {};

  clusters.forEach(cluster => {
    (cluster.items || []).forEach(item => {
      const cellKey = `${item.hour}__${cluster.key}`;
      if (!cells[cellKey]) cells[cellKey] = [];
      cells[cellKey].push({
        ...item,
        directionKey: cluster.key,
        directionLabel: cluster.label,
        directionCenterDeg: cluster.centerDeg
      });
    });
  });

  Object.keys(cells).forEach(key => {
    cells[key].sort((left, right) => {
      const distanceDiff = Number(right?.distanceKm || 0) - Number(left?.distanceKm || 0);
      if (Math.abs(distanceDiff) > 1e-9) return distanceDiff;
      return String(left?.name || '').localeCompare(String(right?.name || ''), 'ja');
    });
  });

  return {
    origin,
    hours,
    directions: clusters.map(cluster => ({
      key: cluster.key,
      label: cluster.label,
      centerDeg: cluster.centerDeg,
      count: cluster.count
    })),
    cells,
    items: normalizedItems
  };
}

function getEstimatedRoadMultiplier(straightKm) {
  const km = Math.max(0, Number(straightKm || 0));
  const freePlanAdjustment = 0.88;
  let baseMultiplier = 1.38;
  if (km <= 3) baseMultiplier = 1.18;
  else if (km <= 10) baseMultiplier = 1.28;
  return Number((baseMultiplier * freePlanAdjustment).toFixed(4));
}

function estimateTravelMinutesByDistance(distanceKm, areaInput = "") {
  if (typeof estimateFallbackTravelMinutes === "function") {
    return Math.max(0, Math.round(estimateFallbackTravelMinutes(distanceKm, areaInput)));
  }
  const km = Math.max(0, Number(distanceKm || 0));
  if (!km) return 0;
  const defaultSpeed = 28;
  return Math.max(1, Math.round((km / defaultSpeed) * 60));
}

function getCastOriginMetrics(castLike = {}, addressOverride = "", originOverride = null) {
  const lat = toNullableNumber(castLike?.latitude ?? castLike?.lat);
  const lng = toNullableNumber(castLike?.longitude ?? castLike?.lng);
  const area = normalizeAreaLabel(castLike?.area || "");
  if (!isValidLatLng(lat, lng)) {
    return {
      straight_km: null,
      distance_km: null,
      travel_minutes: null,
      area
    };
  }

  const overrideLat = toNullableNumber(originOverride?.lat);
  const overrideLng = toNullableNumber(originOverride?.lng);
  const origin = isValidLatLng(overrideLat, overrideLng)
    ? {
        lat: overrideLat,
        lng: overrideLng,
        name: String(originOverride?.name || ORIGIN_LABEL || "起点").trim() || "起点"
      }
    : getCurrentOriginLatLng();
  if (!origin) {
    return {
      straight_km: null,
      distance_km: null,
      travel_minutes: null,
      area
    };
  }

  const straightKm = haversineKm(origin.lat, origin.lng, lat, lng);
  const distanceKmRaw = Number((straightKm * getEstimatedRoadMultiplier(straightKm)).toFixed(1));
  const distanceKm = sanitizeCastComputedDistanceKm(distanceKmRaw);
  const travelMinutes = distanceKm != null ? estimateTravelMinutesByDistance(distanceKm, area || addressOverride || "") : null;

  return {
    straight_km: Number(straightKm.toFixed(1)),
    distance_km: distanceKm,
    travel_minutes: travelMinutes,
    area
  };
}

function getCastTravelMinutesValue(castLike) {
  if (!castLike) return 0;
  const dynamic = getCastOriginMetrics(castLike);
  const dynamicMinutes = getStoredTravelMinutes(dynamic?.travel_minutes);
  if (dynamicMinutes > 0) return dynamicMinutes;
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

function sanitizeGeocodeAddress(address) {
  let query = String(address || "").trim();
  if (!query) return "";
  query = query.replace(/[\u3000\t\r\n]+/g, " ");
  query = query.replace(/^[TＴ〒]\s*\d{3}-?\d{4}\s*/i, "");
  query = query.replace(/^\d{3}-?\d{4}\s*/, "");
  query = query.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  query = query.replace(/[‐‑‒–—―ー−]/g, "-");
  query = query.replace(/\s+/g, " ").trim();
  return query;
}

async function geocodeAddressGoogle(address) {
  const query = sanitizeGeocodeAddress(address);
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

  const tryJson = async (url) => {
    const res = await fetch(url, { headers: { "Accept": "application/json" } });
    if (!res.ok) throw new Error(`HTTP_${res.status}`);
    return res.json();
  };

  try {
    const url = new URL("https://msearch.gsi.go.jp/address-search/AddressSearch");
    url.searchParams.set("q", query);
    const data = await tryJson(url.toString());
    const row = Array.isArray(data) ? data[0] : (Array.isArray(data?.features) ? data.features[0] : null);
    const coords = Array.isArray(row?.geometry?.coordinates) ? row.geometry.coordinates : [];
    const lng = Number(coords?.[0]);
    const lat = Number(coords?.[1]);
    if (isValidLatLng(lat, lng)) return { lat, lng, source: "gsi" };
  } catch (_) {}

  try {
    const url = new URL("https://nominatim.openstreetmap.org/search");
    url.searchParams.set("q", query);
    url.searchParams.set("format", "jsonv2");
    url.searchParams.set("limit", "1");
    url.searchParams.set("countrycodes", "jp");
    const data = await tryJson(url.toString());
    const row = Array.isArray(data) ? data[0] : null;
    const lat = Number(row?.lat);
    const lng = Number(row?.lon);
    if (isValidLatLng(lat, lng)) return { lat, lng, source: "nominatim" };
  } catch (_) {}

  try {
    const url = new URL("https://photon.komoot.io/api/");
    url.searchParams.set("q", query);
    url.searchParams.set("lang", "ja");
    url.searchParams.set("limit", "1");
    const data = await tryJson(url.toString());
    const feat = Array.isArray(data?.features) ? data.features[0] : null;
    const coords = Array.isArray(feat?.geometry?.coordinates) ? feat.geometry.coordinates : [];
    const lng = Number(coords?.[0]);
    const lat = Number(coords?.[1]);
    if (isValidLatLng(lat, lng)) return { lat, lng, source: "photon" };
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

  let metrics = null;
  try {
    const guessedArea = guessArea(Number(geocoded.lat), Number(geocoded.lng), address);
    if (els.castArea && guessedArea) els.castArea.value = normalizeAreaLabel(guessedArea);
    metrics = getCastOriginMetrics({ latitude: geocoded.lat, longitude: geocoded.lng, area: els.castArea?.value || guessedArea || "" }, address);
    if (els.castDistanceKm) els.castDistanceKm.value = formatCastDistanceForEditor(metrics?.distance_km);
    if (els.castTravelMinutes) {
      els.castTravelMinutes.value = metrics?.travel_minutes != null
        ? (typeof formatMinutesAsJa === 'function' ? formatMinutesAsJa(metrics.travel_minutes) : String(metrics.travel_minutes))
        : '';
    }
    if (typeof updateCastDistanceHint === 'function') {
      updateCastDistanceHint({ distanceKm: metrics?.distance_km, invalidDistance: metrics?.distance_km == null });
    }
  } catch (_) {}

  return {
    lat: Number(geocoded.lat),
    lng: Number(geocoded.lng),
    source: geocoded.source || "google",
    metrics_ok: metrics?.distance_km != null,
    metrics_warning: metrics?.distance_km == null ? '距離計算に失敗しました。現在の起点座標を確認してください' : ''
  };
}

function parseLatLngText(text) {
  const raw = String(text || "").trim();
  if (!raw) return null;

  const patterns = [
    /(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)/,
    /@(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/,
    /!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/,
    /[?&](?:q|ll|query|center|destination)=(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)/
  ];

  for (const pattern of patterns) {
    const m = raw.match(pattern);
    if (!m) continue;
    const lat = Number(m[1]);
    const lng = Number(m[2]);
    if (isValidLatLng(lat, lng)) return { lat, lng };
  }

  return null;
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
  return Number((straight * getEstimatedRoadMultiplier(straight)).toFixed(1));
}

function estimateRoadKmFromStation(lat, lng) {
  if (!isValidLatLng(lat, lng)) return 0;
  const origin = getCurrentOriginLatLng();
  if (!origin) return 0;
  return estimateRoadKmBetweenPoints(origin.lat, origin.lng, lat, lng);
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

async function resolveDistanceKmFromOrigin(address, lat, lng, originOverride = null) {
  const latNum = toNullableNumber(lat);
  const lngNum = toNullableNumber(lng);
  if (!isValidLatLng(latNum, lngNum)) return null;

  const dynamic = getCastOriginMetrics({ latitude: latNum, longitude: lngNum }, address, originOverride);
  if (dynamic?.distance_km != null) return Number(dynamic.distance_km);

  return Number(estimateRoadKmFromStation(latNum, lngNum));
}

async function resolveDistanceKmForCastRecord(cast, addressOverride = "", originOverride = null) {
  const address = String(addressOverride || cast?.address || "").trim();
  const lat = toNullableNumber(cast?.latitude ?? cast?.lat);
  const lng = toNullableNumber(cast?.longitude ?? cast?.lng);

  const dynamic = getCastOriginMetrics(cast, address, originOverride);
  if (dynamic?.distance_km != null) return Number(dynamic.distance_km);

  return await resolveDistanceKmFromOrigin(address, lat, lng, originOverride);
}

function findCastByInputValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;

  const casts = Array.isArray(allCastsCache) ? allCastsCache : [];
  if (!casts.length) return null;

  const lowered = raw.toLowerCase();

  // 1) UUID or exact id string
  const byExactId = casts.find(c => String(c.id || "").trim().toLowerCase() === lowered);
  if (byExactId) return byExactId;

  // 2) legacy numeric id support
  const byId = Number(raw);
  if (Number.isFinite(byId) && byId > 0) {
    const castByNumericId = casts.find(c => Number(c.id) === byId);
    if (castByNumericId) return castByNumericId;
  }

  // 3) exact name
  const byExactName = casts.find(c => String(c.name || "").trim().toLowerCase() === lowered);
  if (byExactName) return byExactName;

  // 4) exact searchable composite text shown in candidates
  const byComposite = casts.find(c => getCastSearchText(c).trim().toLowerCase() === lowered);
  if (byComposite) return byComposite;

  // 5) first token before separator, e.g. "るな / 牛久..."
  const head = raw.split('/')[0].trim().toLowerCase();
  if (head) {
    const byHeadName = casts.find(c => String(c.name || "").trim().toLowerCase() === head);
    if (byHeadName) return byHeadName;
  }

  // 6) unique partial match across name/address/area
  const candidates = casts.filter(c => {
    const hay = [
      String(c.name || "").trim(),
      String(c.address || "").trim(),
      String(c.area || "").trim(),
      getCastSearchText(c)
    ].join(' / ').toLowerCase();
    return hay.includes(lowered);
  });
  if (candidates.length === 1) return candidates[0];

  return null;
}

function openGoogleMap(address = "", lat = null, lng = null) {
  const query = String(address || "").trim();
  const destLat = Number(lat);
  const destLng = Number(lng);
  const runtimeOrigin = getCurrentOriginLatLng();

  const originValue = runtimeOrigin && isValidLatLng(runtimeOrigin.lat, runtimeOrigin.lng)
    ? `${runtimeOrigin.lat},${runtimeOrigin.lng}`
    : String(runtimeOrigin?.name || ORIGIN_LABEL || "起点").trim() || "起点";

  const destinationValue = isValidLatLng(destLat, destLng)
    ? `${destLat},${destLng}`
    : query;

  if (!destinationValue) return;

  const url = `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(originValue)}&destination=${encodeURIComponent(destinationValue)}&travelmode=driving`;
  window.open(url, "_blank", "noopener,noreferrer");
}
