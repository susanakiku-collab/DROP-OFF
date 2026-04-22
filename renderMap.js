// renderMap
// dashboard.js から安全に分離した地図表示系

let dispatchOverviewMap = null;

let dispatchOverviewLayer = null;

let dispatchOverviewHost = null;

const dispatchOverviewFilterState = new Map();


let lastDispatchOverviewCards = [];

function getDispatchOverviewVehicleColor(index) {
  const colors = ["#3b82f6", "#22c55e", "#ef4444", "#f59e0b", "#a855f7", "#06b6d4", "#84cc16", "#ec4899"];
  return colors[Math.abs(Number(index) || 0) % colors.length];
}


function getVehicleDisplayName(vehicle) {
  return String(vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle?.id || ""}` || "車両").trim() || "車両";
}


function getDispatchRowLatLng(row) {
  const lat = Number(row?.casts?.latitude ?? row?.latitude ?? row?.lat);
  const lng = Number(row?.casts?.longitude ?? row?.longitude ?? row?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  return { lat, lng };
}


function buildVehicleRouteMapUrl(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) return "";
  const origin = `${ORIGIN_LAT},${ORIGIN_LNG}`;
  const points = rows
    .map(row => {
      const p = getDispatchRowLatLng(row);
      if (p) return `${p.lat},${p.lng}`;
      const address = String(row?.destination_address || row?.casts?.address || "").trim();
      return address || "";
    })
    .filter(Boolean);
  if (!points.length) return "";
  const destination = points[points.length - 1];
  const waypoints = points.slice(0, -1);
  const params = new URLSearchParams({
    api: "1",
    origin,
    destination,
    travelmode: "driving"
  });
  if (waypoints.length) params.set("waypoints", waypoints.join("|"));
  return `https://www.google.com/maps/dir/?${params.toString()}`;
}


function renderDispatchOverviewLegend(cards) {
  const host = document.getElementById("dispatchOverviewLegend");
  if (!host) return;
  host.innerHTML = "";
  (cards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    const isOn = dispatchOverviewFilterState.get(vehicleId) !== false;
    const color = getDispatchOverviewVehicleColor(index);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn ghost";
    button.style.cssText = `display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;border:1px solid ${isOn ? color : 'rgba(148,163,184,.35)'};background:${isOn ? 'rgba(15,23,42,.85)' : 'rgba(15,23,42,.45)'};color:${isOn ? '#fff' : '#94a3b8'};cursor:pointer;`;
    button.innerHTML = `<span style="display:inline-block;width:12px;height:12px;border-radius:999px;background:${color};box-shadow:0 0 0 2px rgba(255,255,255,.18) inset;"></span><span>${escapeHtml(getVehicleDisplayName(vehicle))}</span><span style="font-size:11px;opacity:.8;">${isOn ? '表示中' : '非表示'}</span>`;
    button.addEventListener("click", () => {
      dispatchOverviewFilterState.set(vehicleId, !isOn);
      renderDispatchOverviewLegend(lastDispatchOverviewCards);
      renderDispatchOverviewMap(lastDispatchOverviewCards);
    });
    host.appendChild(button);
  });
}


function renderDispatchOverviewMap(cards) {
  const host = document.getElementById("dispatchOverviewMap");
  lastDispatchOverviewCards = Array.isArray(cards) ? cards : [];
  renderDispatchOverviewLegend(lastDispatchOverviewCards);
  if (!host || !window.L) return;

  if (dispatchOverviewMap && dispatchOverviewHost !== host) {
    try { dispatchOverviewMap.remove(); } catch (_) {}
    dispatchOverviewMap = null;
    dispatchOverviewLayer = null;
  }

  dispatchOverviewHost = host;
  host.innerHTML = "";

  if (!dispatchOverviewMap) {
    dispatchOverviewMap = window.L.map(host, { preferCanvas: true, zoomControl: true });
    window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(dispatchOverviewMap);
  }

  if (dispatchOverviewLayer) {
    try { dispatchOverviewLayer.remove(); } catch (_) {}
  }
  dispatchOverviewLayer = window.L.layerGroup().addTo(dispatchOverviewMap);

  const bounds = [];
  const originMarker = window.L.circleMarker([ORIGIN_LAT, ORIGIN_LNG], {
    radius: 9,
    color: "#111827",
    fillColor: "#111827",
    fillOpacity: 1,
    weight: 2
  }).bindPopup(`<b>${escapeHtml(ORIGIN_LABEL || '起点')}</b>`);
  originMarker.addTo(dispatchOverviewLayer);
  bounds.push([ORIGIN_LAT, ORIGIN_LNG]);

  (lastDispatchOverviewCards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    if (dispatchOverviewFilterState.get(vehicleId) === false) return;
    const color = getDispatchOverviewVehicleColor(index);
    orderedRows.forEach((row, rowIndex) => {
      const point = getDispatchRowLatLng(row);
      if (!point) return;
      const pinNo = rowIndex + 1;
      const marker = window.L.marker([point.lat, point.lng], {
        icon: window.L.divIcon({
          className: 'dispatch-overview-pin',
          html: `<div style="width:26px;height:26px;border-radius:999px;background:${color};border:2px solid rgba(255,255,255,.96);box-shadow:0 6px 16px rgba(15,23,42,.38);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:12px;line-height:1;">${pinNo}</div>`,
          iconSize: [26, 26],
          iconAnchor: [13, 13],
          popupAnchor: [0, -14]
        })
      });
      const vehicleName = getVehicleDisplayName(vehicle);
      const castName = String(row?.casts?.name || "-");
      const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "-");
      const distanceKm = Number(row?.distance_km || 0).toFixed(1);
      marker.bindPopup(`<div style="min-width:180px;line-height:1.7;"><b>${escapeHtml(vehicleName)}-${pinNo}</b><br>キャスト: ${escapeHtml(castName)}<br>方面: ${escapeHtml(area)}<br>距離: ${escapeHtml(distanceKm)}km</div>`);
      marker.bindTooltip(`${escapeHtml(vehicleName)}-${pinNo}`, { direction: "top", opacity: 0.95 });
      marker.addTo(dispatchOverviewLayer);
      bounds.push([point.lat, point.lng]);
    });
  });

  window.setTimeout(() => {
    try {
      dispatchOverviewMap.invalidateSize();
      if (bounds.length >= 2) {
        dispatchOverviewMap.fitBounds(bounds, { padding: [24, 24] });
      } else if (bounds.length === 1) {
        dispatchOverviewMap.setView(bounds[0], 13);
      } else {
        dispatchOverviewMap.setView([ORIGIN_LAT, ORIGIN_LNG], 11);
      }
    } catch (error) {
      console.error("dispatchOverviewMap fit error:", error);
    }
  }, 50);
}

