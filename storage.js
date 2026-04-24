// storage
// dashboard.js から安全に分離した localStorage 系ユーティリティ
function getManualLastVehicleStorageKey(dateStr = getCurrentDispatchDateStr()) {
  return `THEMIS_MANUAL_LAST_VEHICLE_${dateStr}`;
}


function getManualLastVehicleState(dateStr = getCurrentDispatchDateStr()) {
  try {
    const raw = window.localStorage.getItem(getManualLastVehicleStorageKey(dateStr));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    console.error(e);
    return null;
  }
}


function getManualLastVehicleId(dateStr = getCurrentDispatchDateStr()) {
  return Number(getManualLastVehicleState(dateStr)?.vehicle_id || 0);
}


function saveManualLastVehicleState(vehicle, dateStr = getCurrentDispatchDateStr()) {
  const payload = {
    vehicle_id: Number(vehicle?.id || 0),
    driver_name: vehicle?.driver_name || vehicle?.plate_number || "",
    home_area: normalizeAreaLabel(vehicle?.home_area || ""),
    vehicle_area: normalizeAreaLabel(vehicle?.vehicle_area || "")
  };
  window.localStorage.setItem(getManualLastVehicleStorageKey(dateStr), JSON.stringify(payload));
  renderManualLastVehicleInfo();
}


function clearManualLastVehicleState(dateStr = getCurrentDispatchDateStr()) {
  window.localStorage.removeItem(getManualLastVehicleStorageKey(dateStr));
  renderManualLastVehicleInfo();
}


function getDriverLastTripStorageKey(dateStr = getCurrentDispatchDateStr()) {
  return `THEMIS_DRIVER_LAST_TRIP_FLAGS_${dateStr}`;
}


function getDriverLastTripState(dateStr = getCurrentDispatchDateStr()) {
  try {
    const raw = window.localStorage.getItem(getDriverLastTripStorageKey(dateStr));
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    console.error(error);
    return {};
  }
}


function saveDriverLastTripState(state, dateStr = getCurrentDispatchDateStr()) {
  try {
    window.localStorage.setItem(
      getDriverLastTripStorageKey(dateStr),
      JSON.stringify(state || {})
    );
  } catch (error) {
    console.error(error);
  }
}


function isDriverLastTripChecked(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return Boolean(state[String(Number(vehicleId || 0))]);
}


function hasAnyDriverLastTripChecked(dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return Object.values(state).some(Boolean);
}


function clearDriverLastTripState(dateStr = getCurrentDispatchDateStr()) {
  try {
    window.localStorage.removeItem(getDriverLastTripStorageKey(dateStr));
  } catch (error) {
    console.error(error);
  }
}


function setDriverLastTripChecked(vehicleId, checked, dateStr = getCurrentDispatchDateStr()) {
  const key = String(Number(vehicleId || 0));
  const state = getDriverLastTripState(dateStr);

  if (checked) state[key] = true;
  else delete state[key];

  saveDriverLastTripState(state, dateStr);
}

