// lastTrip
// dashboard.js から安全に分離した ラスト便UI / 判定補助系

function renderManualLastVehicleInfo() {
  const state = getManualLastVehicleState();
  const text = state?.vehicle_id
    ? `ラスト便車両: ${state.driver_name || "-"} / 帰宅:${normalizeAreaLabel(state.home_area || "無し")}`
    : "ラスト便車両: なし";
  const checkedNames = getDriverLastTripCheckedNames();
  const extra = checkedNames.length ? ` / ラスト便チェック: ${checkedNames.join("・")}` : "";
  if (els.manualLastVehicleInfo) els.manualLastVehicleInfo.textContent = text + extra;
}
function clearManualLastVehicle() {
  clearManualLastVehicleState();
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
  alert("ラスト便車両を解除しました");
}


function isManualLastVehicle(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  return Number(vehicleId) === Number(getManualLastVehicleId(dateStr));
}


function getDriverLastTripCheckedNames(dateStr = getCurrentDispatchDateStr()) {
  const state = getDriverLastTripState(dateStr);
  return allVehiclesCache
    .filter(vehicle => Boolean(state[String(Number(vehicle.id || 0))]))
    .map(vehicle => vehicle.driver_name || vehicle.plate_number || "-");
}



function isManualLastTripItem() {
  return false;
}


function moveManualLastItemsToEnd(rows) {
  return Array.isArray(rows) ? rows : [];
}


