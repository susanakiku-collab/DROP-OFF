
// ===== direction guard (auto inserted) =====
function __shouldBlockDirectionMerge(items, vehicles){
  try{
    const areas = new Set((items||[]).map(i=>{
      return (i.destination_area||i.cluster_area||i.area||'').toString().trim();
    }).filter(Boolean));
    return (vehicles?.length || 0) >= areas.size;
  }catch(e){return false;}
}
// ==========================================

// THEMIS AI Dispatch v6.9.12
const {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  ORIGIN_LABEL,
  ORIGIN_LAT,
  ORIGIN_LNG
} = window.APP_CONFIG;

const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
);

let currentUser = null;
let currentDispatchId = null;

let editingCastId = null;
let editingVehicleId = null;
let editingPlanId = null;
let editingActualId = null;

let allCastsCache = [];
let allVehiclesCache = [];
let currentPlansCache = [];
let currentActualsCache = [];
let currentDailyReportsCache = [];
let currentMileageExportRows = [];
let activeVehicleIdsForToday = new Set();
let lastAutoDispatchRunAtMinutes = null;
let simulationSlotHour = null;
let lastSimulationResult = null;
let isRefreshingHybridUI = false;
let suppressSimulationSlotChange = false;
const ENABLE_DISTANCE_REBALANCE = false;
const ENFORCE_AREA_PRIORITY_STRICT = true;
const ENABLE_DISPLAY_GROUP_FORCE_BRANCH = false;

const els = {
  plansTimeAreaMatrix: document.getElementById("plansTimeAreaMatrix"),
  userEmail: document.getElementById("userEmail"),
  sendLineBtn: document.getElementById("sendLineBtn"),
  originLabelText: document.getElementById("originLabelText"),
  logoutBtn: document.getElementById("logoutBtn"),

  exportAllBtn: document.getElementById("exportAllBtn"),
  importAllBtn: document.getElementById("importAllBtn"),
  importAllFileInput: document.getElementById("importAllFileInput"),
  openManualBtn: document.getElementById("openManualBtn"),
  dangerResetBtn: document.getElementById("dangerResetBtn"),
  resetCastsBtn: document.getElementById("resetCastsBtn"),
  resetVehiclesBtn: document.getElementById("resetVehiclesBtn"),

  homeCastCount: document.getElementById("homeCastCount"),
  homeVehicleCount: document.getElementById("homeVehicleCount"),
  homePlanCount: document.getElementById("homePlanCount"),
  homeActualCount: document.getElementById("homeActualCount"),
  homeDoneCount: document.getElementById("homeDoneCount"),
  homeCancelCount: document.getElementById("homeCancelCount"),
  homeMonthlyVehicleList: document.getElementById("homeMonthlyVehicleList"),
  resetMonthlySummaryBtn: document.getElementById("resetMonthlySummaryBtn"),

  castName: document.getElementById("castName"),
  castDistanceKm: document.getElementById("castDistanceKm"),
  castTravelMinutes: document.getElementById("castTravelMinutes"),
  fetchCastTravelMinutesBtn: document.getElementById("fetchCastTravelMinutesBtn"),
  castAddress: document.getElementById("castAddress"),
  castArea: document.getElementById("castArea"),
  castMemo: document.getElementById("castMemo"),
  castLatLngText: document.getElementById("castLatLngText"),
  castPhone: document.getElementById("castPhone"),
  castGeoStatus: document.getElementById("castGeoStatus"),
  castLat: document.getElementById("castLat"),
  castLng: document.getElementById("castLng"),
  saveCastBtn: document.getElementById("saveCastBtn"),
  guessAreaBtn: document.getElementById("guessAreaBtn"),
  openGoogleMapBtn: document.getElementById("openGoogleMapBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  importCsvBtn: document.getElementById("importCsvBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  csvFileInput: document.getElementById("csvFileInput"),
  castsTableBody: document.getElementById("castsTableBody"),
  castSearchName: document.getElementById("castSearchName"),
  castSearchArea: document.getElementById("castSearchArea"),
  castSearchAddress: document.getElementById("castSearchAddress"),
  castSearchPhone: document.getElementById("castSearchPhone"),
  castSearchRunBtn: document.getElementById("castSearchRunBtn"),
  castSearchResetBtn: document.getElementById("castSearchResetBtn"),
  castSearchCount: document.getElementById("castSearchCount"),
  castSearchResultWrap: document.getElementById("castSearchResultWrap"),

  vehiclePlateNumber: document.getElementById("vehiclePlateNumber"),
  vehicleArea: document.getElementById("vehicleArea"),
  vehicleHomeArea: document.getElementById("vehicleHomeArea"),
  vehicleSeatCapacity: document.getElementById("vehicleSeatCapacity"),
  vehicleDriverName: document.getElementById("vehicleDriverName"),
  vehicleLineId: document.getElementById("vehicleLineId"),
  vehicleStatus: document.getElementById("vehicleStatus"),
  vehicleMemo: document.getElementById("vehicleMemo"),
  saveVehicleBtn: document.getElementById("saveVehicleBtn"),
  cancelVehicleEditBtn: document.getElementById("cancelVehicleEditBtn"),
  importVehicleCsvBtn: document.getElementById("importVehicleCsvBtn"),
  exportVehicleCsvBtn: document.getElementById("exportVehicleCsvBtn"),
  vehicleCsvFileInput: document.getElementById("vehicleCsvFileInput"),
  vehiclesTableBody: document.getElementById("vehiclesTableBody"),
  mileageReportStartDate: document.getElementById("mileageReportStartDate"),
  mileageReportEndDate: document.getElementById("mileageReportEndDate"),
  previewMileageReportBtn: document.getElementById("previewMileageReportBtn"),
  exportMileageReportBtn: document.getElementById("exportMileageReportBtn"),
  mileageReportTableWrap: document.getElementById("mileageReportTableWrap"),

  dispatchDate: document.getElementById("dispatchDate"),
  optimizeBtn: document.getElementById("optimizeBtn"),
  confirmDailyBtn: document.getElementById("confirmDailyBtn"),
  clearActualBtn: document.getElementById("clearActualBtn"),
  checkAllVehiclesBtn: document.getElementById("checkAllVehiclesBtn"),
  uncheckAllVehiclesBtn: document.getElementById("uncheckAllVehiclesBtn"),
  clearManualLastVehicleBtn: document.getElementById("clearManualLastVehicleBtn"),
  dailyVehicleChecklist: document.getElementById("dailyVehicleChecklist"),
  manualLastVehicleInfo: document.getElementById("manualLastVehicleInfo"),
  dailyMileageInputs: document.getElementById("dailyMileageInputs"),
  saveDailyMileageBtn: document.getElementById("saveDailyMileageBtn"),
  copyResultBtn: document.getElementById("copyResultBtn"),
  dailyDispatchResult: document.getElementById("dailyDispatchResult"),
  operationContextSummary: document.getElementById("operationContextSummary"),
  operationDiagnosis: document.getElementById("operationDiagnosis"),
  simulationSlotSelect: document.getElementById("simulationSlotSelect"),
  simulationIncludePlanInflow: document.getElementById("simulationIncludePlanInflow"),
  runSimulationBtn: document.getElementById("runSimulationBtn"),
  runSimulationDispatchBtn: document.getElementById("runSimulationDispatchBtn"),
  simulationDiagnosis: document.getElementById("simulationDiagnosis"),
  simulationPreview: document.getElementById("simulationPreview"),

  planDate: document.getElementById("planDate"),
  exportPlansCsvBtn: document.getElementById("exportPlansCsvBtn"),
  importPlansCsvBtn: document.getElementById("importPlansCsvBtn"),
  plansCsvFileInput: document.getElementById("plansCsvFileInput"),
  clearPlansBtn: document.getElementById("clearPlansBtn"),
  planCastSelect: document.getElementById("planCastSelect"),
  planHour: document.getElementById("planHour"),
  planDistanceKm: document.getElementById("planDistanceKm"),
  planAddress: document.getElementById("planAddress"),
  planArea: document.getElementById("planArea"),
  planNote: document.getElementById("planNote"),
  savePlanBtn: document.getElementById("savePlanBtn"),
  guessPlanAreaBtn: document.getElementById("guessPlanAreaBtn"),
  cancelPlanEditBtn: document.getElementById("cancelPlanEditBtn"),
  plansGroupedTable: document.getElementById("plansGroupedTable"),
  planCastSuggest: document.getElementById("planCastSuggest"),

  actualDate: document.getElementById("actualDate"),
  addSelectedPlanBtn: document.getElementById("addSelectedPlanBtn"),
  copyActualTableBtn: document.getElementById("copyActualTableBtn"),
  planSelect: document.getElementById("planSelect"),
  castSelect: document.getElementById("castSelect"),
  castSuggest: document.getElementById("castSuggest"),
  actualHour: document.getElementById("actualHour"),
  actualDistanceKm: document.getElementById("actualDistanceKm"),
  actualStatus: document.getElementById("actualStatus"),
  actualAddress: document.getElementById("actualAddress"),
  actualArea: document.getElementById("actualArea"),
  actualNote: document.getElementById("actualNote"),
  saveActualBtn: document.getElementById("saveActualBtn"),
  guessActualAreaBtn: document.getElementById("guessActualAreaBtn"),
  cancelActualEditBtn: document.getElementById("cancelActualEditBtn"),
  actualTableWrap: document.getElementById("actualTableWrap"),
  actualTimeAreaMatrix: document.getElementById("actualTimeAreaMatrix"),

  historyList: document.getElementById("historyList")
};

function todayStr() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}


function getCurrentDispatchDateStr() {
  return els.dispatchDate?.value || els.actualDate?.value || els.planDate?.value || todayStr();
}

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

function renderManualLastVehicleInfo() {
  const state = getManualLastVehicleState();
  const text = state?.vehicle_id
    ? `ラスト便車両: ${state.driver_name || "-"} / 帰宅:${normalizeAreaLabel(state.home_area || "無し")}`
    : "ラスト便車両: なし";
  const checkedNames = getDriverLastTripCheckedNames();
  const extra = checkedNames.length ? ` / ラスト便チェック: ${checkedNames.join("・")}` : "";
  if (els.manualLastVehicleInfo) els.manualLastVehicleInfo.textContent = text + extra;
}function clearManualLastVehicle() {
  clearManualLastVehicleState();
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
  alert("ラスト便車両を解除しました");
}

function isManualLastVehicle(vehicleId, dateStr = getCurrentDispatchDateStr()) {
  return Number(vehicleId) === Number(getManualLastVehicleId(dateStr));
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

function setDriverLastTripChecked(vehicleId, checked, dateStr = getCurrentDispatchDateStr()) {
  const key = String(Number(vehicleId || 0));
  const state = getDriverLastTripState(dateStr);

  if (checked) state[key] = true;
  else delete state[key];

  saveDriverLastTripState(state, dateStr);
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


function formatDateTimeJa(value) {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value || "");
  return d.toLocaleString("ja-JP");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toNullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function normalizeStatus(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  if (status === "assigned") return "assigned";
  return "pending";
}

function getStatusText(status) {
  const s = normalizeStatus(status);
  if (s === "done") return "完了";
  if (s === "cancel") return "キャンセル";
  if (s === "assigned") return "配車済";
  return "未完了";
}

function getHourLabel(hour) {
  const n = Number(hour);
  return `${n}時`;
}

function getMonthKey(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthStartStr(dateStr) {
  const d = new Date(dateStr);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`;
}

function getDayOfWeek(dateStr) {
  const d = new Date(dateStr);
  return d.getDay();
}
function getDefaultLastHour(dateStr) {
  const today = todayStr();
  if (String(dateStr || '') === today) {
    return new Date().getHours();
  }
  return 0;
}

function getEffectiveLastHour(dateStr, rows = []) {
  const rowHours = (Array.isArray(rows) ? rows : [])
    .map(row => Number(row?.actual_hour ?? row?.plan_hour ?? row?.hour))
    .filter(Number.isFinite);
  if (rowHours.length) return Math.max(...rowHours);

  const cacheHours = [];
  (Array.isArray(currentActualsCache) ? currentActualsCache : []).forEach(row => {
    const status = normalizeStatus(row?.status);
    if (status === "done" || status === "cancel") return;
    const hour = Number(row?.actual_hour ?? row?.plan_hour);
    if (Number.isFinite(hour)) cacheHours.push(hour);
  });
  (Array.isArray(currentPlansCache) ? currentPlansCache : []).forEach(row => {
    const hour = Number(row?.plan_hour ?? row?.actual_hour);
    if (Number.isFinite(hour)) cacheHours.push(hour);
  });

  if (cacheHours.length) return Math.max(...cacheHours);
  return getDefaultLastHour(dateStr);
}


function getUnifiedHourSet() {
  const set = new Set();
  (Array.isArray(currentPlansCache) ? currentPlansCache : []).forEach(row => {
    const hour = Number(row?.plan_hour ?? row?.actual_hour);
    if (Number.isFinite(hour)) set.add(hour);
  });
  (Array.isArray(currentActualsCache) ? currentActualsCache : []).forEach(row => {
    const status = normalizeStatus(row?.status);
    if (status === "cancel") return;
    const hour = Number(row?.actual_hour ?? row?.plan_hour);
    if (Number.isFinite(hour)) set.add(hour);
  });
  if (!set.size) set.add(getDefaultLastHour(els.dispatchDate?.value || todayStr()));
  return [...set].sort((a, b) => a - b);
}

function getOperationBaseHour() {
  const activeHours = (Array.isArray(currentActualsCache) ? currentActualsCache : [])
    .filter(row => {
      const status = normalizeStatus(row?.status);
      return status !== "done" && status !== "cancel";
    })
    .map(row => Number(row?.actual_hour ?? row?.plan_hour))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);

  if (activeHours.length) return activeHours[0];

  const plannedHours = (Array.isArray(currentPlansCache) ? currentPlansCache : [])
    .map(row => Number(row?.plan_hour ?? row?.actual_hour))
    .filter(Number.isFinite)
    .sort((a, b) => a - b);

  if (plannedHours.length) return plannedHours[0];
  return getDefaultLastHour(els.dispatchDate?.value || todayStr());
}

function getNextHourSlot(baseHour) {
  const hours = getUnifiedHourSet();
  const found = hours.find(hour => Number(hour) > Number(baseHour));
  return Number.isFinite(found) ? found : null;
}

function getLastHourSlot() {
  const hours = getUnifiedHourSet();
  return hours.length ? hours[hours.length - 1] : null;
}

function parseClockTextToMinutes(value) {
  if (!value || value === "-") return null;
  const m = String(value).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

function buildProjectedRowsForHour(targetHour) {
  const hour = Number(targetHour);
  const actualRows = (Array.isArray(currentActualsCache) ? currentActualsCache : [])
    .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
    .filter(row => {
      const status = normalizeStatus(row?.status);
      return status !== "done" && status !== "cancel";
    });

  const actualPlanIds = new Set(actualRows.map(row => Number(row?.plan_id || 0)).filter(Boolean));
  const actualCastKeys = new Set(actualRows.map(row => `${Number(row?.cast_id || 0)}:${hour}`));
  const doneOrCancelPlanIds = new Set(
    (Array.isArray(currentActualsCache) ? currentActualsCache : [])
      .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
      .filter(row => {
        const status = normalizeStatus(row?.status);
        return status === "done" || status === "cancel";
      })
      .map(row => Number(row?.plan_id || 0))
      .filter(Boolean)
  );

  const planRows = (Array.isArray(currentPlansCache) ? currentPlansCache : [])
    .filter(plan => Number(plan?.plan_hour ?? -1) === hour)
    .filter(plan => !actualPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !doneOrCancelPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !actualCastKeys.has(`${Number(plan?.cast_id || 0)}:${hour}`))
    .map((plan, index) => ({
      id: -100000 - index - Number(plan?.id || 0),
      plan_id: Number(plan?.id || 0),
      cast_id: Number(plan?.cast_id || 0),
      actual_hour: hour,
      plan_hour: hour,
      status: "pending",
      distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
      destination_area: plan?.planned_area || plan?.casts?.area || "",
      destination_address: plan?.destination_address || plan?.casts?.address || "",
      note: plan?.note || "",
      casts: {
        ...(plan?.casts || {}),
        area: plan?.planned_area || plan?.casts?.area || "",
        address: plan?.destination_address || plan?.casts?.address || "",
        distance_km: Number(plan?.distance_km || plan?.casts?.distance_km || 0),
        travel_minutes: Number(plan?.casts?.travel_minutes || 0),
      },
      simulated_from_plan: true
    }));

  return [...actualRows, ...planRows];
}

function getEffectiveVehiclesForHour(targetHour) {
  const slotMinutes = Number(targetHour) * 60;
  const vehicles = Array.isArray(getSelectedVehiclesForToday()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
  const activeItems = (Array.isArray(currentActualsCache) ? currentActualsCache : []).filter(row => {
    const status = normalizeStatus(row?.status);
    return status !== "done" && status !== "cancel";
  });

  return vehicles.filter(vehicle => {
    const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
    if (!rows.length) return true;
    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
    const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
    if (readyMinutes === null) return true;
    return readyMinutes <= slotMinutes;
  });
}

function getVehicleDeadNamesForHour(targetHour) {
  const slotMinutes = Number(targetHour) * 60;
  const vehicles = Array.isArray(getSelectedVehiclesForToday()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
  const activeItems = (Array.isArray(currentActualsCache) ? currentActualsCache : []).filter(row => {
    const status = normalizeStatus(row?.status);
    return status !== "done" && status !== "cancel";
  });

  return vehicles.filter(vehicle => {
    const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
    if (!rows.length) return false;
    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
    const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
    return readyMinutes !== null && readyMinutes > slotMinutes;
  }).map(vehicle => vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle.id}`);
}

