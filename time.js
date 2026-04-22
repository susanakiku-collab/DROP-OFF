// time functions
// dashboard.js から安全に分離した日付・時刻系ユーティリティ
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


function buildMileageDateRange(startDate, endDate) {
  const dates = [];
  const current = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);

  while (current <= end) {
    dates.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

