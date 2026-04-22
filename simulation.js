// simulation
// dashboard.js から安全に分離した シミュレーション診断 / 診断UI系

function getSimulationStateValue(key, fallback) {
  try {
    if (window.stateManager?.get) {
      const value = window.stateManager.get(key);
      return value === undefined ? fallback : value;
    }
  } catch (_) {}
  return fallback;
}

function setSimulationSlotHourSafe(value) {
  if (typeof window.setSimulationSlotHourState === "function") {
    return window.setSimulationSlotHourState(value);
  }
  if (window.stateManager?.setSimulationSlotHour) {
    return window.stateManager.setSimulationSlotHour(value);
  }
  window.simulationSlotHour = value;
  return value;
}

function getSimulationSlotHourSafe() {
  return getSimulationStateValue("simulationSlotHour", window.simulationSlotHour ?? null);
}

function setLastSimulationResultSafe(value) {
  if (typeof window.setLastSimulationResultState === "function") {
    return window.setLastSimulationResultState(value);
  }
  if (window.stateManager?.setLastSimulationResult) {
    return window.stateManager.setLastSimulationResult(value);
  }
  window.lastSimulationResult = value;
  return value;
}

function getLastSimulationResultSafe() {
  return getSimulationStateValue("lastSimulationResult", window.lastSimulationResult ?? null);
}

function setIsRefreshingHybridUISafe(value) {
  if (typeof window.setIsRefreshingHybridUIState === "function") {
    return window.setIsRefreshingHybridUIState(value);
  }
  if (window.stateManager?.setIsRefreshingHybridUI) {
    return window.stateManager.setIsRefreshingHybridUI(value);
  }
  window.isRefreshingHybridUI = Boolean(value);
  return Boolean(value);
}

function getIsRefreshingHybridUISafe() {
  return Boolean(getSimulationStateValue("isRefreshingHybridUI", window.isRefreshingHybridUI ?? false));
}

function setSuppressSimulationSlotChangeSafe(value) {
  if (typeof window.setSuppressSimulationSlotChangeState === "function") {
    return window.setSuppressSimulationSlotChangeState(value);
  }
  if (window.stateManager?.setSuppressSimulationSlotChange) {
    return window.stateManager.setSuppressSimulationSlotChange(value);
  }
  window.suppressSimulationSlotChange = Boolean(value);
  return Boolean(value);
}

function getSuppressSimulationSlotChangeSafe() {
  return Boolean(getSimulationStateValue("suppressSimulationSlotChange", window.suppressSimulationSlotChange ?? false));
}


function __simulationCallDispatchCoreSafe(functionName, ...args) {
  if (typeof window.callDispatchCore === "function") {
    return window.callDispatchCore(functionName, ...args);
  }
  const core = window.DispatchCore || {};
  const fn = typeof window[functionName] === "function" ? window[functionName] : core?.[functionName];
  if (typeof fn !== "function") {
    if (functionName === "getVehicleDeadNamesForHour") return [];
    if (functionName === "getEffectiveVehiclesForHour") return [];
    if (functionName === "buildProjectedRowsForHour" || functionName === "buildSimulationRowsForHour") return functionName === "buildSimulationRowsForHour" ? { rows: [], summary: { slotPlanCount: 0, inflowPlanCount: 0 } } : [];
    return [];
  }
  return fn(...args);
}

function getBuildSimulationRowsForHourSafe(targetHour, options = {}) {
  return __simulationCallDispatchCoreSafe("buildSimulationRowsForHour", targetHour, options);
}

function getBuildProjectedRowsForHourSafe(targetHour) {
  return __simulationCallDispatchCoreSafe("buildProjectedRowsForHour", targetHour);
}

function getEffectiveVehiclesForHourSafe(targetHour) {
  return __simulationCallDispatchCoreSafe("getEffectiveVehiclesForHour", targetHour);
}

function getVehicleDeadNamesForHourSafe(targetHour) {
  return __simulationCallDispatchCoreSafe("getVehicleDeadNamesForHour", targetHour);
}

function diagnoseSimulationHourWindow(targetHour, options = {}) {
  const built = getBuildSimulationRowsForHourSafe(targetHour, options);
  const rows = built.rows;
  const vehicles = getEffectiveVehiclesForHourSafe(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHourSafe(targetHour),
    slotPlanCount: built.summary.slotPlanCount,
    inflowPlanCount: built.summary.inflowPlanCount
  };
}


function diagnoseHourWindow(targetHour) {
  const rows = getBuildProjectedRowsForHourSafe(targetHour);
  const vehicles = getEffectiveVehiclesForHourSafe(targetHour);
  const effectiveVehicleCount = vehicles.length;
  const totalCapacity = vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
  const castCount = rows.length;
  const areaGroups = new Set(rows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し"))));
  const shortageCount = Math.max(0, castCount - totalCapacity);
  let statusKey = "ok";
  let statusText = "OK";
  if (totalCapacity < castCount) {
    statusKey = "danger";
    statusText = "定員オーバー見込み";
  } else if (areaGroups.size > effectiveVehicleCount && castCount > 0) {
    statusKey = "warn";
    statusText = "車両不足見込み";
  }
  return {
    hour: Number(targetHour),
    rows,
    vehicles,
    castCount,
    effectiveVehicleCount,
    totalCapacity,
    areaGroupCount: areaGroups.size,
    shortageCount,
    statusKey,
    statusText,
    deadVehicleNames: getVehicleDeadNamesForHourSafe(targetHour)
  };
}


function getStatusPillClass(statusKey) {
  if (statusKey === "danger") return "danger";
  if (statusKey === "warn") return "warn";
  return "ok";
}


function buildStatusPill(label, value, statusKey = "ok") {
  return `<span class="status-pill ${getStatusPillClass(statusKey)}"><span class="status-pill-label">${escapeHtml(label)}</span><span class="status-pill-value">${escapeHtml(value)}</span></span>`;
}


function buildSummaryItemsHtml(items) {
  return items.map(item => `
    <div class="hybrid-summary-item ${escapeHtml(item.tone || '')}">
      <div class="hybrid-summary-icon">${escapeHtml(item.icon || '•')}</div>
      <div class="hybrid-summary-body">
        <div class="hybrid-summary-label">${escapeHtml(item.label)}</div>
        <div class="hybrid-summary-value">${escapeHtml(item.value)}</div>
      </div>
    </div>
  `).join("");
}


function buildOperationDiagnosisHtml(operationDiag, nextDiag) {
  const nextDead = nextDiag?.deadVehicleNames?.length ? nextDiag.deadVehicleNames.join(" / ") : "なし";
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("今", `${getHourLabel(operationDiag.hour)} / ${operationDiag.castCount}名`, operationDiag.statusKey)}
      ${buildStatusPill("次", nextDiag ? `${getHourLabel(nextDiag.hour)} / ${nextDiag.castCount}名` : '-', nextDiag?.statusKey || 'ok')}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便の状態</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDiag ? nextDiag.statusText : '-')}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(nextDead)}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便 実効/定員</div>
        <div class="hybrid-legend-value">${nextDiag ? `${nextDiag.effectiveVehicleCount}台 / ${nextDiag.totalCapacity}` : '-'}</div>
      </div>
    </div>
  `;
}


function buildSimulationDiagnosisHtml(diag) {
  return `
    <div class="hybrid-state-row">
      ${buildStatusPill("便", `${getHourLabel(diag.hour)}`, diag.statusKey)}
      ${buildStatusPill("人数", `${diag.castCount}名`, diag.statusKey)}
      ${buildStatusPill("車両", `${diag.effectiveVehicleCount}台`, diag.statusKey)}
      ${buildStatusPill("定員", `${diag.totalCapacity}`, diag.statusKey)}
    </div>
    <div class="hybrid-legend-grid">
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">判定</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.statusText)}${diag.shortageCount > 0 ? ` / 未配車見込み ${diag.shortageCount}名` : ''}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">方面系統</div>
        <div class="hybrid-legend-value">${diag.areaGroupCount}</div>
      </div>
      <div class="hybrid-legend-card">
        <div class="hybrid-legend-title">次便NG車両</div>
        <div class="hybrid-legend-value">${escapeHtml(diag.deadVehicleNames.length ? diag.deadVehicleNames.join(' / ') : 'なし')}</div>
      </div>
    </div>
  `;
}


function renderOperationAndSimulationUI() {
  if (getIsRefreshingHybridUISafe()) return;
  setIsRefreshingHybridUISafe(true);
  try {
    const operationBaseHour = getOperationBaseHour();
    const realNextHour = getNextHourSlot(operationBaseHour);
    const operationDiag = diagnoseHourWindow(operationBaseHour);
    const nextDiag = Number.isFinite(realNextHour) ? diagnoseHourWindow(realNextHour) : null;

    if (els.operationContextSummary) {
      els.operationContextSummary.innerHTML = buildSummaryItemsHtml([
        { icon: "📍", label: "今の基準", value: getHourLabel(operationBaseHour), tone: "tone-blue" },
        { icon: "⏭", label: "次に見る便", value: Number.isFinite(realNextHour) ? getHourLabel(realNextHour) : "-", tone: "tone-indigo" },
        { icon: "🚐", label: "次便の使える車両", value: nextDiag ? `${nextDiag.effectiveVehicleCount}台` : "-", tone: "tone-green" }
      ]);
    }

    if (els.operationDiagnosis) {
      const cls = nextDiag?.statusKey === "danger"
        ? "danger"
        : (nextDiag?.statusKey === "warn" ? "warn" : "ok");
      els.operationDiagnosis.className = `hybrid-diagnosis ${cls}`;
      els.operationDiagnosis.innerHTML = buildOperationDiagnosisHtml(operationDiag, nextDiag);
    }

    const hourOptions = getUnifiedHourSet();
    if (els.simulationSlotSelect) {
      const currentSimulationSlotHour = getSimulationSlotHourSafe();
      const previous = Number.isFinite(currentSimulationSlotHour) ? currentSimulationSlotHour : (Number.isFinite(realNextHour) ? realNextHour : operationBaseHour);
      setSuppressSimulationSlotChangeSafe(true);
      try {
        const optionHtml = hourOptions.map(hour => `<option value="${hour}">${getHourLabel(hour)}</option>`).join("");
        if (els.simulationSlotSelect.innerHTML !== optionHtml) {
          els.simulationSlotSelect.innerHTML = optionHtml;
        }
        const targetValue = hourOptions.includes(previous) ? previous : (hourOptions[0] ?? operationBaseHour);
        setSimulationSlotHourSafe(targetValue);
        if (String(els.simulationSlotSelect.value) !== String(targetValue)) {
          els.simulationSlotSelect.value = String(targetValue);
        }
      } finally {
        setTimeout(() => { setSuppressSimulationSlotChangeSafe(false); }, 0);
      }
    }

    if (!getLastSimulationResultSafe() && els.simulationDiagnosis) {
      els.simulationDiagnosis.className = 'hybrid-diagnosis muted';
      els.simulationDiagnosis.textContent = '試算対象便を選択してください';
    }
  } catch (error) {
    console.error(error);
    if (els.operationDiagnosis) {
      els.operationDiagnosis.className = 'hybrid-diagnosis danger';
      els.operationDiagnosis.textContent = `診断エラー: ${error?.message || error}`;
    }
  } finally {
    setIsRefreshingHybridUISafe(false);
  }
}


function runSlotDiagnosisPreview() {
  const hour = Number(els.simulationSlotSelect?.value ?? getSimulationSlotHourSafe() ?? getOperationBaseHour());
  setSimulationSlotHourSafe(hour);
  const diag = diagnoseSimulationHourWindow(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  setLastSimulationResultSafe({ type: 'diagnosis', hour, diag });
  if (els.simulationDiagnosis) {    els.simulationDiagnosis.className = `hybrid-diagnosis ${diag.statusKey}`;
    els.simulationDiagnosis.innerHTML = buildSimulationDiagnosisHtml(diag);
  }
  if (els.simulationPreview) {
    if (!diag.rows.length) {
      els.simulationPreview.className = 'simulation-preview muted';
      els.simulationPreview.textContent = 'この便の予定対象はありません';
    } else {
      const list = diag.rows.map(row => `
        <div class="sim-mini-chip">
          <span class="sim-mini-name">${escapeHtml(row?.casts?.name || '-')}</span>
          <span class="sim-mini-area">${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span>
        </div>
      `).join('');
      els.simulationPreview.className = 'simulation-preview';
      els.simulationPreview.innerHTML = `
        <div class="sim-preview-head">
          <h4 class="sim-preview-title">試算対象一覧</h4>
          <span class="chip">${getHourLabel(diag.hour)}</span>
        </div>
        <div class="sim-preview-meta">対象便予定 ${diag.slotPlanCount}名 / 前便未処理予定 ${diag.inflowPlanCount}名</div>
        <div class="sim-mini-chip-grid">${list}</div>
      `;
    }
  }
}



function __simulationNormalizePendingRows(rows, targetHour) {
  return (Array.isArray(rows) ? rows : []).map((plan, index) => ({
    id: Number(plan?.id || 0) > 0 ? (-1000000 - Number(targetHour || 0) * 1000 - index - Number(plan?.id || 0)) : (-1000000 - Number(targetHour || 0) * 1000 - index),
    plan_id: Number(plan?.id || 0),
    cast_id: Number(plan?.cast_id || 0),
    actual_hour: Number(targetHour),
    plan_hour: Number(plan?.plan_hour ?? targetHour),
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
}

function __simulationBuildProjectedRowsForHourFallback(targetHour) {
  const hour = Number(targetHour);
  const actualRows = (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : [])
    .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
    .filter(row => {
      const status = normalizeStatus(row?.status);
      return status !== "done" && status !== "cancel";
    });

  const actualPlanIds = new Set(actualRows.map(row => Number(row?.plan_id || 0)).filter(Boolean));
  const actualCastKeys = new Set(actualRows.map(row => `${Number(row?.cast_id || 0)}:${hour}`));
  const doneOrCancelPlanIds = new Set(
    (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : [])
      .filter(row => Number(row?.actual_hour ?? row?.plan_hour) === hour)
      .filter(row => {
        const status = normalizeStatus(row?.status);
        return status === "done" || status === "cancel";
      })
      .map(row => Number(row?.plan_id || 0))
      .filter(Boolean)
  );

  const planRows = (Array.isArray(typeof currentPlansCache !== "undefined" ? currentPlansCache : window.currentPlansCache) ? (typeof currentPlansCache !== "undefined" ? currentPlansCache : window.currentPlansCache) : [])
    .filter(plan => Number(plan?.plan_hour ?? -1) === hour)
    .filter(plan => !actualPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !doneOrCancelPlanIds.has(Number(plan?.id || 0)))
    .filter(plan => !actualCastKeys.has(`${Number(plan?.cast_id || 0)}:${hour}`));

  return [...actualRows, ...__simulationNormalizePendingRows(planRows, hour)];
}

function __simulationBuildSimulationRowsForHourFallback(targetHour, options = {}) {
  const hour = Number(targetHour);
  const slotRows = __simulationBuildProjectedRowsForHourFallback(hour);
  const slotPlanIds = new Set(slotRows.map(row => Number(row?.plan_id || 0)).filter(Boolean));
  const slotCount = slotRows.filter(row => Boolean(row?.simulated_from_plan)).length;
  let inflowRows = [];

  if (options?.includePlanInflow) {
    const actualPlanIdsAll = new Set(
      (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : [])
        .filter(row => {
          const status = normalizeStatus(row?.status);
          return status !== "cancel" && status !== "done";
        })
        .map(row => Number(row?.plan_id || 0))
        .filter(Boolean)
    );
    const doneOrCancelPlanIdsAll = new Set(
      (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : [])
        .filter(row => {
          const status = normalizeStatus(row?.status);
          return status === "cancel" || status === "done";
        })
        .map(row => Number(row?.plan_id || 0))
        .filter(Boolean)
    );

    const inflowPlans = (Array.isArray(typeof currentPlansCache !== "undefined" ? currentPlansCache : window.currentPlansCache) ? (typeof currentPlansCache !== "undefined" ? currentPlansCache : window.currentPlansCache) : [])
      .filter(plan => Number(plan?.plan_hour ?? -1) < hour)
      .filter(plan => !slotPlanIds.has(Number(plan?.id || 0)))
      .filter(plan => !doneOrCancelPlanIdsAll.has(Number(plan?.id || 0)))
      .filter(plan => !actualPlanIdsAll.has(Number(plan?.id || 0)));

    inflowRows = __simulationNormalizePendingRows(inflowPlans, hour).map(row => ({
      ...row,
      simulated_inflow: true
    }));
  }

  return {
    rows: [...slotRows, ...inflowRows],
    summary: {
      slotPlanCount: slotCount,
      inflowPlanCount: inflowRows.length
    }
  };
}

(function registerSimulationFallbacks() {
  if (typeof window.buildProjectedRowsForHour !== 'function') {
    window.buildProjectedRowsForHour = __simulationBuildProjectedRowsForHourFallback;
  }
  if (typeof window.buildSimulationRowsForHour !== 'function') {
    window.buildSimulationRowsForHour = __simulationBuildSimulationRowsForHourFallback;
  }
  if (typeof window.getEffectiveVehiclesForHour !== 'function') {
    window.getEffectiveVehiclesForHour = function(targetHour) {
      const slotMinutes = Number(targetHour) * 60;
      const vehicles = Array.isArray(getSelectedVehiclesForToday?.()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
      const activeItems = (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : []).filter(row => {
        const status = normalizeStatus(row?.status);
        return status !== 'done' && status !== 'cancel';
      });
      return vehicles.filter(vehicle => {
        const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
        if (!rows.length) return true;
        const orderedRows = typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function'
          ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
          : rows;
        const forecast = typeof getVehicleRotationForecastSafe === 'function'
          ? getVehicleRotationForecastSafe(vehicle, orderedRows)
          : null;
        const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
        if (readyMinutes === null) return true;
        return readyMinutes <= slotMinutes;
      });
    };
  }
  if (typeof window.getVehicleDeadNamesForHour !== 'function') {
    window.getVehicleDeadNamesForHour = function(targetHour) {
      const slotMinutes = Number(targetHour) * 60;
      const vehicles = Array.isArray(getSelectedVehiclesForToday?.()) ? getSelectedVehiclesForToday().filter(Boolean) : [];
      const activeItems = (Array.isArray(typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) ? (typeof currentActualsCache !== "undefined" ? currentActualsCache : window.currentActualsCache) : []).filter(row => {
        const status = normalizeStatus(row?.status);
        return status !== 'done' && status !== 'cancel';
      });
      return vehicles.filter(vehicle => {
        const rows = activeItems.filter(item => Number(item?.vehicle_id || 0) === Number(vehicle.id));
        if (!rows.length) return false;
        const orderedRows = typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function'
          ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
          : rows;
        const forecast = typeof getVehicleRotationForecastSafe === 'function'
          ? getVehicleRotationForecastSafe(vehicle, orderedRows)
          : null;
        const readyMinutes = parseClockTextToMinutes(forecast?.predictedReadyTime);
        return readyMinutes !== null && readyMinutes > slotMinutes;
      }).map(vehicle => vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle.id}`);
    };
  }
})();

window.renderSimulationDispatchPreview = function renderSimulationDispatchPreview() {
  const hour = Number(els.simulationSlotSelect?.value ?? getSimulationSlotHourSafe() ?? getOperationBaseHour());
  setSimulationSlotHourSafe(hour);

  const built = __simulationBuildSimulationRowsForHourFallback(hour, { includePlanInflow: Boolean(els.simulationIncludePlanInflow?.checked) });
  const rows = Array.isArray(built?.rows) ? built.rows.filter(Boolean) : [];
  const vehicles = Array.isArray(getSelectedVehiclesForToday?.()) ? getSelectedVehiclesForToday().filter(Boolean) : [];

  if (!vehicles.length) {
    alert('可能車両を選択してください');
    return [];
  }
  if (!rows.length) {
    alert('この便の試算対象がありません');
    return [];
  }

  const monthlyMap = typeof buildMonthlyDistanceMapForCurrentMonth === 'function'
    ? buildMonthlyDistanceMapForCurrentMonth()
    : new Map();

  let assignments = [];
  if (typeof optimizeAssignments === 'function') {
    assignments = optimizeAssignments(rows, vehicles, monthlyMap, { mode: 'simulation_preview' });
  } else if (window.DispatchCore?.optimizeAssignments) {
    assignments = window.DispatchCore.optimizeAssignments(rows, vehicles, monthlyMap, { mode: 'simulation_preview' });
  }
  if (typeof resolveCapacityOverflowLocally === 'function') {
    assignments = resolveCapacityOverflowLocally(assignments, rows, vehicles, monthlyMap);
  }
  if (!Array.isArray(assignments)) assignments = [];

  const byItemId = new Map(assignments.map(a => [Number(a?.item_id || 0), Number(a?.vehicle_id || 0)]));
  const previewRows = rows.map(row => ({
    ...row,
    vehicle_id: byItemId.get(Number(row.id)) || 0
  }));
  const assignedCount = previewRows.filter(row => Number(row.vehicle_id || 0) > 0).length;
  const unassignedRows = previewRows.filter(row => Number(row.vehicle_id || 0) <= 0);

  setLastSimulationResultSafe({
    type: 'dispatch_preview',
    hour,
    assignments,
    rows: previewRows,
    summary: built.summary
  });

  if (els.simulationDiagnosis) {
    const diagStatus = unassignedRows.length ? 'warn' : 'ok';
    els.simulationDiagnosis.className = `hybrid-diagnosis ${diagStatus}`;
    els.simulationDiagnosis.innerHTML = `
      <div class="hybrid-state-row">
        ${buildStatusPill('便', `${getHourLabel(hour)}`, diagStatus)}
        ${buildStatusPill('対象', `${rows.length}名`, diagStatus)}
        ${buildStatusPill('割当', `${assignedCount}名`, diagStatus)}
        ${buildStatusPill('未割当', `${unassignedRows.length}名`, unassignedRows.length ? 'warn' : 'ok')}
      </div>
      <div class="hybrid-legend-grid">
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">対象便予定</div>
          <div class="hybrid-legend-value">${Number(built?.summary?.slotPlanCount || 0)}名</div>
        </div>
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">前便未処理予定</div>
          <div class="hybrid-legend-value">${Number(built?.summary?.inflowPlanCount || 0)}名</div>
        </div>
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">表示</div>
          <div class="hybrid-legend-value">試算のみ / 実績保存なし / 実際の配車結果はそのまま表示</div>
        </div>
      </div>
    `;
  }

  if (els.simulationPreview) {
    const formatMinutesJaSafe = minutes => {
      if (typeof formatMinutesAsJa === 'function') return formatMinutesAsJa(minutes);
      const safe = Math.max(0, Math.round(Number(minutes || 0)));
      const h = Math.floor(safe / 60);
      const m = safe % 60;
      if (h <= 0) return `${m}分`;
      if (m === 0) return `${h}時間`;
      return `${h}時間${m}分`;
    };

    const sortRowsForPreview = sourceRows => {
      const base = Array.isArray(sourceRows) ? sourceRows.slice() : [];
      if (typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function') {
        try {
          return moveManualLastItemsToEnd(sortItemsByNearestRoute(base));
        } catch (_) {}
      }
      return base.sort((a, b) => Number(a?.distance_km || 0) - Number(b?.distance_km || 0));
    };

    const buildVehicleSummary = assignedRows => {
      const orderedRows = sortRowsForPreview(assignedRows);
      const forecast = typeof calcVehicleRotationForecastGlobal === 'function'
        ? calcVehicleRotationForecastGlobal(null, orderedRows)
        : (typeof getVehicleRotationForecastSafe === 'function' ? getVehicleRotationForecastSafe(null, orderedRows) : null);
      const routeDistanceKm = Number(forecast?.routeDistanceKm ?? (typeof calculateRouteDistanceGlobal === 'function' ? calculateRouteDistanceGlobal(orderedRows) : 0) ?? 0);
      const lastRow = orderedRows[orderedRows.length - 1] || {};
      const returnDistanceKm = Number(forecast?.returnDistanceKm ?? lastRow?.distance_km ?? lastRow?.casts?.distance_km ?? 0);
      const roundTripKm = Number((routeDistanceKm + returnDistanceKm).toFixed(1));
      const totalMinutes = Number(forecast?.predictedReturnMinutes ?? (typeof getRowsTravelTimeSummary === 'function' ? getRowsTravelTimeSummary(orderedRows)?.totalMinutes : 0) ?? 0);
      return {
        orderedRows,
        roundTripKm,
        totalMinutes,
        totalMinutesLabel: formatMinutesJaSafe(totalMinutes)
      };
    };

    const assignedVehicleGroups = vehicles.map(vehicle => {
      const vehicleRows = previewRows.filter(row => Number(row?.vehicle_id || 0) === Number(vehicle?.id || 0));
      const summary = buildVehicleSummary(vehicleRows);
      return {
        vehicle,
        rows: summary.orderedRows,
        roundTripKm: summary.roundTripKm,
        totalMinutes: summary.totalMinutes,
        totalMinutesLabel: summary.totalMinutesLabel
      };
    }).filter(group => group.rows.length > 0)
      .sort((a, b) => Number(a.vehicle?.id || 0) - Number(b.vehicle?.id || 0));

    const buildCompactCastLabel = row => {
      const name = escapeHtml(row?.casts?.name || '-');
      const area = escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'));
      return `${name}(${area})${row?.simulated_inflow ? ' [流入]' : ''}`;
    };

    const assignedHtml = assignedVehicleGroups.length
      ? assignedVehicleGroups.map(group => {
          const vehicleLabel = group.vehicle?.driver_name || group.vehicle?.plate_number || `車両${group.vehicle?.id || '-'}`;
          const castLine = group.rows.map(buildCompactCastLabel).join(' ・ ');
          return `
            <div class="hybrid-legend-card" style="display:flex;flex-direction:column;gap:6px;padding:10px 12px;">
              <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
                <div class="hybrid-legend-title" style="margin:0;">${escapeHtml(vehicleLabel)}</div>
                <span class="chip">${group.rows.length}名</span>
                <span class="chip">往復 ${group.roundTripKm.toFixed(1)}km</span>
                <span class="chip">往復 ${escapeHtml(group.totalMinutesLabel)}</span>
              </div>
              <div style="font-size:13px;line-height:1.6;word-break:break-word;">${castLine || '-'}</div>
            </div>
          `;
        }).join('')
      : '<div class="muted">割当車両はありません</div>';

    const unassignedHtml = unassignedRows.length
      ? `
        <div class="hybrid-legend-card" style="display:flex;flex-direction:column;gap:6px;padding:10px 12px;">
          <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
            <div class="hybrid-legend-title" style="margin:0;">未割当</div>
            <span class="chip">${unassignedRows.length}名</span>
          </div>
          <div style="font-size:13px;line-height:1.6;word-break:break-word;">
            ${unassignedRows.map(buildCompactCastLabel).join(' ・ ')}
          </div>
        </div>
      ` : '';

    els.simulationPreview.className = 'simulation-preview';
    els.simulationPreview.innerHTML = `
      <div class="sim-preview-head">
        <h4 class="sim-preview-title">試算対象一覧</h4>
        <span class="chip">${getHourLabel(hour)}</span>
      </div>
      <div class="sim-preview-meta">対象 ${rows.length}名 / 割当 ${assignedCount}名 / 未割当 ${unassignedRows.length}名</div>
      <div style="display:flex;flex-direction:column;gap:8px;">${assignedHtml}${unassignedHtml}</div>
    `;
  }

  return previewRows;
};
