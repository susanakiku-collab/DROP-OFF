// renderVehicleTimeline
// dashboard.js から安全に分離した車両タイムライン描画系

function getVehicleRotationForecastSafe(vehicle, orderedRows) {
  try {
    if (typeof calcVehicleRotationForecastGlobal === "function") {
      return calcVehicleRotationForecastGlobal(vehicle, orderedRows);
    }
  } catch (e) {
    console.warn("calcVehicleRotationForecastGlobal fallback:", e);
  }
  try {
    if (typeof calcVehicleRotationForecast === "function") {
      return calcVehicleRotationForecast(vehicle, orderedRows);
    }
  } catch (e) {
    console.warn("calcVehicleRotationForecast fallback:", e);
  }

  const totalDistance = Array.isArray(orderedRows)
    ? orderedRows.reduce((sum, row) => sum + Number(row?.distance_km || 0), 0)
    : 0;

  return {
    routeDistanceKm: totalDistance,
    returnDistanceKm: 0,
    zoneLabel: "-",
    predictedDepartureTime: "-",
    predictedReturnTime: "-",
    predictedReadyTime: "-",
    predictedReturnMinutes: 0,
    extraSharedDelayMinutes: 0,
    returnAfterLabel: "0分後",
    stopCount: Array.isArray(orderedRows) ? orderedRows.length : 0,
    totalKm: totalDistance,
    dailyDistanceKm: totalDistance,
    jobCount: Array.isArray(orderedRows) ? orderedRows.length : 0,
    count: Array.isArray(orderedRows) ? orderedRows.length : 0
  };
}


function buildRotationTimelineHtmlSafe(vehicles, activeItems) {
  try {
    const timeline = (Array.isArray(vehicles) ? vehicles : [])
      .map(vehicle => {
        const rows = (Array.isArray(activeItems) ? activeItems : []).filter(
          item => Number(item.vehicle_id) === Number(vehicle.id)
        );

        if (!rows.length) return null;

        const orderedRows = (typeof moveManualLastItemsToEnd === "function" && typeof sortItemsByNearestRoute === "function")
          ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
          : rows;

        const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
        const summary = getVehicleDailySummary(vehicle, orderedRows);

        return {
          name: vehicle?.driver_name || vehicle?.plate_number || "-",
          lineId: vehicle?.line_id || "",
          returnAfterLabel: forecast?.returnAfterLabel || `${Number(forecast?.predictedReturnMinutes || 0)}分後`,
          nextRunTime: forecast?.predictedReadyTime || "-",
          totalKm: summary.totalKm,
          totalJobs: summary.jobCount
        };
      })
      .filter(Boolean);

    if (!timeline.length) return "";

    return `
      <div class="panel-card" style="margin-bottom:16px;">
        <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
        <div style="display:flex; flex-wrap:wrap; gap:8px;">
          ${timeline.map(item => `
            <div class="chip" style="padding:8px 12px;">
              <strong>${escapeHtml(item.name)}</strong>
              / 戻り ${escapeHtml(item.returnAfterLabel)}
              / 次便可能 ${escapeHtml(item.nextRunTime)}
              / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
              / ${Number(item.totalJobs || 0)}件
            </div>
          `).join("")}
        </div>
      </div>
    `;
  } catch (e) {
    console.error("buildRotationTimelineHtmlSafe error:", e);
    return "";
  }
}



function formatMinutesAsJa(totalMinutes) {
  const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
  const hours = Math.floor(safe / 60);
  const minutes = safe % 60;
  if (hours <= 0) return `${minutes}分`;
  if (minutes === 0) return `${hours}時間`;
  return `${hours}時間${minutes}分`;
}


function getVehiclePersistentDailyStats(vehicleId, orderedRows) {
  const numericVehicleId = Number(vehicleId || 0);
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  const reportDate = els.dispatchDate?.value || els.actualDate?.value || todayStr();

  const reportedRow = Array.isArray(currentDailyReportsCache)
    ? currentDailyReportsCache.find(
        row =>
          String(row.report_date || "") === String(reportDate || "") &&
          Number(row.vehicle_id || 0) === numericVehicleId
      )
    : null;

  const actualRows = Array.isArray(currentActualsCache)
    ? currentActualsCache.filter(
        item =>
          Number(item?.vehicle_id || 0) === numericVehicleId &&
          normalizeStatus(item?.status) !== "cancel"
      )
    : [];

  const baseRows = actualRows.length
    ? moveManualLastItemsToEnd(
        sortItemsByNearestRoute(
          [...actualRows].sort((a, b) => {
            const ah = Number(a?.actual_hour ?? a?.plan_hour ?? 0);
            const bh = Number(b?.actual_hour ?? b?.plan_hour ?? 0);
            if (ah !== bh) return ah - bh;
            return Number(a?.stop_order || 0) - Number(b?.stop_order || 0);
          })
        )
      )
    : rows;

  if (reportedRow && Number.isFinite(Number(reportedRow.distance_km))) {
    const reportedDistance = Number(Number(reportedRow.distance_km || 0).toFixed(1));
    const jobCount = actualRows.length || rows.length || 0;
    const driveMinutes = Math.round(
      getStoredTravelMinutes(baseRows[baseRows.length - 1]?.casts?.travel_minutes) ||
      estimateFallbackTravelMinutes(reportedDistance, getRepresentativeAreaFromRows(baseRows)) + jobCount
    );

    return {
      sendKm: reportedDistance,
      returnKm: 0,
      totalKm: reportedDistance,
      driveMinutes,
      jobCount,
      hasFixedReport: true
    };
  }

  if (!baseRows.length) {
    return {
      sendKm: 0,
      returnKm: 0,
      totalKm: 0,
      driveMinutes: 0,
      jobCount: 0,
      hasFixedReport: false
    };
  }

  const sendKm = Number(calculateRouteDistanceGlobal(baseRows) || 0);
  const lastRow = baseRows[baseRows.length - 1] || {};
  const returnKm = Number(lastRow.distance_km || 0);
  const totalKm = Number((sendKm + returnKm).toFixed(1));
  const driveMinutes = Math.round(getRowsTravelTimeSummary(baseRows).totalMinutes);

  return {
    sendKm: Number(sendKm.toFixed(1)),
    returnKm: Number(returnKm.toFixed(1)),
    totalKm,
    driveMinutes,
    jobCount: baseRows.length,
    hasFixedReport: false
  };
}


function getVehicleDailySummary(vehicle, orderedRows) {
  const vehicleId = Number(vehicle?.id || 0);
  const summary = getVehiclePersistentDailyStats(vehicleId, orderedRows);
  const isLastTripDriver = isDriverLastTripChecked(vehicleId);
  const displayTotalKm = (!summary.hasFixedReport && isLastTripDriver)
    ? Number(summary.sendKm || 0)
    : Number(summary.totalKm || 0);
  const sendOnlyMinutes = Math.round(getRowsTravelTimeSummary(orderedRows).sendOnlyMinutes);
  const displayDriveMinutes = (!summary.hasFixedReport && isLastTripDriver)
    ? sendOnlyMinutes
    : Math.round(Number(summary.driveMinutes || 0));

  return {
    sendKm: Number(summary.sendKm || 0),
    returnKm: Number(summary.returnKm || 0),
    totalKm: Number(displayTotalKm || 0),
    driveMinutes: Number(displayDriveMinutes || 0),
    jobCount: Number(summary.jobCount || 0),
    hasFixedReport: Boolean(summary.hasFixedReport)
  };
}


function getVehicleProjectedMonthlyDistance(vehicleId, monthlyMap, orderedRows) {
  const currentMonth = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0 };
  const todaySummary = getVehicleDailySummary({ id: vehicleId }, orderedRows);
  if (todaySummary.hasFixedReport) {
    return Number(Number(currentMonth.totalDistance || 0).toFixed(1));
  }
  return Number(Number(currentMonth.totalDistance || 0) + Number(todaySummary.totalKm || 0));
}
function buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap) {
  return vehicles.map(vehicle => {
    const rows = activeItems
      .filter(item => Number(item.vehicle_id) === Number(vehicle.id))
      .sort((a, b) => {
        const ah = Number(a.actual_hour ?? 0);
        const bh = Number(b.actual_hour ?? 0);
        if (ah !== bh) return ah - bh;

        const aa = normalizeAreaLabel(a.destination_area || "");
        const ba = normalizeAreaLabel(b.destination_area || "");
        if (aa !== ba) return aa.localeCompare(ba, "ja");

        return Number(a.stop_order || 0) - Number(b.stop_order || 0);
      });

    const orderedRows = moveManualLastItemsToEnd(sortItemsByNearestRoute(rows));
    return { vehicle, rows, orderedRows };
  });
}


function buildLineVehicleBlock(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) return "";

  const summary = getVehicleDailySummary(vehicle, rows);
  const forecast = getVehicleRotationForecastSafe(vehicle, rows);
  const driverName = getVehicleDisplayName(vehicle);
  const lineId = String(vehicle?.line_id || "").trim();
  const lastTripTag = isDriverLastTripChecked(vehicle?.id) ? " 【ラスト便】" : "";
  const routeUrl = buildVehicleRouteMapUrl(vehicle, rows);

  const header = [lineId, `🚗 ${driverName}${lastTripTag}`].filter(Boolean).join(" ");
  const body = rows.map((row, index) => {
    const castName = String(row?.casts?.name || "-").trim() || "-";
    const areaLabel = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "無し");
    return `${index + 1}️⃣ ${castName}　${areaLabel}`;
  });

  const footer = [
    `戻り ${forecast?.returnAfterLabel || "-"} / 次便可能 ${forecast?.predictedReadyTime || "-"}`,
    `距離 ${Number(summary?.totalKm || 0).toFixed(1)}km / 時間 ${formatMinutesAsJa(summary?.driveMinutes || 0)}`
  ];

  if (routeUrl) {
    footer.push("📍ルート");
    footer.push(routeUrl);
  }

  return [header, ...body, ...footer].join("\n");
}


function buildLineResultText() {
  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  const vehicles = getSelectedVehiclesForToday();
  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);
  return cards
    .map(({ vehicle, orderedRows }) => buildLineVehicleBlock(vehicle, orderedRows))
    .filter(Boolean)
    .join("\n\n")
    .trim();
}


