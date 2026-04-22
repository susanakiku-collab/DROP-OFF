// exportExcel
// dashboard.js から安全に分離した Excel / CSV 出力系

function buildMileageCsvSectionRows(title, days, drivers, options = {}) {
  const showOverallColumns = Boolean(options.showOverallColumns);
  const rows = [];

  rows.push([title]);

  const header = [
    "ドライバー",
    ...days.map(day => day.label),
    "合計距離",
    "合計出勤数",
    "1日平均距離"
  ];
  if (showOverallColumns) {
    header.push("全合計距離", "全出勤日数", "全平均距離");
  }
  rows.push(header);

  drivers.forEach(driver => {
    const sectionValues = days.map(day => Number(driver.byDate.get(day.key) || 0));
    const sectionDistance = Number(sectionValues.reduce((sum, value) => sum + value, 0).toFixed(1));
    const sectionWorkedDays = sectionValues.filter(value => value > 0).length;
    const sectionAvgDistance = sectionWorkedDays > 0 ? Number((sectionDistance / sectionWorkedDays).toFixed(1)) : 0;

    const row = [
      driver.driver_name || "-",
      ...sectionValues.map(value => value > 0 ? Number(value.toFixed(1)) : ""),
      sectionDistance,
      sectionWorkedDays,
      sectionAvgDistance
    ];

    if (showOverallColumns) {
      row.push(
        Number(Number(driver.total_distance_km || 0).toFixed(1)),
        Number(driver.worked_days || 0),
        Number(Number(driver.avg_distance_km || 0).toFixed(1))
      );
    }

    rows.push(row);
  });

  const dailyTotals = days.map(day =>
    Number(drivers.reduce((sum, driver) => sum + Number(driver.byDate.get(day.key) || 0), 0).toFixed(1))
  );
  const dailyWorkedCounts = days.map(day =>
    drivers.reduce((count, driver) => count + (Number(driver.byDate.get(day.key) || 0) > 0 ? 1 : 0), 0)
  );
  const dailyAverages = dailyTotals.map((total, index) =>
    dailyWorkedCounts[index] > 0 ? Number((total / dailyWorkedCounts[index]).toFixed(1)) : ""
  );

  const sectionTotalDistance = Number(dailyTotals.reduce((sum, value) => sum + Number(value || 0), 0).toFixed(1));
  const sectionWorkedDays = dailyWorkedCounts.reduce((sum, value) => sum + Number(value || 0), 0);
  const sectionAverageDistance = sectionWorkedDays > 0 ? Number((sectionTotalDistance / sectionWorkedDays).toFixed(1)) : 0;

  const overallTotalDistance = Number(drivers.reduce((sum, driver) => sum + Number(driver.total_distance_km || 0), 0).toFixed(1));
  const overallWorkedDays = drivers.reduce((sum, driver) => sum + Number(driver.worked_days || 0), 0);
  const overallAvgDistance = overallWorkedDays > 0 ? Number((overallTotalDistance / overallWorkedDays).toFixed(1)) : 0;

  const totalRow = [
    "日別合計",
    ...dailyTotals.map(v => v || ""),
    sectionTotalDistance,
    sectionWorkedDays,
    sectionAverageDistance
  ];
  if (showOverallColumns) totalRow.push(overallTotalDistance, overallWorkedDays, overallAvgDistance);
  rows.push(totalRow);

  const avgRow = [
    "日別平均",
    ...dailyAverages,
    drivers.length ? Number((sectionTotalDistance / drivers.length).toFixed(1)) : 0,
    "",
    sectionAverageDistance
  ];
  if (showOverallColumns) {
    avgRow.push(
      drivers.length ? Number((overallTotalDistance / drivers.length).toFixed(1)) : 0,
      drivers.length ? Number((overallWorkedDays / drivers.length).toFixed(1)) : 0,
      overallAvgDistance
    );
  }
  rows.push(avgRow);

  return rows;
}


function buildMileageMatrixCsvRows(rows, startDate, endDate) {
  const calendar = buildMileageCalendarRows(rows, startDate, endDate);
  const result = [];
  result.push([calendar.monthLabel, `${startDate.replaceAll("-", "/")} ～ ${endDate.replaceAll("-", "/")}`]);
  result.push([]);
  result.push(...buildMileageCsvSectionRows("1日～16日", calendar.firstHalfDays, calendar.drivers));
  result.push([]);
  result.push(...buildMileageCsvSectionRows("17日～月末", calendar.secondHalfDays, calendar.drivers, {
    showOverallColumns: true
  }));
  return result;
}


async function exportDriverMileageReportCsv() {
  if (!currentMileageExportRows.length) {
    await previewDriverMileageReport();
    if (!currentMileageExportRows.length) return;
  }

  const startDate = els.mileageReportStartDate?.value || todayStr();
  const endDate = els.mileageReportEndDate?.value || todayStr();

  const aoa = buildMileageMatrixCsvRows(currentMileageExportRows, startDate, endDate);
  const csv = aoa.map(row => row.map(value => csvEscape(value ?? "")).join(",")).join("\n");
  downloadTextFile(`driver_mileage_matrix_${startDate}_${endDate}.csv`, csv, "text/csv;charset=utf-8");
}


async function exportDriverMileageReportXlsx() {
  return exportDriverMileageReportCsv();
}



function exportVehiclesCsv() {
  const headers = [
    "plate_number",
    "vehicle_area",
    "home_area",
    "seat_capacity",
    "driver_name",
    "line_id",
    "status",
    "memo"
  ];

  const rows = allVehiclesCache.map(vehicle => [
    vehicle.plate_number || "",
    normalizeAreaLabel(vehicle.vehicle_area || ""),
    normalizeAreaLabel(vehicle.home_area || ""),
    vehicle.seat_capacity ?? "",
    vehicle.driver_name || "",
    vehicle.line_id || "",
    vehicle.status || "waiting",
    vehicle.memo || ""
  ]);

  const csv = [headers.join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`vehicles_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}



function exportPlansCsv() {
  const planDate = els.planDate?.value || todayStr();
  const rows = [...currentPlansCache]
    .sort((a, b) => Number(a.plan_hour || 0) - Number(b.plan_hour || 0) || Number(a.id || 0) - Number(b.id || 0))
    .map(plan => ({
      plan_date: planDate,
      plan_hour: Number(plan.plan_hour || 0),
      cast_id: Number(plan.cast_id || plan.casts?.id || 0),
      cast_name: plan.casts?.name || "",
      destination_address: plan.destination_address || plan.casts?.address || "",
      planned_area: normalizeAreaLabel(plan.planned_area || plan.casts?.area || ""),
      distance_km: plan.distance_km ?? "",
      note: plan.note || "",
      status: plan.status || "planned",
      vehicle_group: plan.vehicle_group || ""
    }));

  const headers = [
    "plan_date",
    "plan_hour",
    "cast_id",
    "cast_name",
    "destination_address",
    "planned_area",
    "distance_km",
    "note",
    "status",
    "vehicle_group"
  ];

  const csv = [headers.join(","), ...rows.map(row => headers.map(key => csvEscape(row[key] ?? "")).join(","))].join("\n");
  downloadTextFile(`plans_${planDate}.csv`, csv, "text/csv;charset=utf-8");
}

