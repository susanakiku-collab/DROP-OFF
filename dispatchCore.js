(function (global) {
  'use strict';

  const VERSION = 'dispatchcore-2026-04-14-normal-pick-debug-fix8';
  const AREA_MEMBER_ANGLE_THRESHOLD = 30;
  const SAME_DIRECTION_MERGE_ANGLE_THRESHOLD = 18;
  const SAME_DIRECTION_MERGE_POINT_DISTANCE_KM = 18;
  const LONG_ROUTE_BUNDLE_MINUTES = 55;
  const LAST_TRIP_MAX_DIRECTION_DIFF = 90;
  const DISPATCH_DEBUG = global.DISPATCH_DEBUG === true;

  function debugLog(...args) {
    if (!DISPATCH_DEBUG) return;
    console.log(...args);
  }

  function debugWarn(...args) {
    if (!DISPATCH_DEBUG) return;
    console.warn(...args);
  }

  function toNumber(value, fallback = 0) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
  }

  function normalizeId(value) {
    if (value === null || value === undefined) return '';
    const text = String(value).trim();
    return text;
  }

  function normalizeAngle(angle) {
    const normalized = angle % 360;
    return normalized < 0 ? normalized + 360 : normalized;
  }

  function angleDiff(a, b) {
    const diff = Math.abs(normalizeAngle(a) - normalizeAngle(b));
    return diff > 180 ? 360 - diff : diff;
  }

  function safeNormalizeArea(area) {
    if (typeof global.normalizeAreaLabel === 'function') {
      return String(global.normalizeAreaLabel(area || '') || '').trim();
    }
    return String(area || '').trim();
  }

  function getCanonicalAreaSafe(area) {
    return typeof global.getCanonicalArea === 'function'
      ? String(global.getCanonicalArea(area || '') || '').trim()
      : '';
  }

  function getAreaDisplayGroupSafe(area) {
    return typeof global.getAreaDisplayGroup === 'function'
      ? String(global.getAreaDisplayGroup(area || '') || '').trim()
      : '';
  }

  function getAreaAffinityScoreSafe(areaA, areaB) {
    if (typeof global.getAreaAffinityScore === 'function') {
      return toNumber(global.getAreaAffinityScore(areaA || '', areaB || ''), 0);
    }
    return 0;
  }

  function getDirectionAffinityScoreSafe(areaA, areaB) {
    if (typeof global.getDirectionAffinityScore === 'function') {
      return toNumber(global.getDirectionAffinityScore(areaA || '', areaB || ''), 0);
    }
    return 0;
  }

  function getStoredTravelMinutesSafe(value) {
    if (typeof global.getStoredTravelMinutes === 'function') {
      return Math.max(0, Math.round(toNumber(global.getStoredTravelMinutes(value), 0)));
    }
    return Math.max(0, Math.round(toNumber(value, 0)));
  }

  function estimateFallbackTravelMinutesSafe(distanceKm, area) {
    if (typeof global.estimateFallbackTravelMinutes === 'function') {
      return Math.max(0, Math.round(toNumber(global.estimateFallbackTravelMinutes(distanceKm, area), 0)));
    }
    const safeKm = Math.max(0, toNumber(distanceKm, 0));
    const fallback = safeKm <= 0 ? 0 : Math.round((safeKm / 30) * 60);
    return Math.max(0, fallback);
  }

  function haversineKm(a, b) {
    const lat1 = toNumber(a?.lat, NaN);
    const lng1 = toNumber(a?.lng, NaN);
    const lat2 = toNumber(b?.lat, NaN);
    const lng2 = toNumber(b?.lng, NaN);
    if (![lat1, lng1, lat2, lng2].every(Number.isFinite)) return Infinity;

    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const rad1 = lat1 * Math.PI / 180;
    const rad2 = lat2 * Math.PI / 180;
    const sinLat = Math.sin(dLat / 2);
    const sinLng = Math.sin(dLng / 2);
    const h = sinLat * sinLat + Math.cos(rad1) * Math.cos(rad2) * sinLng * sinLng;
    return 2 * R * Math.asin(Math.sqrt(h));
  }

  function calcAngleFromOrigin(origin, point) {
    const dx = point.lng - origin.lng;
    const dy = point.lat - origin.lat;
    return normalizeAngle(Math.atan2(dy, dx) * 180 / Math.PI);
  }

  function getOrigin() {
    const cfg = global.APP_CONFIG || {};
    const lat = toNumber(cfg.ORIGIN_LAT, NaN);
    const lng = toNumber(cfg.ORIGIN_LNG, NaN);
    return { lat, lng, name: cfg.ORIGIN_LABEL || '起点' };
  }

  function getItemPoint(item) {
    const lat = toNumber(item?.casts?.latitude ?? item?.latitude ?? item?.lat, NaN);
    const lng = toNumber(item?.casts?.longitude ?? item?.longitude ?? item?.lng, NaN);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { lat, lng };
  }

  function getItemDistanceKm(item, origin, point) {
    const stored = toNumber(item?.distance_km ?? item?.casts?.distance_km, NaN);
    if (Number.isFinite(stored) && stored > 0) return stored;
    if (!point) return Infinity;
    return haversineKm(origin, point);
  }

  function getItemArea(item) {
    return safeNormalizeArea(
      item?.destination_area || item?.planned_area || item?.cluster_area || item?.casts?.area || ''
    );
  }

  function getItemTravelMinutes(item, area, distanceKm) {
    const stored = getStoredTravelMinutesSafe(item?.travel_minutes ?? item?.casts?.travel_minutes);
    if (stored > 0) return stored;
    return estimateFallbackTravelMinutesSafe(distanceKm, area);
  }

  function normalizeItem(item, origin) {
    const point = getItemPoint(item);
    const distanceKm = getItemDistanceKm(item, origin, point);
    const angleFromOrigin = point ? calcAngleFromOrigin(origin, point) : null;
    const area = getItemArea(item);
    const itemId = normalizeId(item?.id);
    const castId = normalizeId(item?.cast_id);
    const travelMinutes = getItemTravelMinutes(item, area, distanceKm);

    return {
      itemId,
      castId,
      actualHour: Math.trunc(toNumber(item?.actual_hour ?? item?.plan_hour ?? 0, 0)),
      name: String(item?.casts?.name || item?.name || `item_${itemId}` || '').trim() || '-',
      area,
      distanceKm,
      angleFromOrigin,
      point,
      travelMinutes,
      raw: item
    };
  }

  function normalizeVehicle(vehicle, monthlyMap) {
    const vehicleId = toNumber(vehicle?.id, 0);
    const monthly = resolveMonthlyStatsEntry(monthlyMap, vehicle) || {};
    const directAvgDistance = toNumber(
      vehicle?.avgDistance ??
      vehicle?.avg_distance ??
      vehicle?.averageDistance ??
      vehicle?.daily_avg_km ??
      vehicle?.avg_km_per_day,
      0
    );
    const directTotalDistance = toNumber(
      vehicle?.totalDistance ??
      vehicle?.total_distance ??
      vehicle?.monthlyDistance ??
      vehicle?.monthly_distance ??
      vehicle?.monthly_distance_km,
      0
    );
    const directWorkedDays = Math.max(
      0,
      Math.trunc(
        toNumber(
          vehicle?.workedDays ??
          vehicle?.worked_days ??
          vehicle?.workDays ??
          vehicle?.work_days,
          0
        )
      )
    );
    const avgDistance = toNumber(
      monthly?.avgDistance ??
      monthly?.avg_distance ??
      monthly?.averageDistance ??
      directAvgDistance,
      0
    );
    const totalDistance = toNumber(
      monthly?.totalDistance ??
      monthly?.total_distance ??
      directTotalDistance,
      0
    );
    const workedDays = Math.max(
      0,
      Math.trunc(
        toNumber(
          monthly?.workedDays ??
          monthly?.worked_days ??
          directWorkedDays,
          0
        )
      )
    );
    const homeLat = toNumber(vehicle?.home_lat, NaN);
    const homeLng = toNumber(vehicle?.home_lng, NaN);
    const isLastTrip = typeof global.isDriverLastTripChecked === 'function'
      ? Boolean(global.isDriverLastTripChecked(vehicleId))
      : Boolean(vehicle?.is_last_trip || vehicle?.isLastTrip);
    return {
      vehicleId,
      driverName: String(vehicle?.driver_name || vehicle?.name || vehicle?.plate_number || '').trim() || '-',
      capacity: Math.max(0, Math.trunc(toNumber(vehicle?.seat_capacity ?? vehicle?.capacity, 0))),
      avgDistance,
      totalDistance,
      workedDays,
      todayDistance: Math.max(0, toNumber(vehicle?.todayDistance ?? vehicle?.today_distance, 0)),
      score: avgDistance,
      homeArea: safeNormalizeArea(vehicle?.home_area || ''),
      homePoint: Number.isFinite(homeLat) && Number.isFinite(homeLng) ? { lat: homeLat, lng: homeLng } : null,
      isLastTrip,
      raw: vehicle
    };
  }

  function byDistanceDesc(a, b) {
    if (b.distanceKm !== a.distanceKm) return b.distanceKm - a.distanceKm;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  }

  function byDistanceAsc(a, b) {
    if (a.distanceKm !== b.distanceKm) return a.distanceKm - b.distanceKm;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  }

  function buildAxis(anchor, index, vehicle) {
    return {
      axisId: `axis_${index + 1}`,
      vehicleId: vehicle.vehicleId,
      driverName: vehicle.driverName,
      capacity: vehicle.capacity,
      anchorItemId: anchor.itemId,
      anchorName: anchor.name,
      anchorAngle: anchor.angleFromOrigin,
      anchorDistanceKm: anchor.distanceKm,
      members: [anchor],
      pending: [],
      overflowed: []
    };
  }

  function getPrimaryCandidateAxes(item, axes, threshold) {
    return axes
      .map(axis => ({ axis, diff: angleDiff(item.angleFromOrigin, axis.anchorAngle) }))
      .filter(entry => Number.isFinite(entry.diff) && entry.diff <= threshold)
      .sort((a, b) => {
        if (a.diff !== b.diff) return a.diff - b.diff;
        if (b.axis.anchorDistanceKm !== a.axis.anchorDistanceKm) return b.axis.anchorDistanceKm - a.axis.anchorDistanceKm;
        return String(a.axis.axisId).localeCompare(String(b.axis.axisId), 'ja');
      });
  }

  function assignStagewiseCoveredIds(people, axes, threshold) {
    const coveredIds = new Set(axes.map(axis => axis.anchorItemId));
    people.forEach(item => {
      if (coveredIds.has(item.itemId)) return;
      const matches = getPrimaryCandidateAxes(item, axes, threshold);
      if (!matches.length) return;
      coveredIds.add(item.itemId);
    });
    return coveredIds;
  }

  function selectFallbackAxisCandidate(sorted, selectedAnchorIds, axes) {
    const remaining = sorted.filter(item => !selectedAnchorIds.has(item.itemId));
    if (!remaining.length) return null;
    if (!axes.length) return remaining[0] || null;

    return remaining
      .map(item => {
        const minAngleDiff = axes.reduce((min, axis) => {
          const diff = angleDiff(item.angleFromOrigin, axis.anchorAngle);
          return Math.min(min, diff);
        }, Infinity);
        return { item, minAngleDiff };
      })
      .sort((a, b) => {
        if (b.minAngleDiff !== a.minAngleDiff) return b.minAngleDiff - a.minAngleDiff;
        if (b.item.distanceKm !== a.item.distanceKm) return b.item.distanceKm - a.item.distanceKm;
        return String(a.item.name || '').localeCompare(String(b.item.name || ''), 'ja');
      })[0]?.item || null;
  }

  function selectFixedAxes(people, vehicles, threshold) {
    const sorted = [...people].sort(byDistanceDesc);
    const axes = [];
    if (!sorted.length || !vehicles.length) return axes;

    const selectedAnchorIds = new Set();
    let coveredIds = new Set();

    for (let i = 0; i < vehicles.length; i += 1) {
      let candidate = sorted.find(item => !coveredIds.has(item.itemId) && !selectedAnchorIds.has(item.itemId));

      if (!candidate) {
        candidate = selectFallbackAxisCandidate(sorted, selectedAnchorIds, axes);
      }

      if (!candidate) break;

      const axis = buildAxis(candidate, i, vehicles[i]);
      axes.push(axis);
      selectedAnchorIds.add(candidate.itemId);
      coveredIds = assignStagewiseCoveredIds(sorted, axes, threshold);
    }
    return axes;
  }

  function pushPendingMembers(people, axes, threshold) {
    const anchorIds = new Set(axes.map(axis => axis.anchorItemId));
    const unresolved = [];

    people.forEach(item => {
      if (anchorIds.has(item.itemId)) return;
      const matches = getPrimaryCandidateAxes(item, axes, threshold);
      if (!matches.length) {
        unresolved.push(item);
        return;
      }
      if (matches.length === 1) {
        matches[0].axis.pending.push(item);
        return;
      }
      unresolved.push(item);
    });

    axes.forEach(axis => {
      axis.pending.sort(byDistanceDesc);
      const freeSeats = Math.max(0, axis.capacity - axis.members.length);
      axis.members.push(...axis.pending.slice(0, freeSeats));
      axis.overflowed.push(...axis.pending.slice(freeSeats));
      axis.pending = [];
    });

    return unresolved.concat(axes.flatMap(axis => axis.overflowed));
  }

  function getAllAxisChoices(item, axes) {
    return axes
      .map(axis => ({ axis, diff: angleDiff(item.angleFromOrigin, axis.anchorAngle) }))
      .sort((a, b) => {
        if (a.diff !== b.diff) return a.diff - b.diff;
        if (b.axis.anchorDistanceKm !== a.axis.anchorDistanceKm) return b.axis.anchorDistanceKm - a.axis.anchorDistanceKm;
        return String(a.axis.axisId).localeCompare(String(b.axis.axisId), 'ja');
      });
  }

  function resolveUnsettledMembers(unresolved, axes) {
    const overflow = [];
    const sorted = [...unresolved].sort(byDistanceDesc);

    sorted.forEach(item => {
      const choices = getAllAxisChoices(item, axes);
      const target = choices.find(choice => choice.axis.members.length < choice.axis.capacity);
      if (!target) {
        overflow.push(item);
        return;
      }
      target.axis.members.push(item);
    });

    return overflow;
  }

  function finalizeAxisOrder(axes) {
    axes.forEach(axis => {
      axis.members.sort(byDistanceAsc);
    });
  }

  function buildAssignments(axes) {
    const assignments = [];
    axes.forEach(axis => {
      axis.members.forEach((item, index) => {
        assignments.push({
          item_id: item.itemId,
          actual_hour: item.actualHour,
          vehicle_id: axis.vehicleId,
          driver_name: axis.driverName,
          distance_km: item.distanceKm,
          stop_order: index + 1
        });
      });
    });
    return assignments;
  }

  function buildOverflowMeta(overflow) {
    return {
      overflowGroups: [],
      overflowEvaluations: [],
      capacityOverflowCount: overflow.length,
      capacityOverflowItems: overflow.map(item => ({
        itemId: item.itemId,
        hour: item.actualHour,
        group: item.area || '無し',
        distanceKm: item.distanceKm
      }))
    };
  }

  function getAxisRepresentative(axis, origin) {
    const members = Array.isArray(axis?.members) ? axis.members.filter(Boolean) : [];
    const sorted = [...members].sort(byDistanceDesc);
    const farthest = sorted[0] || null;
    return {
      axis,
      farthest,
      point: farthest?.point || null,
      distanceKm: toNumber(farthest?.distanceKm, 0),
      angleFromOrigin: Number.isFinite(farthest?.angleFromOrigin) ? farthest.angleFromOrigin : null,
      size: members.length
    };
  }

  function getMaxVehicleCapacity(vehicles) {
    return Math.max(0, ...(Array.isArray(vehicles) ? vehicles : []).map(vehicle => Number(vehicle?.capacity || 0)));
  }

  function buildAxisMergeCandidate(leftRep, rightRep, maxCapacity) {
    if (!leftRep?.axis || !rightRep?.axis) return null;
    if (!Number.isFinite(leftRep.angleFromOrigin) || !Number.isFinite(rightRep.angleFromOrigin)) return null;
    const combinedSize = Number(leftRep.size || 0) + Number(rightRep.size || 0);
    if (combinedSize <= 0 || combinedSize > maxCapacity) return null;

    const angleGap = angleDiff(leftRep.angleFromOrigin, rightRep.angleFromOrigin);
    if (!Number.isFinite(angleGap) || angleGap > SAME_DIRECTION_MERGE_ANGLE_THRESHOLD) return null;

    const pointDistanceKm = (leftRep.point && rightRep.point)
      ? haversineKm(leftRep.point, rightRep.point)
      : Infinity;
    if (Number.isFinite(pointDistanceKm) && pointDistanceKm > SAME_DIRECTION_MERGE_POINT_DISTANCE_KM) return null;

    const primary = Number(leftRep.distanceKm || 0) >= Number(rightRep.distanceKm || 0) ? leftRep : rightRep;
    const secondary = primary === leftRep ? rightRep : leftRep;

    return {
      primaryAxisId: primary.axis.axisId,
      secondaryAxisId: secondary.axis.axisId,
      angleGap,
      pointDistanceKm,
      combinedSize,
      primaryDistanceKm: Number(primary.distanceKm || 0),
      secondaryDistanceKm: Number(secondary.distanceKm || 0)
    };
  }

  function pickSameDirectionMergePair(axes, vehicles, origin) {
    if (!Array.isArray(axes) || axes.length < 2) return null;
    const maxCapacity = getMaxVehicleCapacity(vehicles);
    if (maxCapacity <= 0) return null;

    const reps = axes.map(axis => getAxisRepresentative(axis, origin));
    const candidates = [];

    for (let i = 0; i < reps.length; i += 1) {
      for (let j = i + 1; j < reps.length; j += 1) {
        const candidate = buildAxisMergeCandidate(reps[i], reps[j], maxCapacity);
        if (candidate) candidates.push(candidate);
      }
    }

    candidates.sort((a, b) => {
      if (a.angleGap !== b.angleGap) return a.angleGap - b.angleGap;
      if (a.pointDistanceKm !== b.pointDistanceKm) return a.pointDistanceKm - b.pointDistanceKm;
      if (b.primaryDistanceKm !== a.primaryDistanceKm) return b.primaryDistanceKm - a.primaryDistanceKm;
      if (a.combinedSize !== b.combinedSize) return a.combinedSize - b.combinedSize;
      return String(a.primaryAxisId).localeCompare(String(b.primaryAxisId), 'ja');
    });

    return candidates[0] || null;
  }

  function mergeFixedAxesByCandidate(axes, candidate) {
    if (!candidate?.primaryAxisId || !candidate?.secondaryAxisId) return Array.isArray(axes) ? [...axes] : [];
    const primary = axes.find(axis => axis.axisId === candidate.primaryAxisId);
    const secondary = axes.find(axis => axis.axisId === candidate.secondaryAxisId);
    if (!primary || !secondary) return [...axes];

    primary.members = [...(Array.isArray(primary.members) ? primary.members : []), ...(Array.isArray(secondary.members) ? secondary.members : [])];
    primary.overflowed = [...(Array.isArray(primary.overflowed) ? primary.overflowed : []), ...(Array.isArray(secondary.overflowed) ? secondary.overflowed : [])];
    primary.pending = [];
    return axes.filter(axis => axis.axisId !== candidate.secondaryAxisId);
  }

  function collapseSameDirectionAxes(axes, vehicles, origin) {
    let current = Array.isArray(axes) ? [...axes] : [];
    const standbyCount = Math.max(0, (Array.isArray(vehicles) ? vehicles.length : 0) - current.length);
    if (current.length < 2) return { axes: current, mergeLogs: [], standbyCount };

    const mergeLogs = [];
    while (current.length >= 2) {
      const candidate = pickSameDirectionMergePair(current, vehicles, origin);
      if (!candidate) break;
      mergeLogs.push(candidate);
      current = mergeFixedAxesByCandidate(current, candidate);
    }

    current.forEach(axis => {
      axis.members.sort(byDistanceAsc);
      const rep = getAxisRepresentative(axis, origin);
      if (rep?.farthest) {
        axis.anchorItemId = rep.farthest.itemId;
        axis.anchorName = rep.farthest.name;
        axis.anchorAngle = rep.farthest.angleFromOrigin;
        axis.anchorDistanceKm = rep.farthest.distanceKm;
      }
    });

    return {
      axes: current,
      mergeLogs,
      standbyCount: Math.max(0, (Array.isArray(vehicles) ? vehicles.length : 0) - current.length)
    };
  }

  function vehicleCanTakeAxis(vehicle, axisRep) {
    if (!vehicle || !axisRep) return false;
    if (Number(vehicle.capacity || 0) <= 0) return false;
    const need = Math.max(1, Number(axisRep.size || 0));
    return Number(vehicle.capacity || 0) >= need;
  }

  function estimateAxisReferenceDistanceKm(axisRep, vehicle) {
    const baseDistanceKm = Math.max(0, toNumber(axisRep?.distanceKm ?? axisRep?.anchorDistanceKm, 0));
    if (!(baseDistanceKm > 0)) return 0;
    const multiplier = vehicle?.isLastTrip ? 1 : 2;
    return Number((baseDistanceKm * multiplier).toFixed(1));
  }

  function estimateRowsReferenceDistanceKm(rows, itemMap, vehicle) {
    const items = (Array.isArray(rows) ? rows : [])
      .map(row => itemMap.get(normalizeId(row?.item_id)))
      .filter(Boolean);
    const farthest = [...items].sort(byDistanceDesc)[0] || null;
    return estimateAxisReferenceDistanceKm({ distanceKm: farthest?.distanceKm || 0 }, vehicle);
  }


  function resolveMonthlyStatsEntry(monthlyMap, vehicleOrId, fallbackVehicle = null) {
    if (!(monthlyMap instanceof Map) || monthlyMap.size <= 0) return {};
    const vehicle = (vehicleOrId && typeof vehicleOrId === 'object') ? vehicleOrId : fallbackVehicle;
    const rawId = (vehicleOrId && typeof vehicleOrId !== 'object')
      ? vehicleOrId
      : (vehicle?.id ?? vehicle?.vehicleId ?? vehicle?.vehicle_id ?? vehicle?.cloud_id ?? vehicle?.uuid ?? vehicle?.raw?.id);

    const candidateKeys = [];
    const pushKey = value => {
      if (value == null) return;
      const raw = String(value).trim();
      if (!raw || raw === 'undefined' || raw === 'null') return;
      if (!candidateKeys.includes(value)) candidateKeys.push(value);
      if (!candidateKeys.includes(raw)) candidateKeys.push(raw);
      const num = Number(raw);
      if (Number.isFinite(num) && !candidateKeys.includes(num)) candidateKeys.push(num);
    };

    pushKey(rawId);
    pushKey(vehicle?.id);
    pushKey(vehicle?.vehicleId);
    pushKey(vehicle?.vehicle_id);
    pushKey(vehicle?.cloud_id);
    pushKey(vehicle?.uuid);
    pushKey(vehicle?.raw?.id);
    pushKey(vehicle?.driver_name);
    pushKey(vehicle?.driverName);
    pushKey(vehicle?.plate_number);
    pushKey(vehicle?.raw?.driver_name);
    pushKey(vehicle?.raw?.plate_number);

    for (const key of candidateKeys) {
      if (monthlyMap.has(key)) return monthlyMap.get(key) || {};
    }

    for (const [mapKey, entry] of monthlyMap.entries()) {
      const entryCandidates = [
        mapKey,
        entry?.vehicleId,
        entry?.vehicle_id,
        entry?.id,
        entry?.cloud_id,
        entry?.uuid,
        entry?.driverName,
        entry?.driver_name,
        entry?.plate_number
      ];
      const normalizedEntryKeys = new Set();
      entryCandidates.forEach(value => {
        if (value == null) return;
        const raw = String(value).trim();
        if (!raw || raw === 'undefined' || raw === 'null') return;
        normalizedEntryKeys.add(raw);
        const num = Number(raw);
        if (Number.isFinite(num)) normalizedEntryKeys.add(String(num));
      });
      const hit = candidateKeys.some(key => normalizedEntryKeys.has(String(key).trim()));
      if (hit) return entry || {};
    }

    return {};
  }

  function getMonthlyStatsForVehicle(monthlyMap, vehicleId, fallbackVehicle) {
    const monthly = resolveMonthlyStatsEntry(monthlyMap, vehicleId, fallbackVehicle) || {};
    return {
      totalDistance: Math.max(0, toNumber(monthly?.totalDistance ?? monthly?.total_distance ?? fallbackVehicle?.totalDistance, 0)),
      workedDays: Math.max(0, Math.trunc(toNumber(monthly?.workedDays ?? monthly?.worked_days ?? fallbackVehicle?.workedDays, 0))),
      avgDistance: Math.max(0, toNumber(monthly?.avgDistance ?? monthly?.avg_distance ?? monthly?.averageDistance ?? fallbackVehicle?.avgDistance, 0))
    };
  }

  function calcProjectedAvgDistance(baseTotalDistance, baseWorkedDays, currentAssignedDistanceKm, hasWorkAlready, extraDistanceKm) {
    const monthlyDistance = Math.max(0, toNumber(baseTotalDistance, 0));
    const currentDistance = Math.max(0, toNumber(currentAssignedDistanceKm, 0));
    const extraDistance = Math.max(0, toNumber(extraDistanceKm, 0));
    const workedDays = Math.max(0, Math.trunc(toNumber(baseWorkedDays, 0)));
    const projectedWorkedDays = workedDays + ((hasWorkAlready || extraDistance <= 0) ? 0 : 1);
    const projectedTotalDistance = monthlyDistance + currentDistance + extraDistance;
    if (projectedWorkedDays <= 0) return projectedTotalDistance > 0 ? projectedTotalDistance : 0;
    return projectedTotalDistance / projectedWorkedDays;
  }

  function getProjectedAxisAverageDistance(vehicle, axisRep) {
    const extraDistanceKm = estimateAxisReferenceDistanceKm(axisRep, vehicle);
    return calcProjectedAvgDistance(vehicle?.totalDistance, vehicle?.workedDays, 0, false, extraDistanceKm);
  }


  function buildNormalVehicleDebugRow(vehicle, axisRep, extra = {}) {
    const projectedAvg = getProjectedAxisAverageDistance(vehicle, axisRep);
    const avgDistance = Math.max(0, toNumber(vehicle?.avgDistance, 0));
    const totalDistance = Math.max(0, toNumber(vehicle?.totalDistance, 0));
    const workedDays = Math.max(0, Math.trunc(toNumber(vehicle?.workedDays, 0)));
    const todayDistance = Math.max(0, toNumber(vehicle?.todayDistance, 0));
    const unused = todayDistance <= 0 && (workedDays <= 0 || avgDistance <= 0);
    return {
      vehicleId: vehicle?.vehicleId,
      driverName: vehicle?.driverName,
      unused,
      avgDistance,
      projectedAvg,
      todayDistance,
      totalDistance,
      workedDays,
      capacity: vehicle?.capacity,
      axisId: axisRep?.axis?.axisId || axisRep?.axisId || '',
      axisAnchor: axisRep?.axis?.anchorName || axisRep?.anchorName || '',
      axisDistanceKm: Number(axisRep?.distanceKm || axisRep?.axis?.anchorDistanceKm || 0),
      axisSize: Number(axisRep?.size || axisRep?.axis?.members?.length || 0),
      canTake: typeof extra.canTake === 'boolean' ? extra.canTake : true,
      reason: extra.reason || ''
    };
  }

  function logNormalVehicleSelection(label, axisRep, freeVehicles, candidates, picked) {
    if (!DISPATCH_DEBUG) return;
    try {
      const rows = (Array.isArray(freeVehicles) ? freeVehicles : []).map(vehicle => {
        const canTake = (Array.isArray(candidates) ? candidates : []).some(candidate => candidate?.vehicleId === vehicle?.vehicleId);
        const reason = canTake ? 'candidate' : 'filtered_out_by_vehicleCanTakeAxis';
        return buildNormalVehicleDebugRow(vehicle, axisRep, { canTake, reason });
      });
      console.groupCollapsed(`[DispatchCore][NORMAL_PICK][${label}] axis=${axisRep?.axis?.axisId || axisRep?.axisId || '-'} anchor=${axisRep?.axis?.anchorName || axisRep?.anchorName || '-'} picked=${picked?.driverName || '-'} (${picked?.vehicleId || '-'})`);
      console.table(rows);
      if (picked) {
        console.log('[DispatchCore][NORMAL_PICK][PICKED]', buildNormalVehicleDebugRow(picked, axisRep, { canTake: true, reason: 'picked' }));
      } else {
        console.warn('[DispatchCore][NORMAL_PICK] no vehicle picked');
      }
      console.groupEnd();
    } catch (error) {
      console.warn('[DispatchCore][NORMAL_PICK] log failed:', error);
    }
  }

  function logComponentVehicleSelection(componentRows, states, best, itemMap, monthlyMap, origin) {
    if (!DISPATCH_DEBUG) return;
    try {
      const componentIds = (Array.isArray(componentRows) ? componentRows : []).map(row => normalizeId(row?.item_id));
      const componentNames = componentIds.map(id => itemMap.get(id)?.name || id);
      const rows = [];
      for (const [vehicleId, state] of states.entries()) {
        const vehicle = state?.vehicle;
        const monthly = getMonthlyStatsForVehicle(monthlyMap, vehicleId, vehicle);
        const projectedAvg = calcProjectedAvgDistance(
          monthly.totalDistance,
          monthly.workedDays,
          toNumber(state?.assignedDistanceKm, 0),
          toNumber(state?.count, 0) > 0,
          estimateRowsReferenceDistanceKm(componentRows, itemMap, vehicle)
        );
        const currentAssignedDistanceKm = Math.max(0, toNumber(state?.assignedDistanceKm, 0));
        const monthlyAvgDistance = Math.max(0, toNumber(monthly?.avgDistance, 0));
        const isUnusedVehicle = currentAssignedDistanceKm <= 0 && (
          toNumber(monthly?.workedDays, 0) <= 0 || monthlyAvgDistance <= 0
        );
        rows.push({
          vehicleId,
          driverName: vehicle?.driverName || vehicle?.driver_name || '-',
          picked: normalizeId(best?.vehicleId) === normalizeId(vehicleId),
          canAccept: vehicleCanAcceptComponent(state, vehicle, componentRows, itemMap),
          count: Number(state?.count || 0),
          assignedDistanceKm: currentAssignedDistanceKm,
          monthlyAvgDistance,
          projectedAvg,
          totalDistance: monthly.totalDistance,
          workedDays: monthly.workedDays,
          isUnusedVehicle
        });
      }
      console.groupCollapsed(`[DispatchCore][REBUNDLE_PICK] component=${componentNames.join(', ')} picked=${best?.vehicleId || '-'} (${best?.monthlyAvgDistance ?? '-'})`);
      console.table(rows);
      console.groupEnd();
    } catch (error) {
      console.warn('[DispatchCore][REBUNDLE_PICK] log failed:', error);
    }
  }

  function compareNormalVehiclePreference(a, b, axisRep) {
    const aProjectedAvg = getProjectedAxisAverageDistance(a, axisRep);
    const bProjectedAvg = getProjectedAxisAverageDistance(b, axisRep);

    const aUnused = toNumber(a?.todayDistance, 0) <= 0 && (
      toNumber(a?.workedDays, 0) <= 0 || toNumber(a?.avgDistance, 0) <= 0
    );
    const bUnused = toNumber(b?.todayDistance, 0) <= 0 && (
      toNumber(b?.workedDays, 0) <= 0 || toNumber(b?.avgDistance, 0) <= 0
    );
    if (aUnused !== bUnused) return aUnused ? -1 : 1;

    if (toNumber(a?.avgDistance, 0) !== toNumber(b?.avgDistance, 0)) {
      return toNumber(a?.avgDistance, 0) - toNumber(b?.avgDistance, 0);
    }
    if (aProjectedAvg !== bProjectedAvg) return aProjectedAvg - bProjectedAvg;
    if (toNumber(a?.todayDistance, 0) !== toNumber(b?.todayDistance, 0)) {
      return toNumber(a?.todayDistance, 0) - toNumber(b?.todayDistance, 0);
    }
    if (toNumber(a?.totalDistance, 0) !== toNumber(b?.totalDistance, 0)) {
      return toNumber(a?.totalDistance, 0) - toNumber(b?.totalDistance, 0);
    }
    if (toNumber(a?.workedDays, 0) !== toNumber(b?.workedDays, 0)) {
      return toNumber(a?.workedDays, 0) - toNumber(b?.workedDays, 0);
    }
    return String(a.driverName).localeCompare(String(b.driverName), 'ja');
  }

  function chooseNormalVehicleAssignments(axisReps, vehicles) {
    const pool = Array.isArray(vehicles) ? [...vehicles] : [];
    const assignments = new Map();
    const used = new Set();

    [...axisReps].sort((a, b) => (b.distanceKm - a.distanceKm) || (b.size - a.size)).forEach(axisRep => {
      const freeVehicles = pool.filter(vehicle => !used.has(vehicle.vehicleId));
      const candidates = freeVehicles
        .filter(vehicle => vehicleCanTakeAxis(vehicle, axisRep))
        .sort((a, b) => compareNormalVehiclePreference(a, b, axisRep));

      const fallbackVehicles = [...freeVehicles].sort((a, b) => compareNormalVehiclePreference(a, b, axisRep));
      const picked = candidates[0] || fallbackVehicles[0];

      logNormalVehicleSelection('chooseNormalVehicleAssignments', axisRep, freeVehicles, candidates, picked);

      if (!picked) return;
      assignments.set(axisRep.axis.axisId, picked);
      used.add(picked.vehicleId);
    });

    return assignments;
  }

  function buildLastTripCandidate(axisRep, vehicle, origin) {
    if (!axisRep?.point || !vehicle?.homePoint) return null;
    if (!vehicleCanTakeAxis(vehicle, axisRep)) return null;
    const axisAngle = axisRep.angleFromOrigin;
    if (!Number.isFinite(axisAngle)) return null;
    const homeAngle = calcAngleFromOrigin(origin, vehicle.homePoint);
    const diff = angleDiff(axisAngle, homeAngle);
    return {
      axisId: axisRep.axis.axisId,
      vehicleId: vehicle.vehicleId,
      diff,
      homeDistanceKm: haversineKm(axisRep.point, vehicle.homePoint),
      axisDistanceKm: axisRep.distanceKm,
      size: axisRep.size
    };
  }

  function assignVehiclesToFixedAxes(fixedAxes, normalizedVehicles, origin) {
    const axisReps = fixedAxes.map(axis => getAxisRepresentative(axis, origin));
    const lastTripVehicles = normalizedVehicles.filter(vehicle => vehicle.isLastTrip);
    const normalVehicles = normalizedVehicles.filter(vehicle => !vehicle.isLastTrip);
    const assigned = new Map();
    const usedVehicles = new Set();
    const usedAxes = new Set();

    if (!axisReps.length) return assigned;

    if (!lastTripVehicles.length) {
      return chooseNormalVehicleAssignments(axisReps, normalizedVehicles);
    }

    const pairCandidates = [];
    axisReps.forEach(axisRep => {
      lastTripVehicles.forEach(vehicle => {
        const entry = buildLastTripCandidate(axisRep, vehicle, origin);
        if (!entry || entry.diff > LAST_TRIP_MAX_DIRECTION_DIFF) return;
        pairCandidates.push(entry);
      });
    });

    pairCandidates.sort((a, b) => {
      if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
      if (a.diff !== b.diff) return a.diff - b.diff;
      if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
      return a.vehicleId - b.vehicleId;
    });

    pairCandidates.forEach(entry => {
      if (usedVehicles.has(entry.vehicleId) || usedAxes.has(entry.axisId)) return;
      const vehicle = normalizedVehicles.find(v => v.vehicleId === entry.vehicleId);
      if (!vehicle) return;
      assigned.set(entry.axisId, vehicle);
      usedVehicles.add(entry.vehicleId);
      usedAxes.add(entry.axisId);
    });

    const remainingLastTrip = lastTripVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));

    remainingLastTrip.forEach(vehicle => {
      const choices = axisReps
        .filter(axisRep => !usedAxes.has(axisRep.axis.axisId))
        .map(axisRep => buildLastTripCandidate(axisRep, vehicle, origin))
        .filter(Boolean)
        .filter(entry => entry.diff <= LAST_TRIP_MAX_DIRECTION_DIFF)
        .sort((a, b) => {
          if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
          if (a.diff !== b.diff) return a.diff - b.diff;
          if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
          return a.vehicleId - b.vehicleId;
        });
      const picked = choices[0];
      if (!picked) return;
      assigned.set(picked.axisId, vehicle);
      usedVehicles.add(vehicle.vehicleId);
      usedAxes.add(picked.axisId);
    });

    let unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
    let freeNormalVehicles = normalVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));
    let freeLastTripVehicles = lastTripVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));

    if (unassignedAxes.length > 0 && freeLastTripVehicles.length > 0 && freeNormalVehicles.length < unassignedAxes.length) {
      freeLastTripVehicles.forEach(vehicle => {
        if (!unassignedAxes.length) return;
        if (freeNormalVehicles.length >= unassignedAxes.length) return;
        const choices = unassignedAxes
          .map(axisRep => buildLastTripCandidate(axisRep, vehicle, origin) || {
            axisId: axisRep.axis.axisId,
            vehicleId: vehicle.vehicleId,
            diff: Infinity,
            homeDistanceKm: vehicle.homePoint && axisRep.point ? haversineKm(axisRep.point, vehicle.homePoint) : Infinity,
            axisDistanceKm: axisRep.distanceKm,
            size: axisRep.size
          })
          .sort((a, b) => {
            if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
            if (a.diff !== b.diff) return a.diff - b.diff;
            if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
            return a.vehicleId - b.vehicleId;
          });
        const picked = choices[0];
        if (!picked) return;
        assigned.set(picked.axisId, vehicle);
        usedVehicles.add(vehicle.vehicleId);
        usedAxes.add(picked.axisId);
        unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
        freeNormalVehicles = normalVehicles.filter(v => !usedVehicles.has(v.vehicleId));
      });
    }

    unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
    const remainingVehicles = normalizedVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));
    const normalAssignments = chooseNormalVehicleAssignments(unassignedAxes, remainingVehicles);
    normalAssignments.forEach((vehicle, axisId) => {
      assigned.set(axisId, vehicle);
      usedVehicles.add(vehicle.vehicleId);
      usedAxes.add(axisId);
    });

    return assigned;
  }

  function applyVehicleAssignmentsToAxes(fixedAxes, vehicleAssignments) {
    fixedAxes.forEach(axis => {
      const vehicle = vehicleAssignments.get(axis.axisId);
      if (!vehicle) return;
      axis.vehicleId = vehicle.vehicleId;
      axis.driverName = vehicle.driverName;
      axis.capacity = vehicle.capacity;
    });
  }

  function isFriendlyArea(areaA, areaB) {
    const a = safeNormalizeArea(areaA || '');
    const b = safeNormalizeArea(areaB || '');
    if (!a || !b) return false;

    const canonicalA = getCanonicalAreaSafe(a);
    const canonicalB = getCanonicalAreaSafe(b);
    const groupA = getAreaDisplayGroupSafe(a);
    const groupB = getAreaDisplayGroupSafe(b);

    if (canonicalA && canonicalB && canonicalA === canonicalB) return true;
    if (groupA && groupB && groupA === groupB && getDirectionAffinityScoreSafe(a, b) > -20) return true;
    if (getAreaAffinityScoreSafe(a, b) >= 60) return true;
    if (getDirectionAffinityScoreSafe(a, b) >= 22) return true;
    return false;
  }

  function isHardReverseArea(areaA, areaB) {
    const a = safeNormalizeArea(areaA || '');
    const b = safeNormalizeArea(areaB || '');
    if (!a || !b) return false;
    return getAreaAffinityScoreSafe(a, b) <= 26 || getDirectionAffinityScoreSafe(a, b) <= -34;
  }

  function getRoundTripMinutesForItem(item) {
    const sendOnlyMinutes = Math.max(0, Math.round(toNumber(item?.travelMinutes, 0)));
    return sendOnlyMinutes * 2;
  }

  function buildItemMap(normalizedItems) {
    return new Map((Array.isArray(normalizedItems) ? normalizedItems : []).map(item => [item.itemId, item]));
  }

  function buildVehicleMap(vehicles) {
    return new Map((Array.isArray(vehicles) ? vehicles : []).map(vehicle => [toNumber(vehicle?.id, 0), vehicle]));
  }

  function buildAssignmentVehicleStates(assignments, itemMap, vehicles, hour) {
    const states = new Map();
    (Array.isArray(vehicles) ? vehicles : []).forEach(vehicle => {
      const vehicleId = toNumber(vehicle?.id, 0);
      states.set(vehicleId, { vehicle, count: 0, rows: [], areas: [], assignedDistanceKm: 0 });
    });

    (Array.isArray(assignments) ? assignments : [])
      .filter(row => toNumber(row?.actual_hour, 0) === toNumber(hour, 0))
      .forEach(row => {
        const vehicleId = toNumber(row?.vehicle_id, 0);
        if (!states.has(vehicleId)) return;
        const state = states.get(vehicleId);
        state.count += 1;
        state.rows.push(row);
        const item = itemMap.get(normalizeId(row?.item_id));
        const area = safeNormalizeArea(item?.area || '');
        if (area) state.areas.push(area);
      });

    states.forEach(state => {
      state.assignedDistanceKm = estimateRowsReferenceDistanceKm(state.rows, itemMap, state.vehicle);
    });

    return states;
  }

  function findLongFriendlyComponents(hourRows, itemMap) {
    const eligibleRows = (Array.isArray(hourRows) ? hourRows : [])
      .filter(Boolean)
      .filter(row => {
        const item = itemMap.get(normalizeId(row?.item_id));
        return item && getRoundTripMinutesForItem(item) >= LONG_ROUTE_BUNDLE_MINUTES;
      });

    const visited = new Set();
    const components = [];

    for (const row of eligibleRows) {
      const rowId = normalizeId(row?.item_id);
      if (!rowId || visited.has(rowId)) continue;

      const stack = [row];
      const component = [];
      visited.add(rowId);

      while (stack.length) {
        const current = stack.pop();
        const currentId = normalizeId(current?.item_id);
        const currentItem = itemMap.get(currentId);
        if (!currentItem) continue;
        component.push(current);
        const currentArea = currentItem.area;

        for (const other of eligibleRows) {
          const otherId = normalizeId(other?.item_id);
          if (!otherId || visited.has(otherId)) continue;
          const otherItem = itemMap.get(otherId);
          if (!otherItem) continue;
          const otherArea = otherItem.area;
          if (isFriendlyArea(currentArea, otherArea) && !isHardReverseArea(currentArea, otherArea)) {
            visited.add(otherId);
            stack.push(other);
          }
        }
      }

      if (component.length >= 2) components.push(component);
    }

    return components;
  }

  function vehicleCanAcceptComponent(state, vehicle, componentRows, itemMap) {
    const seat = Math.max(1, toNumber(vehicle?.seat_capacity ?? vehicle?.capacity, 4));
    const componentIds = new Set((Array.isArray(componentRows) ? componentRows : []).map(row => normalizeId(row?.item_id)));
    const currentRows = Array.isArray(state?.rows) ? state.rows : [];
    const nonComponentRows = currentRows.filter(row => !componentIds.has(normalizeId(row?.item_id)));
    if (nonComponentRows.length + componentRows.length > seat) return false;

    const existingAreas = nonComponentRows
      .map(row => itemMap.get(normalizeId(row?.item_id)))
      .map(item => safeNormalizeArea(item?.area || ''))
      .filter(Boolean);

    for (const row of componentRows) {
      const item = itemMap.get(normalizeId(row?.item_id));
      const area = safeNormalizeArea(item?.area || '');
      if (existingAreas.some(existingArea => isHardReverseArea(area, existingArea))) return false;
    }

    return true;
  }

  function getComponentHomeBonus(componentRows, itemMap, vehicle, origin) {
    if (!vehicle?.homePoint) return 0;
    const componentItems = (Array.isArray(componentRows) ? componentRows : [])
      .map(row => itemMap.get(normalizeId(row?.item_id)))
      .filter(Boolean);
    const farthestItem = [...componentItems].sort(byDistanceDesc)[0] || null;
    if (!farthestItem?.point) return 0;

    const axisAngle = Number.isFinite(farthestItem.angleFromOrigin) ? farthestItem.angleFromOrigin : null;
    if (!Number.isFinite(axisAngle)) return 0;
    const homeAngle = calcAngleFromOrigin(origin, vehicle.homePoint);
    const diff = angleDiff(axisAngle, homeAngle);
    const homeDistanceKm = haversineKm(farthestItem.point, vehicle.homePoint);

    let score = 0;
    if (vehicle.isLastTrip && diff <= LAST_TRIP_MAX_DIRECTION_DIFF) score += 140;
    score -= diff * 0.7;
    if (Number.isFinite(homeDistanceKm)) score -= homeDistanceKm * 2.2;
    return score;
  }

  function pickTargetVehicleForComponent(componentRows, states, monthlyMap, itemMap, origin) {
    const componentIds = new Set(componentRows.map(row => normalizeId(row?.item_id)));
    const componentItems = componentRows
      .map(row => itemMap.get(normalizeId(row?.item_id)))
      .filter(Boolean);
    const componentAreas = componentItems.map(item => safeNormalizeArea(item?.area || '')).filter(Boolean);

    let best = null;

    for (const [vehicleId, state] of states.entries()) {
      const vehicle = state?.vehicle;
      if (!vehicle) continue;
      if (!vehicleCanAcceptComponent(state, vehicle, componentRows, itemMap)) continue;

      const currentRows = Array.isArray(state.rows) ? state.rows : [];
      const componentAlreadyHere = currentRows.filter(row => componentIds.has(normalizeId(row?.item_id))).length;
      const nonComponentAreas = currentRows
        .filter(row => !componentIds.has(normalizeId(row?.item_id)))
        .map(row => itemMap.get(normalizeId(row?.item_id)))
        .map(item => safeNormalizeArea(item?.area || ''))
        .filter(Boolean);

      let compatibilityScore = 0;
      compatibilityScore += componentAlreadyHere * 900;
      if (componentAlreadyHere > 0) compatibilityScore += 250;

      for (const componentArea of componentAreas) {
        for (const existingArea of nonComponentAreas) {
          if (isFriendlyArea(componentArea, existingArea)) compatibilityScore += 120;
          if (isHardReverseArea(componentArea, existingArea)) compatibilityScore -= 600;
        }
      }

      compatibilityScore -= toNumber(state?.count, 0) * 6;
      compatibilityScore += getComponentHomeBonus(componentRows, itemMap, vehicle, origin);

      const monthly = getMonthlyStatsForVehicle(monthlyMap, vehicleId, vehicle);
      const projectedAvg = calcProjectedAvgDistance(
        monthly.totalDistance,
        monthly.workedDays,
        toNumber(state?.assignedDistanceKm, 0),
        toNumber(state?.count, 0) > 0,
        estimateRowsReferenceDistanceKm(componentRows, itemMap, vehicle)
      );

      const tieCurrentAvg = calcProjectedAvgDistance(
        monthly.totalDistance,
        monthly.workedDays,
        toNumber(state?.assignedDistanceKm, 0),
        toNumber(state?.count, 0) > 0,
        0
      );

      const monthlyAvgDistance = Math.max(0, toNumber(monthly?.avgDistance, 0));
      const currentAssignedDistanceKm = Math.max(0, toNumber(state?.assignedDistanceKm, 0));
      const isUnusedVehicle = currentAssignedDistanceKm <= 0 && (
        toNumber(monthly?.workedDays, 0) <= 0 || monthlyAvgDistance <= 0
      );
      const preferByAverageFirst = !vehicle?.isLastTrip;

      const isBetterNormal = !best ||
        (isUnusedVehicle !== best.isUnusedVehicle && isUnusedVehicle) ||
        (isUnusedVehicle === best.isUnusedVehicle && monthlyAvgDistance !== best.monthlyAvgDistance && monthlyAvgDistance < best.monthlyAvgDistance) ||
        (isUnusedVehicle === best.isUnusedVehicle && monthlyAvgDistance === best.monthlyAvgDistance && projectedAvg !== best.projectedAvg && projectedAvg < best.projectedAvg) ||
        (isUnusedVehicle === best.isUnusedVehicle && monthlyAvgDistance === best.monthlyAvgDistance && projectedAvg === best.projectedAvg && currentAssignedDistanceKm !== best.currentAssignedDistanceKm && currentAssignedDistanceKm < best.currentAssignedDistanceKm) ||
        (isUnusedVehicle === best.isUnusedVehicle && monthlyAvgDistance === best.monthlyAvgDistance && projectedAvg === best.projectedAvg && currentAssignedDistanceKm === best.currentAssignedDistanceKm && compatibilityScore !== best.compatibilityScore && compatibilityScore > best.compatibilityScore) ||
        (isUnusedVehicle === best.isUnusedVehicle && monthlyAvgDistance === best.monthlyAvgDistance && projectedAvg === best.projectedAvg && currentAssignedDistanceKm === best.currentAssignedDistanceKm && compatibilityScore === best.compatibilityScore && monthly.totalDistance < best.totalDistance);

      const isBetterLastTrip = !best ||
        (compatibilityScore !== best.compatibilityScore && compatibilityScore > best.compatibilityScore) ||
        (compatibilityScore === best.compatibilityScore && projectedAvg !== best.projectedAvg && projectedAvg < best.projectedAvg) ||
        (compatibilityScore === best.compatibilityScore && projectedAvg === best.projectedAvg && currentAssignedDistanceKm !== best.currentAssignedDistanceKm && currentAssignedDistanceKm < best.currentAssignedDistanceKm) ||
        (compatibilityScore === best.compatibilityScore && projectedAvg === best.projectedAvg && currentAssignedDistanceKm === best.currentAssignedDistanceKm && monthlyAvgDistance !== best.monthlyAvgDistance && monthlyAvgDistance < best.monthlyAvgDistance) ||
        (compatibilityScore === best.compatibilityScore && projectedAvg === best.projectedAvg && currentAssignedDistanceKm === best.currentAssignedDistanceKm && monthlyAvgDistance === best.monthlyAvgDistance && monthly.totalDistance < best.totalDistance);

      if ((preferByAverageFirst && isBetterNormal) || (!preferByAverageFirst && isBetterLastTrip)) {
        best = {
          vehicleId,
          projectedAvg,
          compatibilityScore,
          tieCurrentAvg,
          currentAssignedDistanceKm,
          monthlyAvgDistance,
          totalDistance: monthly.totalDistance,
          isUnusedVehicle
        };
      }
    }

    logComponentVehicleSelection(componentRows, states, best, itemMap, monthlyMap, origin);
    return best?.vehicleId || null;
  }

  function reindexAssignments(assignments) {
    const working = (Array.isArray(assignments) ? assignments : []).map(row => ({ ...row }));
    const grouped = new Map();

    working.forEach(row => {
      const vehicleId = toNumber(row?.vehicle_id, 0);
      const hour = toNumber(row?.actual_hour, 0);
      const key = `${vehicleId}__${hour}`;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(row);
    });

    grouped.forEach(rows => {
      rows.sort((a, b) => {
        if (toNumber(a?.distance_km, 0) !== toNumber(b?.distance_km, 0)) {
          return toNumber(a?.distance_km, 0) - toNumber(b?.distance_km, 0);
        }
        return String(normalizeId(a?.item_id)).localeCompare(String(normalizeId(b?.item_id)), 'ja');
      });
      rows.forEach((row, index) => {
        row.stop_order = index + 1;
      });
    });

    return working;
  }

  function rebundleLongFriendlyAssignments(assignments, normalizedItems, vehicles, monthlyMap, origin) {
    let rows = (Array.isArray(assignments) ? assignments : []).map(row => ({ ...row }));
    if (!rows.length) return rows;

    const itemMap = buildItemMap(normalizedItems);
    const vehicleMap = buildVehicleMap(vehicles);
    const hours = [...new Set(rows.map(row => toNumber(row?.actual_hour, 0)))];

    for (const hour of hours) {
      const hourRows = rows.filter(row => toNumber(row?.actual_hour, 0) === hour);
      const components = findLongFriendlyComponents(hourRows, itemMap);
      if (!components.length) continue;

      for (const componentRows of components) {
        const currentVehicles = [...new Set(componentRows.map(row => toNumber(row?.vehicle_id, 0)).filter(Boolean))];
        if (currentVehicles.length <= 1) continue;

        const states = buildAssignmentVehicleStates(rows, itemMap, vehicles, hour);
        const targetVehicleId = pickTargetVehicleForComponent(componentRows, states, monthlyMap, itemMap, origin);
        if (!targetVehicleId) continue;

        const targetVehicle = vehicleMap.get(targetVehicleId);
        if (!targetVehicle) continue;

        const componentIds = new Set(componentRows.map(row => normalizeId(row?.item_id)));
        rows = rows.map(row => {
          if (!componentIds.has(normalizeId(row?.item_id))) return row;
          return {
            ...row,
            vehicle_id: targetVehicleId,
            driver_name: String(targetVehicle?.driver_name || targetVehicle?.name || targetVehicle?.plate_number || '').trim() || '-'
          };
        });
      }
    }

    return reindexAssignments(rows);
  }

  function optimizeAssignments(items, vehicles, monthlyMap, options = {}) {
    const origin = getOrigin();
    const normalizedVehicles = (Array.isArray(vehicles) ? vehicles : [])
      .map(vehicle => normalizeVehicle(vehicle, monthlyMap))
      .filter(vehicle => vehicle.vehicleId > 0 && vehicle.capacity > 0);

    const normalizedItems = (Array.isArray(items) ? items : [])
      .map(item => normalizeItem(item, origin))
      .filter(item => item.itemId && Number.isFinite(item.distanceKm) && Number.isFinite(item.angleFromOrigin));

    if (!normalizedVehicles.length || !normalizedItems.length) {
      global.__THEMIS_LAST_OVERFLOW__ = buildOverflowMeta([]);
      return [];
    }

    const vehicleCount = Math.min(normalizedVehicles.length, normalizedItems.length);
    const threshold = Number(options?.axisThreshold || AREA_MEMBER_ANGLE_THRESHOLD);

    const fixedAxes = selectFixedAxes(normalizedItems, normalizedVehicles, threshold);
    try {
      debugLog('[DispatchCore][AXES]', fixedAxes.map(axis => ({
        axisId: axis.axisId,
        vehicleId: axis.vehicleId,
        driverName: axis.driverName,
        anchorName: axis.anchorName,
        anchorDistanceKm: Number(axis.anchorDistanceKm || 0),
        anchorAngle: Number(axis.anchorAngle || 0)
      })));
    } catch (error) {
      debugWarn('[DispatchCore][AXES] log failed:', error);
    }

    const unsettled = pushPendingMembers(normalizedItems, fixedAxes, threshold);
    const overflow = resolveUnsettledMembers(unsettled, fixedAxes);
    finalizeAxisOrder(fixedAxes);

    const mergeResult = collapseSameDirectionAxes(fixedAxes, normalizedVehicles, origin);
    const mergedAxes = Array.isArray(mergeResult?.axes) ? mergeResult.axes : fixedAxes;

    const vehicleAssignments = assignVehiclesToFixedAxes(mergedAxes, normalizedVehicles, origin);
    applyVehicleAssignmentsToAxes(mergedAxes, vehicleAssignments);

    let assignments = buildAssignments(mergedAxes);
    assignments = rebundleLongFriendlyAssignments(assignments, normalizedItems, vehicles, monthlyMap, origin);

    if (DISPATCH_DEBUG) {
      try {
        console.groupCollapsed(`[DispatchCore][FINAL_ASSIGNMENTS] version=${VERSION}`);
        console.table(normalizedVehicles.map(vehicle => ({
          vehicleId: vehicle.vehicleId,
          driverName: vehicle.driverName,
          avgDistance: vehicle.avgDistance,
          totalDistance: vehicle.totalDistance,
          workedDays: vehicle.workedDays,
          todayDistance: vehicle.todayDistance,
          isLastTrip: vehicle.isLastTrip
        })));
        console.table(assignments.map(row => ({
          item_id: row.item_id,
          item_name: normalizedItems.find(item => normalizeId(item.itemId) === normalizeId(row.item_id))?.name || '',
          vehicle_id: row.vehicle_id,
          driver_name: row.driver_name,
          actual_hour: row.actual_hour,
          stop_order: row.stop_order
        })));
        console.groupEnd();
        debugLog('[DispatchCore][RESULT]', mergedAxes.map(axis => ({
          axisId: axis.axisId,
          vehicleId: axis.vehicleId,
          driverName: axis.driverName,
          anchorName: axis.anchorName,
          memberNames: axis.members.map(item => item.name),
          memberCount: axis.members.length,
          capacity: axis.capacity
        })));
        debugLog('[DispatchCore][ASSIGNMENTS]', assignments.map(row => ({
          item_id: row.item_id,
          vehicle_id: row.vehicle_id,
          driver_name: row.driver_name,
          actual_hour: row.actual_hour,
          stop_order: row.stop_order
        })));
      } catch (error) {
        debugWarn('[DispatchCore][RESULT] log failed:', error);
      }
    }

    global.__THEMIS_LAST_OVERFLOW__ = {
      ...buildOverflowMeta(overflow),
      fixedAxes: mergedAxes.map(axis => ({
        axisId: axis.axisId,
        vehicleId: axis.vehicleId,
        driverName: axis.driverName,
        anchorItemId: axis.anchorItemId,
        anchorName: axis.anchorName,
        anchorAngle: axis.anchorAngle,
        anchorDistanceKm: axis.anchorDistanceKm,
        memberItemIds: axis.members.map(item => item.itemId)
      })),
      totalSeatCapacity: mergedAxes.reduce((sum, axis) => sum + Number(axis.capacity || 0), 0),
      totalCastCount: normalizedItems.length,
      standbyVehicleCount: Number(mergeResult?.standbyCount || 0),
      mergedSameDirectionAxes: Array.isArray(mergeResult?.mergeLogs) ? mergeResult.mergeLogs : [],
      version: VERSION
    };

    return assignments;
  }

  function runDispatchPlan(origin, vehicles, people) {
    return optimizeAssignments(people, vehicles, null, {});
  }

  global.DispatchCore = {
    VERSION,
    optimizeAssignments,
    runDispatchPlan
  };
})(window);
