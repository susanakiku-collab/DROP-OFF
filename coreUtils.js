// coreUtils
// dispatchCore / dashboard 共通の不足依存を吸収する軽量ブリッジ
(function () {
  const ROOT = typeof window !== 'undefined' ? window : globalThis;

  function getStateValue(key, fallback) {
    try {
      if (ROOT.stateManager?.get) {
        const value = ROOT.stateManager.get(key);
        return value === undefined ? fallback : value;
      }
    } catch (_) {}
    return fallback;
  }

  function asArray(value) {
    return Array.isArray(value) ? value : [];
  }

  if (typeof ROOT.__getCoreStateValue !== 'function') {
    ROOT.__getCoreStateValue = getStateValue;
  }

  ROOT.__getCoreActuals = function () {
    return asArray(getStateValue('currentActualsCache', ROOT.currentActualsCache ?? []));
  };

  ROOT.__getCorePlans = function () {
    return asArray(getStateValue('currentPlansCache', ROOT.currentPlansCache ?? []));
  };

  ROOT.__getCoreDailyReports = function () {
    return asArray(getStateValue('currentDailyReportsCache', ROOT.currentDailyReportsCache ?? []));
  };

  ROOT.__getCoreVehicles = function () {
    return asArray(getStateValue('allVehiclesCache', ROOT.allVehiclesCache ?? []));
  };

  ROOT.__getCoreEls = function () {
    return ROOT.els || ROOT.__THEMIS_ELS__ || {};
  };

  if (typeof ROOT.getExpectedDepartureDelayMinutes !== 'function') {
    ROOT.getExpectedDepartureDelayMinutes = function (baseHour) {
      const hour = Number(baseHour || 0);
      if (hour <= 2) return 20;
      if (hour === 3) return 18;
      if (hour === 4) return 12;
      if (hour >= 5) return 8;
      return 20;
    };
  }

  if (typeof ROOT.parseClockTextToMinutes !== 'function') {
    ROOT.parseClockTextToMinutes = function (value) {
      if (!value || value === '-') return null;
      const match = String(value).match(/^(\d{1,2}):(\d{2})$/);
      if (!match) return null;
      return Number(match[1]) * 60 + Number(match[2]);
    };
  }

  ROOT.getVehicleRotationForecastSafe = function (vehicle, orderedRows) {
    try {
      if (ROOT.DispatchCore?.calcVehicleRotationForecast) {
        return ROOT.DispatchCore.calcVehicleRotationForecast(vehicle, orderedRows);
      }
    } catch (error) {
      console.warn('DispatchCore.calcVehicleRotationForecast fallback:', error);
    }

    try {
      if (typeof ROOT.calcVehicleRotationForecast === 'function') {
        return ROOT.calcVehicleRotationForecast(vehicle, orderedRows);
      }
    } catch (error) {
      console.warn('calcVehicleRotationForecast fallback:', error);
    }

    try {
      if (typeof ROOT.calcVehicleRotationForecastGlobal === 'function') {
        return ROOT.calcVehicleRotationForecastGlobal(vehicle, orderedRows);
      }
    } catch (error) {
      console.warn('calcVehicleRotationForecastGlobal fallback:', error);
    }

    return {
      routeDistanceKm: 0,
      returnDistanceKm: 0,
      zoneLabel: '-',
      predictedDepartureTime: '-',
      predictedReturnTime: '-',
      predictedReadyTime: '-',
      predictedReturnMinutes: 0,
      extraSharedDelayMinutes: 0,
      stopCount: Array.isArray(orderedRows) ? orderedRows.length : 0,
      returnAfterLabel: '-'
    };
  };

  ROOT.getSelectedVehiclesForToday = ROOT.getSelectedVehiclesForToday || function () {
    const vehicles = ROOT.__getCoreVehicles();
    const activeIdsValue = getStateValue('activeVehicleIdsForToday', ROOT.activeVehicleIdsForToday ?? new Set());
    const activeIds = activeIdsValue instanceof Set ? activeIdsValue : new Set(activeIdsValue || []);
    return vehicles.filter(vehicle => activeIds.has(Number(vehicle?.id || 0)));
  };
})();
