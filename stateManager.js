// stateManager.js
// THEMIS AI Dispatch 用 共通 state 管理
// 目的:
// - dashboard.js に溜まっている共有stateを一元管理する
// - まずは supabaseService.js の代入先をここへ寄せる
// - 段階移行のため、window 経由でも読めるようにする

(function () {
  const ROOT = typeof window !== "undefined" ? window : globalThis;

  const DEFAULT_STATE = {
    currentUser: null,
    currentDispatchId: null,

    allCastsCache: [],
    allVehiclesCache: [],
    currentPlansCache: [],
    currentActualsCache: [],
    currentDailyReportsCache: [],
    currentMileageExportRows: [],

    activeVehicleIdsForToday: new Set(),

    lastAutoDispatchRunAtMinutes: null,
    simulationSlotHour: null,
    lastSimulationResult: null,
    isRefreshingHybridUI: false,
    suppressSimulationSlotChange: false
  };

  const state = {
    ...DEFAULT_STATE,
    activeVehicleIdsForToday: new Set(DEFAULT_STATE.activeVehicleIdsForToday)
  };

  function isPlainObject(value) {
    return !!value && typeof value === "object" && !Array.isArray(value) && !(value instanceof Set);
  }

  function cloneValue(value) {
    if (value instanceof Set) return new Set(value);
    if (Array.isArray(value)) return [...value];
    if (isPlainObject(value)) return { ...value };
    return value;
  }

  function normalizeArray(value) {
    return Array.isArray(value) ? value : [];
  }

  function normalizeSet(value) {
    if (value instanceof Set) return new Set(value);
    if (Array.isArray(value)) return new Set(value);
    return new Set();
  }

  function syncWindowRefs() {
    // 段階移行用の可視化ミラー
    // 注意:
    // dashboard.js の top-level let そのものは直接更新できないため、
    // 今後は参照側も stateManager 経由に寄せていく
    ROOT.__APP_STATE__ = state;
    ROOT.__APP_STATE_SNAPSHOT__ = getAll();

    ROOT.currentUser = state.currentUser;
    ROOT.currentDispatchId = state.currentDispatchId;

    ROOT.allCastsCache = [...state.allCastsCache];
    ROOT.allVehiclesCache = [...state.allVehiclesCache];
    ROOT.currentPlansCache = [...state.currentPlansCache];
    ROOT.currentActualsCache = [...state.currentActualsCache];
    ROOT.currentDailyReportsCache = [...state.currentDailyReportsCache];
    ROOT.currentMileageExportRows = [...state.currentMileageExportRows];

    ROOT.activeVehicleIdsForToday = new Set(state.activeVehicleIdsForToday);

    ROOT.lastAutoDispatchRunAtMinutes = state.lastAutoDispatchRunAtMinutes;
    ROOT.simulationSlotHour = state.simulationSlotHour;
    ROOT.lastSimulationResult = state.lastSimulationResult;
    ROOT.isRefreshingHybridUI = state.isRefreshingHybridUI;
    ROOT.suppressSimulationSlotChange = state.suppressSimulationSlotChange;
  }

  function get(key) {
    if (!(key in state)) {
      throw new Error(`stateManager.get: unknown key "${key}"`);
    }
    return cloneValue(state[key]);
  }

  function set(key, value) {
    if (!(key in state)) {
      throw new Error(`stateManager.set: unknown key "${key}"`);
    }

    switch (key) {
      case "allCastsCache":
      case "allVehiclesCache":
      case "currentPlansCache":
      case "currentActualsCache":
      case "currentDailyReportsCache":
      case "currentMileageExportRows":
        state[key] = normalizeArray(value);
        break;

      case "activeVehicleIdsForToday":
        state[key] = normalizeSet(value);
        break;

      case "isRefreshingHybridUI":
      case "suppressSimulationSlotChange":
        state[key] = Boolean(value);
        break;

      default:
        state[key] = value;
        break;
    }

    syncWindowRefs();
    return cloneValue(state[key]);
  }

  function patch(key, partial) {
    if (!(key in state)) {
      throw new Error(`stateManager.patch: unknown key "${key}"`);
    }
    if (!isPlainObject(state[key])) {
      throw new Error(`stateManager.patch: "${key}" is not a plain object`);
    }
    if (!isPlainObject(partial)) {
      throw new Error(`stateManager.patch: partial for "${key}" must be a plain object`);
    }

    state[key] = {
      ...state[key],
      ...partial
    };

    syncWindowRefs();
    return { ...state[key] };
  }

  function reset(key) {
    if (typeof key === "string") {
      if (!(key in DEFAULT_STATE)) {
        throw new Error(`stateManager.reset: unknown key "${key}"`);
      }
      state[key] = cloneValue(DEFAULT_STATE[key]);
      syncWindowRefs();
      return get(key);
    }

    Object.keys(DEFAULT_STATE).forEach(k => {
      state[k] = cloneValue(DEFAULT_STATE[k]);
    });
    syncWindowRefs();
    return getAll();
  }

  function getAll() {
    return {
      currentUser: state.currentUser,
      currentDispatchId: state.currentDispatchId,

      allCastsCache: [...state.allCastsCache],
      allVehiclesCache: [...state.allVehiclesCache],
      currentPlansCache: [...state.currentPlansCache],
      currentActualsCache: [...state.currentActualsCache],
      currentDailyReportsCache: [...state.currentDailyReportsCache],
      currentMileageExportRows: [...state.currentMileageExportRows],

      activeVehicleIdsForToday: new Set(state.activeVehicleIdsForToday),

      lastAutoDispatchRunAtMinutes: state.lastAutoDispatchRunAtMinutes,
      simulationSlotHour: state.simulationSlotHour,
      lastSimulationResult: state.lastSimulationResult,
      isRefreshingHybridUI: state.isRefreshingHybridUI,
      suppressSimulationSlotChange: state.suppressSimulationSlotChange
    };
  }

  function replaceAll(nextState = {}) {
    Object.keys(DEFAULT_STATE).forEach(key => {
      if (Object.prototype.hasOwnProperty.call(nextState, key)) {
        set(key, nextState[key]);
      } else {
        state[key] = cloneValue(DEFAULT_STATE[key]);
      }
    });
    syncWindowRefs();
    return getAll();
  }

  const api = {
    get,
    set,
    patch,
    reset,
    getAll,
    replaceAll,

    getCurrentUser() {
      return state.currentUser;
    },
    setCurrentUser(value) {
      return set("currentUser", value);
    },

    getCurrentDispatchId() {
      return state.currentDispatchId;
    },
    setCurrentDispatchId(value) {
      return set("currentDispatchId", value);
    },

    getCasts() {
      return [...state.allCastsCache];
    },
    setCasts(value) {
      return set("allCastsCache", value);
    },

    getVehicles() {
      return [...state.allVehiclesCache];
    },
    setVehicles(value) {
      return set("allVehiclesCache", value);
    },

    getPlans() {
      return [...state.currentPlansCache];
    },
    setPlans(value) {
      return set("currentPlansCache", value);
    },

    getActuals() {
      return [...state.currentActualsCache];
    },
    setActuals(value) {
      return set("currentActualsCache", value);
    },

    getDailyReports() {
      return [...state.currentDailyReportsCache];
    },
    setDailyReports(value) {
      return set("currentDailyReportsCache", value);
    },

    getMileageExportRows() {
      return [...state.currentMileageExportRows];
    },
    setMileageExportRows(value) {
      return set("currentMileageExportRows", value);
    },

    getActiveVehicleIds() {
      return new Set(state.activeVehicleIdsForToday);
    },
    setActiveVehicleIds(value) {
      return set("activeVehicleIdsForToday", value);
    },
    addActiveVehicleId(vehicleId) {
      const next = new Set(state.activeVehicleIdsForToday);
      next.add(Number(vehicleId || 0));
      return set("activeVehicleIdsForToday", next);
    },
    removeActiveVehicleId(vehicleId) {
      const next = new Set(state.activeVehicleIdsForToday);
      next.delete(Number(vehicleId || 0));
      return set("activeVehicleIdsForToday", next);
    },
    hasActiveVehicleId(vehicleId) {
      return state.activeVehicleIdsForToday.has(Number(vehicleId || 0));
    },

    getLastAutoDispatchRunAtMinutes() {
      return state.lastAutoDispatchRunAtMinutes;
    },
    setLastAutoDispatchRunAtMinutes(value) {
      return set("lastAutoDispatchRunAtMinutes", value);
    },

    getSimulationSlotHour() {
      return state.simulationSlotHour;
    },
    setSimulationSlotHour(value) {
      return set("simulationSlotHour", value);
    },

    getLastSimulationResult() {
      return state.lastSimulationResult;
    },
    setLastSimulationResult(value) {
      return set("lastSimulationResult", value);
    },

    getIsRefreshingHybridUI() {
      return Boolean(state.isRefreshingHybridUI);
    },
    setIsRefreshingHybridUI(value) {
      return set("isRefreshingHybridUI", value);
    },

    getSuppressSimulationSlotChange() {
      return Boolean(state.suppressSimulationSlotChange);
    },
    setSuppressSimulationSlotChange(value) {
      return set("suppressSimulationSlotChange", value);
    }
  };

  ROOT.stateManager = api;

  // dashboard.js 側のローカルstateとは別物なので、
  // 各 setter 関数から同期させる前提で使う。
  syncWindowRefs();
})();