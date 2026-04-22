// supabaseService
// supabase table mapping for DROP OFF
window.DROP_OFF_TABLES = Object.assign({
  profiles: "dropoff_profiles",
  casts: "dropoff_people",
  vehicles: "dropoff_vehicles",
  origins: "dropoff_origins",
  teams: "dropoff_teams",
  team_members: "dropoff_team_members",
  invitations: "dropoff_invitations",
  dispatches: "dropoff_dispatches",
  dispatch_plans: "dropoff_dispatch_plans",
  dispatch_items: "dropoff_dispatch_items",
  dispatch_history: "dropoff_dispatch_history",
  vehicle_daily_reports: "dropoff_vehicle_daily_reports",
  vehicle_daily_runs: "dropoff_vehicle_daily_runs"
}, window.DROP_OFF_TABLES || {});

function getTableName(logicalName) {
  return window.DROP_OFF_TABLES?.[logicalName] || logicalName;
}

function getRelationName(logicalName) {
  return getTableName(logicalName);
}

function remapRelationSelect(selectText) {
  return String(selectText || "")
    .replace(/\bcasts\s*\(/g, `${getRelationName("casts")} (`)
    .replace(/\bvehicles\s*\(/g, `${getRelationName("vehicles")} (`);
}

function isMissingTableError(error) {
  const code = String(error?.code || "").trim();
  const message = String(error?.message || "");
  return code === "PGRST205" || code === "42P01" || /Could not find the table/i.test(message) || /schema cache/i.test(message);
}

const __dropOffMissingTableWarnings = new Set();
const __dropOffMissingTables = new Set();
function markKnownMissingTable(logicalName) {
  __dropOffMissingTables.add(String(logicalName || ""));
  __dropOffMissingTables.add(String(getTableName(logicalName) || ""));
}
function isKnownMissingTable(logicalName) {
  return __dropOffMissingTables.has(String(logicalName || "")) || __dropOffMissingTables.has(String(getTableName(logicalName) || ""));
}
function warnMissingTableOnce(logicalName, error) {
  markKnownMissingTable(logicalName);
  const key = `${logicalName}:${String(error?.code || "")}:${String(error?.message || "")}`;
  if (__dropOffMissingTableWarnings.has(key)) return;
  __dropOffMissingTableWarnings.add(key);
  console.warn(`${getTableName(logicalName)} テーブルが未作成のため処理をスキップ:`, error);
}

function isMissingColumnError(error) {
  const code = String(error?.code || "").trim();
  const message = String(error?.message || "");
  return code === "PGRST204" || code === "42703" || /Could not find the '.+?' column/i.test(message) || /column .+ does not exist/i.test(message);
}

function getMissingColumnName(error) {
  const message = String(error?.message || "");
  let m = message.match(/Could not find the '([^']+)' column/i);
  if (m) return m[1];
  m = message.match(/column\s+[^.]+\.([^\s]+)\s+does not exist/i);
  if (m) return m[1].replace(/['"]/g, "");
  return null;
}

const DROP_OFF_OPTIONAL_DISABLED_TABLES = new Set(
  Array.isArray(window.DROP_OFF_DISABLED_TABLES)
    ? window.DROP_OFF_DISABLED_TABLES.map(String)
    : ["dispatch_plans", "dispatch_items", "dispatch_history", "vehicle_daily_reports"].map(String)
);
for (const logicalName of DROP_OFF_OPTIONAL_DISABLED_TABLES) {
  markKnownMissingTable(logicalName);
}

function compareValuesForSort(a, b) {
  const aEmpty = a === null || a === undefined || a === "";
  const bEmpty = b === null || b === undefined || b === "";
  if (aEmpty && bEmpty) return 0;
  if (aEmpty) return 1;
  if (bEmpty) return -1;
  const aNum = Number(a);
  const bNum = Number(b);
  const aIsNum = Number.isFinite(aNum) && String(a).trim() !== "";
  const bIsNum = Number.isFinite(bNum) && String(b).trim() !== "";
  if (aIsNum && bIsNum) return aNum - bNum;
  return String(a).localeCompare(String(b), "ja");
}

function sortRowsClientSide(rows, orderSpecs) {
  const specs = Array.isArray(orderSpecs) ? orderSpecs.filter(spec => spec?.column) : [];
  if (!specs.length) return Array.isArray(rows) ? [...rows] : [];
  return [...(Array.isArray(rows) ? rows : [])].sort((left, right) => {
    for (const spec of specs) {
      const result = compareValuesForSort(left?.[spec.column], right?.[spec.column]);
      if (result !== 0) return spec.ascending === false ? -result : result;
    }
    return 0;
  });
}

async function selectRowsClientSideSafe(tableName, logicalName, orderSpecs, options = {}) {
  if (isKnownMissingTable(logicalName)) {
    return { data: [], error: null };
  }

  const requestedTeamId = String(options?.teamId || '').trim() || null;

  let query = supabaseClient
    .from(tableName)
    .select("*");

  if (requestedTeamId) {
    query = query.eq('team_id', requestedTeamId);
  }

  let { data, error } = await query;

  if (error && requestedTeamId && isMissingColumnError(error)) {
    const fallback = await supabaseClient
      .from(tableName)
      .select('*');
    data = fallback.data;
    error = fallback.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce(logicalName, error);
    }
    return { data: null, error };
  }

  let rows = Array.isArray(data) ? [...data] : [];
  rows = rows.filter(row => row?.is_active !== false);
  if (requestedTeamId) {
    rows = rows.filter(row => String(row?.team_id || '').trim() === requestedTeamId);
  }
  rows = sortRowsClientSide(rows, orderSpecs);
  return { data: rows, error: null };
}

function omitKeys(obj, keys) {
  const drop = new Set((Array.isArray(keys) ? keys : [keys]).filter(Boolean));
  return Object.fromEntries(Object.entries(obj || {}).filter(([k]) => !drop.has(k)));
}

function normalizeCastRecordForApp(row = {}) {
  return {
    ...row,
    latitude: row.latitude ?? row.lat ?? null,
    longitude: row.longitude ?? row.lng ?? null,
    is_active: row.is_active !== false
  };
}

function buildCastTablePayload(basePayload = {}) {
  return {
    ...basePayload,
    lat: basePayload.latitude ?? basePayload.lat ?? null,
    lng: basePayload.longitude ?? basePayload.lng ?? null
  };
}

const DROP_OFF_VEHICLE_META_STORAGE_KEY = "dropoff_vehicle_meta_v1";

function createVehicleMetaStore() {
  return { next_app_id: 1, entries: {} };
}

function readVehicleMetaStore() {
  try {
    const raw = localStorage.getItem(DROP_OFF_VEHICLE_META_STORAGE_KEY);
    if (!raw) return createVehicleMetaStore();
    const parsed = JSON.parse(raw);
    return {
      next_app_id: Math.max(1, Number(parsed?.next_app_id || 1)),
      entries: typeof parsed?.entries === "object" && parsed.entries ? parsed.entries : {}
    };
  } catch (_) {
    return createVehicleMetaStore();
  }
}

function writeVehicleMetaStore(store) {
  try {
    localStorage.setItem(DROP_OFF_VEHICLE_META_STORAGE_KEY, JSON.stringify({
      next_app_id: Math.max(1, Number(store?.next_app_id || 1)),
      entries: typeof store?.entries === "object" && store.entries ? store.entries : {}
    }));
  } catch (_) {}
}

function normalizeVehicleMetaKey(value, prefix) {
  const str = String(value ?? "").trim();
  return str ? `${prefix}:${str}` : "";
}

function allocateVehicleAppId(store) {
  const nextId = Math.max(1, Number(store?.next_app_id || 1));
  store.next_app_id = nextId + 1;
  return nextId;
}

function normalizeVehicleMetaPayload(source = {}) {
  return {
    plate_number: String(source.plate_number ?? source.vehicle_id ?? source.name ?? "").trim(),
    vehicle_area: String(source.vehicle_area ?? source.area ?? "").trim(),
    home_area: String(source.home_area ?? source.home_direction ?? "").trim(),
    home_lat: source.home_lat ?? source.home_latitude ?? null,
    home_lng: source.home_lng ?? source.home_longitude ?? null,
    seat_capacity: source.seat_capacity ?? source.capacity ?? 4,
    driver_name: String(source.driver_name ?? "").trim(),
    line_id: String(source.line_id ?? "").trim(),
    status: String(source.status || "waiting").trim() || "waiting",
    memo: String(source.memo ?? "").trim()
  };
}

function findVehicleMetaEntry(store, { cloudId = null, name = "", appId = null } = {}) {
  const keys = [
    normalizeVehicleMetaKey(cloudId, "cloud"),
    normalizeVehicleMetaKey(name, "name"),
    normalizeVehicleMetaKey(appId, "app")
  ].filter(Boolean);

  for (const key of keys) {
    const entry = store?.entries?.[key];
    if (entry) return entry;
  }
  return null;
}

function upsertVehicleMetaEntry(store, source = {}, { cloudId = null, name = "", appId = null } = {}) {
  const existing = findVehicleMetaEntry(store, { cloudId, name, appId });
  const next = {
    ...(existing || {}),
    ...normalizeVehicleMetaPayload(existing || {}),
    ...normalizeVehicleMetaPayload(source || {}),
    app_id: Number(appId || existing?.app_id || allocateVehicleAppId(store)),
    cloud_row_id: cloudId || existing?.cloud_row_id || null,
    cloud_name: String(name || source?.name || existing?.cloud_name || source?.plate_number || "").trim()
  };

  const keys = new Set([
    normalizeVehicleMetaKey(next.cloud_row_id, "cloud"),
    normalizeVehicleMetaKey(next.cloud_name, "name"),
    normalizeVehicleMetaKey(next.plate_number, "name"),
    normalizeVehicleMetaKey(next.app_id, "app")
  ].filter(Boolean));

  for (const key of keys) store.entries[key] = next;
  return next;
}

function deleteVehicleMetaEntry(store, { cloudId = null, name = "", appId = null } = {}) {
  const entries = store?.entries || {};
  Object.keys(entries).forEach(key => {
    const entry = entries[key];
    if (!entry) return;
    if ((cloudId && entry.cloud_row_id === cloudId) ||
        (name && (entry.cloud_name === name || entry.plate_number === name)) ||
        (appId != null && Number(entry.app_id) === Number(appId))) {
      delete entries[key];
    }
  });
}

function normalizeVehicleRecordForApp(row = {}, meta = null) {
  const merged = {
    ...normalizeVehicleMetaPayload(meta || {}),
    ...normalizeVehicleMetaPayload(row || {})
  };

  return {
    ...row,
    id: Number(meta?.app_id || row.app_id || 0),
    cloud_row_id: row.id ?? meta?.cloud_row_id ?? null,
    db_id: row.id ?? meta?.cloud_row_id ?? null,
    plate_number: merged.plate_number || String(row.name || row.vehicle_id || row.plate_number || "").trim(),
    vehicle_area: merged.vehicle_area || "",
    home_area: merged.home_area || "",
    home_lat: merged.home_lat,
    home_lng: merged.home_lng,
    seat_capacity: Number(merged.seat_capacity || 4),
    driver_name: merged.driver_name || "",
    line_id: merged.line_id || "",
    status: merged.status || "waiting",
    memo: merged.memo || "",
    is_active: row.is_active !== false
  };
}

function buildVehicleTablePayload(basePayload = {}) {
  const payload = {
    name: String(basePayload.plate_number ?? basePayload.vehicle_id ?? basePayload.name ?? "").trim(),
    vehicle_area: String(basePayload.vehicle_area ?? basePayload.area ?? "").trim() || null,
    home_area: String(basePayload.home_area ?? basePayload.home_direction ?? "").trim() || null,
    home_lat: basePayload.home_lat ?? basePayload.home_latitude ?? null,
    home_lng: basePayload.home_lng ?? basePayload.home_longitude ?? null,
    seat_capacity: Number(basePayload.seat_capacity || basePayload.capacity || 4),
    driver_name: String(basePayload.driver_name ?? "").trim() || null,
    line_id: String(basePayload.line_id ?? "").trim() || null,
    status: String(basePayload.status || "waiting").trim() || "waiting",
    memo: String(basePayload.memo ?? "").trim() || null,
    updated_at: new Date().toISOString()
  };
  if (basePayload.team_id) payload.team_id = basePayload.team_id;
  return payload;
}

async function insertOrUpdateWithColumnFallback(tableName, mode, payload, matcherColumn, matcherValue) {
  let workingPayload = { ...(payload || {}) };
  for (let attempt = 0; attempt < 5; attempt++) {
    let query = supabaseClient.from(tableName);
    if (mode === "insert") query = query.insert(workingPayload);
    else if (mode === "upsert") query = query.upsert(workingPayload);
    else if (mode === "update") query = query.update(workingPayload).eq(matcherColumn, matcherValue);
    const { error } = await query;
    if (!error) return { error: null, payload: workingPayload };
    if (isMissingTableError(error) || !isMissingColumnError(error)) return { error, payload: workingPayload };
    const missingColumn = getMissingColumnName(error);
    if (!missingColumn || !(missingColumn in workingPayload)) return { error, payload: workingPayload };
    workingPayload = omitKeys(workingPayload, missingColumn);
  }
  return { error: new Error("column fallback failed"), payload: workingPayload };
}



async function insertSelectSingleWithColumnFallback(tableName, payload) {
  let workingPayload = { ...(payload || {}) };
  for (let attempt = 0; attempt < 6; attempt++) {
    const { data, error } = await supabaseClient
      .from(tableName)
      .insert(workingPayload)
      .select('*')
      .single();
    if (!error) return { data, error: null, payload: workingPayload };
    if (isMissingTableError(error) || !isMissingColumnError(error)) return { data: null, error, payload: workingPayload };
    const missingColumn = getMissingColumnName(error);
    if (!missingColumn || !(missingColumn in workingPayload)) return { data: null, error, payload: workingPayload };
    workingPayload = omitKeys(workingPayload, missingColumn);
  }
  return { data: null, error: new Error('column fallback failed'), payload: workingPayload };
}


function getSupabaseClientSafe() {
  try {
    if (window.supabaseClient) return window.supabaseClient;
  } catch (_) {}
  try {
    if (typeof supabaseClient !== "undefined" && supabaseClient) return supabaseClient;
  } catch (_) {}
  return null;
}


const DROP_OFF_TEAM_META_CACHE_KEY = 'dropoff_team_meta_cache_v1';

function readDropOffTeamMetaCache() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(DROP_OFF_TEAM_META_CACHE_KEY) || '{}');
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function writeDropOffTeamMetaCache(cache) {
  try {
    window.localStorage.setItem(DROP_OFF_TEAM_META_CACHE_KEY, JSON.stringify(cache && typeof cache === 'object' ? cache : {}));
  } catch (_) {}
}

function cacheDropOffTeamMeta(meta = {}) {
  const normalized = normalizeTeamRow(meta || {});
  const teamId = String(normalized?.id || '').trim();
  if (!teamId) return null;
  const cache = readDropOffTeamMetaCache();
  cache[teamId] = {
    id: teamId,
    name: String(normalized?.name || '').trim() || null,
    team_name: String(normalized?.name || '').trim() || null,
    current_origin_slot: normalized?.current_origin_slot ?? null,
    updated_at: new Date().toISOString()
  };
  writeDropOffTeamMetaCache(cache);
  try {
    window.__DROP_OFF_TEAM_META_CACHE__ = cache;
  } catch (_) {}
  return cache[teamId];
}

function getCachedDropOffTeamMeta(teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return null;
  try {
    const runtimeCache = window.__DROP_OFF_TEAM_META_CACHE__;
    if (runtimeCache && runtimeCache[safeTeamId]) {
      return normalizeTeamRow({ id: safeTeamId, ...runtimeCache[safeTeamId] });
    }
  } catch (_) {}
  const cache = readDropOffTeamMetaCache();
  return cache[safeTeamId] ? normalizeTeamRow({ id: safeTeamId, ...cache[safeTeamId] }) : null;
}

function persistResolvedWorkspaceTeamId(userId, teamId) {
  const normalized = String(teamId || '').trim();
  if (!normalized) return null;
  try {
    window.__DROP_OFF_WORKSPACE_CACHE__ = window.__DROP_OFF_WORKSPACE_CACHE__ || {};
    if (userId) window.__DROP_OFF_WORKSPACE_CACHE__[userId] = normalized;
  } catch (_) {}
  try { window.localStorage.setItem('dropoff_workspace_team_id', normalized); } catch (_) {}
  try { window.localStorage.setItem('workspaceTeamId', normalized); } catch (_) {}
  try { window.localStorage.setItem('current_dropoff_team_id', normalized); } catch (_) {}
  try { window.localStorage.setItem('__DROP_OFF_LAST_TEAM_ID__', normalized); } catch (_) {}
  try { if (userId) window.localStorage.setItem(getDropOffWorkspaceCacheKey(userId), normalized); } catch (_) {}
  try {
    window.__DROP_OFF_CURRENT_TEAM_ID__ = normalized;
    if (window.currentUserProfile && typeof window.currentUserProfile === 'object') {
      window.currentUserProfile.team_id = normalized;
      window.currentUserProfile.current_dropoff_team_id = normalized;
    }
  } catch (_) {}
  return normalized;
}

async function getCurrentAuthIdentity() {
  const directUser = currentUser || window.currentUser || null;
  if (directUser?.id) return directUser;
  try {
    const result = await supabaseClient?.auth?.getUser?.();
    return result?.data?.user || null;
  } catch (_) {
    return null;
  }
}

function getCachedWorkspaceTeamId(userId) {
  try {
    return window.localStorage.getItem(getDropOffWorkspaceCacheKey(userId));
  } catch (_) {
    return null;
  }
}

async function resolveProfileIdByEmail(email) {
  if (!email) return null;
  const tableName = getTableName('profiles');
  const { data, error } = await supabaseClient
    .from(tableName)
    .select('id,email')
    .ilike('email', String(email).trim())
    .limit(1)
    .maybeSingle();
  if (error) return null;
  return data?.id || null;
}

function getDropOffWorkspaceCacheKey(userId) {
  return userId ? `dropoff_workspace_team_id_v1_${userId}` : 'dropoff_workspace_team_id_v1';
}


function getAdminForcedWorkspaceTeamIdSafe() {
  try {
    return String(window.localStorage.getItem('admin_force_team_id') || '').trim() || null;
  } catch (_) {
    return null;
  }
}

function getSharedWorkspaceTeamCandidates() {
  const values = [];
  try { values.push(window.localStorage.getItem('dropoff_workspace_team_id')); } catch (_) {}
  try { values.push(window.localStorage.getItem('workspaceTeamId')); } catch (_) {}
  try { values.push(window.localStorage.getItem('current_dropoff_team_id')); } catch (_) {}
  try { values.push(window.localStorage.getItem('__DROP_OFF_LAST_TEAM_ID__')); } catch (_) {}
  return values.map(v => String(v || '').trim()).filter(Boolean);
}

function uniqueWorkspaceValues(values = []) {
  return Array.from(new Set((Array.isArray(values) ? values : []).map(v => String(v || '').trim()).filter(Boolean)));
}

async function workspaceTeamExists(client, teamsTable, teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return false;
  try {
    const { data, error } = await client
      .from(teamsTable)
      .select('id')
      .eq('id', safeTeamId)
      .maybeSingle();
    return !error && !!data?.id;
  } catch (_) {
    return false;
  }
}

async function findWorkspaceMembershipRows(client, teamMembersTable, identityIds) {
  const rows = [];
  for (const identityId of uniqueWorkspaceValues(identityIds)) {
    const { data, error } = await client
      .from(teamMembersTable)
      .select('team_id,user_id,role,created_at')
      .eq('user_id', identityId)
      .order('created_at', { ascending: true });
    if (error) {
      if (isMissingTableError(error)) {
        warnMissingTableOnce('team_members', error);
        return [];
      }
      continue;
    }
    if (Array.isArray(data)) rows.push(...data);
  }
  const seen = new Set();
  return rows.filter(row => {
    const key = String(row?.team_id || '').trim();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function ensureDropOffWorkspaceId() {
  const authUser = await getCurrentAuthIdentity();
  const uid = authUser?.id || null;
  const email = authUser?.email || currentUser?.email || window.currentUser?.email || '';
  const client = getSupabaseClientSafe();
  if (!uid || !client) return null;

  const isPlatformAdminRuntime = !!window.isPlatformAdminUser;
  const teamsTable = getTableName('teams');
  const teamMembersTable = getTableName('team_members');
  const currentProfile = window.currentUserProfile || null;
  const profileId = currentProfile?.id || currentProfile?.user_id || await resolveProfileIdByEmail(email);
  const identityIds = uniqueWorkspaceValues([uid, currentProfile?.id, currentProfile?.user_id, profileId]);

  const forcedTeamId = isPlatformAdminRuntime ? getAdminForcedWorkspaceTeamIdSafe() : null;
  if (forcedTeamId && await workspaceTeamExists(client, teamsTable, forcedTeamId)) {
    return persistResolvedWorkspaceTeamId(uid, forcedTeamId);
  }

  window.__DROP_OFF_WORKSPACE_CACHE__ = window.__DROP_OFF_WORKSPACE_CACHE__ || {};
  const runtimeCached = String(window.__DROP_OFF_WORKSPACE_CACHE__[uid] || '').trim() || null;
  if (runtimeCached) {
    const memberships = await findWorkspaceMembershipRows(client, teamMembersTable, identityIds);
    if (memberships.some(row => String(row?.team_id || '').trim() === runtimeCached)) {
      return persistResolvedWorkspaceTeamId(uid, runtimeCached);
    }
  }

  const candidateList = uniqueWorkspaceValues([
    getCachedWorkspaceTeamId(uid),
    currentProfile?.current_dropoff_team_id,
    window.__DROP_OFF_CURRENT_TEAM_ID__,
    window.currentWorkspaceTeamId,
    ...getSharedWorkspaceTeamCandidates()
  ]);

  if (!window.__DROP_OFF_WORKSPACE_PROMISES__) window.__DROP_OFF_WORKSPACE_PROMISES__ = {};
  if (window.__DROP_OFF_WORKSPACE_PROMISES__[uid]) return window.__DROP_OFF_WORKSPACE_PROMISES__[uid];

  window.__DROP_OFF_WORKSPACE_PROMISES__[uid] = (async () => {
    const memberships = await findWorkspaceMembershipRows(client, teamMembersTable, identityIds);

    for (const candidate of candidateList) {
      if (!await workspaceTeamExists(client, teamsTable, candidate)) continue;
      if (!memberships.some(row => String(row?.team_id || '').trim() === String(candidate).trim())) continue;
      persistResolvedWorkspaceTeamId(uid, candidate);
      return candidate;
    }

    if (memberships.length) {
      const prioritized = memberships.slice().sort((a, b) => {
        const ar = String(a?.role || 'user');
        const br = String(b?.role || 'user');
        const aw = ar === 'owner' ? 0 : ar === 'admin' ? 1 : 2;
        const bw = br === 'owner' ? 0 : br === 'admin' ? 1 : 2;
        if (aw != bw) return aw - bw;
        return String(a?.created_at || '').localeCompare(String(b?.created_at || ''));
      })[0];
      const teamId = String(prioritized?.team_id || '').trim() || null;
      if (teamId) {
        persistResolvedWorkspaceTeamId(uid, teamId);
        return teamId;
      }
    }

    return null;
  })();

  try {
    return await window.__DROP_OFF_WORKSPACE_PROMISES__[uid];
  } finally {
    delete window.__DROP_OFF_WORKSPACE_PROMISES__[uid];
  }
}

async function selectRowsPossiblyWithoutActive(tableName, logicalName, orderSpecs) {
  let useActiveFilter = !window.__DROP_OFF_NO_IS_ACTIVE_FILTER__;
  let activeOrderSpecs = Array.isArray(orderSpecs) ? [...orderSpecs] : [];

  for (let attempt = 0; attempt < 8; attempt++) {
    let query = supabaseClient.from(tableName).select("*");
    if (useActiveFilter) {
      query = query.eq("is_active", true);
    }
    for (const spec of activeOrderSpecs) {
      query = query.order(spec.column, { ascending: spec.ascending !== false });
    }

    let { data, error } = await query;
    if (!error) {
      const rows = Array.isArray(data) ? data.filter(row => row?.is_active !== false) : [];
      return { data: rows, error: null };
    }

    if (isMissingTableError(error)) {
      return { data: null, error };
    }

    if (!isMissingColumnError(error)) {
      return { data: null, error };
    }

    const missingColumn = getMissingColumnName(error);
    if (!missingColumn) {
      return { data: null, error };
    }

    if (missingColumn === "is_active" && useActiveFilter) {
      useActiveFilter = false;
      window.__DROP_OFF_NO_IS_ACTIVE_FILTER__ = true;
      continue;
    }

    const filteredSpecs = activeOrderSpecs.filter(spec => spec.column !== missingColumn);
    if (filteredSpecs.length !== activeOrderSpecs.length) {
      activeOrderSpecs = filteredSpecs;
      continue;
    }

    return { data: null, error };
  }

  return { data: [], error: null };
}


function normalizeProfileRole(role) {
  const value = String(role || '').trim().toLowerCase();
  if (value === 'owner' || value === 'admin') return value;
  return 'user';
}

function formatDisplayNameFromEmail(email) {
  const raw = String(email || '').trim();
  if (!raw) return 'ユーザー';
  return raw.split('@')[0] || raw;
}

function normalizeProfileRow(row, fallbackUser = null) {
  const authId = fallbackUser?.id || null;
  const email = String(row?.email || fallbackUser?.email || '').trim();
  const id = row?.id || row?.user_id || authId || null;
  return {
    id,
    user_id: row?.user_id || authId || id,
    email,
    display_name: String(row?.display_name || row?.name || formatDisplayNameFromEmail(email)).trim() || formatDisplayNameFromEmail(email),
    role: normalizeProfileRole(row?.role),
    is_active: row?.is_active !== false,
    created_at: row?.created_at || null,
    updated_at: row?.updated_at || null,
    team_id: row?.team_id || row?.current_dropoff_team_id || null,
    team_name: String(row?.team_name || row?.current_dropoff_team_name || row?.workspace_name || '').trim()
  };
}

async function selectSingleProfileByUser(user) {
  const tableName = getTableName('profiles');
  const attempts = [
    () => supabaseClient.from(tableName).select('*').eq('id', user.id).maybeSingle(),
    () => supabaseClient.from(tableName).select('*').eq('user_id', user.id).maybeSingle(),
    () => user.email ? supabaseClient.from(tableName).select('*').eq('email', user.email).limit(1).maybeSingle() : Promise.resolve({ data: null, error: null })
  ];

  for (const run of attempts) {
    const { data, error } = await run();
    if (!error && data) return { data, error: null };
    if (error && isMissingTableError(error)) return { data: null, error };
  }
  return { data: null, error: null };
}

async function countOwnerProfiles() {
  const tableName = getTableName('profiles');
  const { data, error } = await supabaseClient
    .from(tableName)
    .select('id,role')
    .eq('role', 'owner')
    .limit(1);

  if (error) {
    if (isMissingTableError(error)) return { count: 0, error };
    if (isMissingColumnError(error)) return { count: 0, error };
  }
  return { count: Array.isArray(data) ? data.length : 0, error: null };
}

async function upsertProfileWithFallback(payload) {
  const tableName = getTableName('profiles');
  let workingPayload = { ...(payload || {}) };
  const onConflictModes = ['id', 'user_id'];
  let lastError = null;

  for (const onConflict of onConflictModes) {
    let currentPayload = { ...workingPayload };
    for (let attempt = 0; attempt < 8; attempt++) {
      const { data, error } = await supabaseClient
        .from(tableName)
        .upsert(currentPayload, { onConflict })
        .select('*')
        .single();
      if (!error) return { data, error: null, payload: currentPayload };
      lastError = error;
      if (isMissingTableError(error)) return { data: null, error, payload: currentPayload };
      if (!isMissingColumnError(error)) break;
      const missingColumn = getMissingColumnName(error);
      if (!missingColumn || !(missingColumn in currentPayload)) break;
      currentPayload = omitKeys(currentPayload, missingColumn);
    }
  }

  return { data: null, error: lastError, payload: workingPayload };
}

async function ensureCurrentUserProfileCloud(user) {
  if (!user?.id) return null;

  const selected = await selectSingleProfileByUser(user);
  if (selected?.error && isMissingTableError(selected.error)) {
    warnMissingTableOnce('profiles', selected.error);
    const fallback = normalizeProfileRow({ id: user.id, user_id: user.id, email: user.email, display_name: formatDisplayNameFromEmail(user.email), role: 'owner', is_active: true }, user);
    window.currentUserProfile = fallback;
    return fallback;
  }

  const existing = selected?.data ? normalizeProfileRow(selected.data, user) : null;
  const ownerInfo = await countOwnerProfiles();
  const shouldBecomeOwner = ownerInfo.count === 0 || normalizeProfileRole(existing?.role) === 'owner';
  const payload = {
    id: existing?.id || user.id,
    user_id: user.id,
    email: String(existing?.email || user.email || '').trim(),
    display_name: String(existing?.display_name || user.user_metadata?.display_name || formatDisplayNameFromEmail(user.email)).trim() || formatDisplayNameFromEmail(user.email),
    role: shouldBecomeOwner ? 'owner' : normalizeProfileRole(existing?.role),
    is_active: existing?.is_active !== false,
    updated_at: new Date().toISOString()
  };
  if (!existing?.created_at) payload.created_at = new Date().toISOString();

  const result = await upsertProfileWithFallback(payload);
  if (result?.error) {
    if (isMissingTableError(result.error)) warnMissingTableOnce('profiles', result.error);
    const fallback = normalizeProfileRow({ ...(existing || {}), ...payload }, user);
    window.currentUserProfile = fallback;
    return fallback;
  }

  const profile = normalizeProfileRow(result.data || { ...(existing || {}), ...(result.payload || payload) }, user);
  window.currentUserProfile = profile;
  return profile;
}

async function loadUserProfilesCloud(teamId = null) {
  const workspaceTeamId = String(teamId || await ensureDropOffWorkspaceId() || '').trim() || null;
  if (!workspaceTeamId) return [];

  const teamMembersTable = getTableName('team_members');
  const profilesTable = getTableName('profiles');
  const teamsTable = getTableName('teams');

  const membersResult = await supabaseClient
    .from(teamMembersTable)
    .select('*')
    .eq('team_id', workspaceTeamId);

  if (membersResult.error) {
    if (isMissingTableError(membersResult.error)) {
      warnMissingTableOnce('team_members', membersResult.error);
      return [];
    }
    throw membersResult.error;
  }

  const members = Array.isArray(membersResult.data) ? membersResult.data : [];
  const memberIds = [...new Set(members.map(row => String(row?.user_id || '').trim()).filter(Boolean))];
  let teamName = String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || '').trim();
  if (!teamName) {
    try {
      const teamRes = await supabaseClient.from(teamsTable).select('*').eq('id', workspaceTeamId).maybeSingle();
      if (!teamRes.error && teamRes.data) {
        teamName = String(teamRes.data?.name || teamRes.data?.team_name || teamRes.data?.workspace_name || teamRes.data?.team_label || '').trim();
      }
    } catch (_) {}
  }

  let rows = [];
  if (memberIds.length) {
    let data = null;
    let error = null;
    ({ data, error } = await supabaseClient.from(profilesTable).select('*').in('id', memberIds));
    if (error && isMissingColumnError(error)) {
      const second = await supabaseClient.from(profilesTable).select('*').in('user_id', memberIds);
      data = second.data;
      error = second.error;
    }
    if (error) {
      if (isMissingTableError(error)) {
        warnMissingTableOnce('profiles', error);
        data = [];
        error = null;
      } else {
        throw error;
      }
    }
    rows = Array.isArray(data) ? data : [];
  }

  const profileMap = new Map(rows.map(row => [String(row?.id || row?.user_id || '').trim(), row]));
  return members
    .map(member => {
      const key = String(member?.user_id || '').trim();
      const linkedProfile = profileMap.get(key) || null;
      const profile = linkedProfile || {
        id: key,
        user_id: key,
        email: '',
        display_name: '未連携ユーザー',
        created_at: member?.created_at || null
      };
      const normalized = normalizeProfileRow(profile);
      normalized.role = normalizeProfileRole(member?.role || normalized.role);
      normalized.team_id = workspaceTeamId;
      normalized.team_name = teamName || '現在のワークスペース';
      normalized.profile_linked = Boolean(linkedProfile);
      normalized.profile_missing = !linkedProfile;
      if (normalized.profile_missing) {
        normalized.display_name = '未連携ユーザー';
        normalized.email = '';
      }
      if (!normalized.created_at) normalized.created_at = member?.created_at || null;
      return normalized;
    })
    .sort((a, b) => {
      const roleRank = { owner: 0, admin: 1, user: 2 };
      const roleDiff = (roleRank[a.role] ?? 9) - (roleRank[b.role] ?? 9);
      if (roleDiff !== 0) return roleDiff;
      return String(a.email || a.display_name || '').localeCompare(String(b.email || b.display_name || ''), 'ja', { sensitivity: 'base' });
    });
}

async function saveUserProfileCloud(profile) {
  const workspaceTeamId = String(profile?.team_id || await ensureDropOffWorkspaceId() || '').trim() || null;
  const userId = String(profile?.user_id || profile?.id || '').trim() || null;
  const email = String(profile?.email || '').trim();
  const normalizedRole = normalizeProfileRole(profile?.role);
  const now = new Date().toISOString();

  let savedMember = null;
  if (workspaceTeamId && (userId || email)) {
    const teamMembersTable = getTableName('team_members');
    let memberRow = null;

    if (userId) {
      const { data, error } = await supabaseClient
        .from(teamMembersTable)
        .select('*')
        .eq('team_id', workspaceTeamId)
        .eq('user_id', userId)
        .limit(1);
      if (error) {
        if (isMissingTableError(error)) throw error;
        throw error;
      }
      memberRow = Array.isArray(data) ? (data[0] || null) : null;
    }

    if (!memberRow && email) {
      const { data, error } = await supabaseClient
        .from(teamMembersTable)
        .select('*')
        .eq('team_id', workspaceTeamId)
        .eq('member_email', email)
        .limit(1);
      if (error) {
        if (isMissingTableError(error)) throw error;
        throw error;
      }
      memberRow = Array.isArray(data) ? (data[0] || null) : null;
    }

    if (memberRow) {
      let memberPayload = {
        role: normalizedRole,
        updated_at: now
      };

      const runMemberUpdate = async (payload) => {
        let query = supabaseClient.from(teamMembersTable).update(payload).eq('team_id', workspaceTeamId);
        if (memberRow?.user_id) query = query.eq('user_id', memberRow.user_id);
        else if (memberRow?.member_email) query = query.eq('member_email', memberRow.member_email);
        return await query.select('*').maybeSingle();
      };

      let memberResult = await runMemberUpdate(memberPayload);
      while (memberResult?.error && isMissingColumnError(memberResult.error)) {
        const missingColumn = getMissingColumnName(memberResult.error);
        if (!missingColumn || !(missingColumn in memberPayload)) break;
        memberPayload = omitKeys(memberPayload, missingColumn);
        memberResult = await runMemberUpdate(memberPayload);
      }

      if (memberResult?.error) throw memberResult.error;
      savedMember = memberResult?.data || { ...(memberRow || {}), ...(memberPayload || {}) };
    }
  }

  const payload = {
    id: profile?.id || profile?.user_id,
    user_id: profile?.user_id || profile?.id,
    email,
    display_name: String(profile?.display_name || '').trim(),
    role: normalizedRole,
    is_active: profile?.is_active !== false,
    team_id: workspaceTeamId,
    updated_at: now
  };

  const result = await upsertProfileWithFallback(payload);
  const profileWriteError = result?.error || null;
  const profileWriteFailedByPolicy = profileWriteError && (
    String(profileWriteError?.code || '').trim() === '42501' ||
    /row-level security/i.test(String(profileWriteError?.message || '')) ||
    Number(profileWriteError?.status || 0) === 401 ||
    Number(profileWriteError?.status || 0) === 403
  );

  if (profileWriteError && !profileWriteFailedByPolicy) throw profileWriteError;

  const merged = normalizeProfileRow(result?.data || { ...(profile || {}), ...(result?.payload || payload) });
  merged.role = normalizedRole;
  merged.team_id = workspaceTeamId;
  if (savedMember?.created_at && !merged.created_at) merged.created_at = savedMember.created_at;
  return merged;
}

async function deleteUserProfileCloud(profileOrId) {
  const tableName = getTableName('profiles');
  const teamMembersTable = getTableName('team_members');

  const input = (profileOrId && typeof profileOrId === 'object') ? profileOrId : { id: profileOrId, user_id: profileOrId };
  const workspaceTeamId = String(input?.team_id || await ensureDropOffWorkspaceId() || '').trim() || null;
  const profileId = String(input?.id || input?.user_id || '').trim() || null;
  const userId = String(input?.user_id || input?.id || '').trim() || null;
  const email = String(input?.email || '').trim() || null;

  let deletedMembership = false;

  if (workspaceTeamId && (userId || email)) {
    let memberDeleteQuery = supabaseClient.from(teamMembersTable).delete().eq('team_id', workspaceTeamId);
    if (userId) memberDeleteQuery = memberDeleteQuery.eq('user_id', userId);
    else if (email) memberDeleteQuery = memberDeleteQuery.eq('member_email', email);

    let memberDeleteResult = await memberDeleteQuery;
    if (memberDeleteResult?.error && userId && email) {
      memberDeleteResult = await supabaseClient
        .from(teamMembersTable)
        .delete()
        .eq('team_id', workspaceTeamId)
        .eq('member_email', email);
    }
    if (memberDeleteResult?.error) throw memberDeleteResult.error;
    deletedMembership = true;
  }

  let shouldDeleteProfile = Boolean(profileId || userId);
  if (shouldDeleteProfile && userId) {
    const remainingMembershipResult = await supabaseClient
      .from(teamMembersTable)
      .select('team_id', { count: 'exact', head: true })
      .eq('user_id', userId);
    if (remainingMembershipResult?.error && !isMissingTableError(remainingMembershipResult.error)) {
      throw remainingMembershipResult.error;
    }
    const remainingMembershipCount = Number(remainingMembershipResult?.count || 0);
    if (remainingMembershipCount > 0) shouldDeleteProfile = false;
  }

  if (shouldDeleteProfile && (profileId || userId)) {
    let profileDeleteResult = null;
    if (profileId) {
      profileDeleteResult = await supabaseClient.from(tableName).delete().eq('id', profileId);
      if (profileDeleteResult?.error && isMissingColumnError(profileDeleteResult.error)) {
        profileDeleteResult = null;
      }
    }
    if (!profileDeleteResult && userId) {
      profileDeleteResult = await supabaseClient.from(tableName).delete().eq('user_id', userId);
    }
    if (profileDeleteResult?.error && !isMissingTableError(profileDeleteResult.error)) throw profileDeleteResult.error;
  }

  return deletedMembership || shouldDeleteProfile;
}

function normalizeTeamRow(row = {}) {
  const currentOriginSlotRaw = Number(row?.current_origin_slot ?? row?.currentOriginSlot ?? row?.active_origin_slot ?? row?.selected_origin_slot);
  const resolvedName = String(
    row?.team_name
    || row?.name
    || row?.workspace_name
    || row?.team_label
    || row?.display_name
    || row?.title
    || row?.label
    || ''
  ).trim();
  const safeId = String(row?.id || row?.team_id || '').trim();
  return {
    id: safeId || null,
    name: resolvedName || (safeId ? `team-${safeId.slice(0, 8)}` : 'チーム'),
    current_origin_slot: Number.isInteger(currentOriginSlotRaw) && currentOriginSlotRaw >= 1 && currentOriginSlotRaw <= 5 ? currentOriginSlotRaw : null,
    created_at: row?.created_at || null,
    updated_at: row?.updated_at || null
  };
}

function normalizeInvitationRole(role) {
  const value = String(role || '').trim().toLowerCase();
  return value === 'admin' ? 'admin' : 'user';
}

function normalizeInvitationStatus(status, expiresAt = null) {
  const value = String(status || '').trim().toLowerCase();
  if (value === 'accepted' || value === 'revoked' || value === 'expired') return value;
  const expires = expiresAt ? new Date(expiresAt) : null;
  if (expires && !Number.isNaN(expires.getTime()) && expires.getTime() < Date.now()) return 'expired';
  return 'pending';
}

function normalizeInvitationRow(row = {}) {
  return {
    id: row?.id || null,
    email: String(row?.email || '').trim(),
    display_name: String(row?.display_name || row?.name || formatDisplayNameFromEmail(row?.email || '')).trim() || formatDisplayNameFromEmail(row?.email || ''),
    invited_role: normalizeInvitationRole(row?.invited_role || row?.role),
    team_id: row?.team_id || null,
    invited_by_user_id: row?.invited_by_user_id || row?.created_by || null,
    status: normalizeInvitationStatus(row?.status, row?.expires_at),
    expires_at: row?.expires_at || null,
    accepted_at: row?.accepted_at || null,
    created_at: row?.created_at || null,
    updated_at: row?.updated_at || null
  };
}

function generateUuidV4Safe() {
  try {
    if (typeof crypto !== 'undefined') {
      if (typeof crypto.randomUUID === 'function') return crypto.randomUUID();
      if (typeof crypto.getRandomValues === 'function') {
        const bytes = crypto.getRandomValues(new Uint8Array(16));
        bytes[6] = (bytes[6] & 0x0f) | 0x40;
        bytes[8] = (bytes[8] & 0x3f) | 0x80;
        const hex = Array.from(bytes, b => b.toString(16).padStart(2, '0'));
        return `${hex.slice(0, 4).join('')}-${hex.slice(4, 6).join('')}-${hex.slice(6, 8).join('')}-${hex.slice(8, 10).join('')}-${hex.slice(10, 16).join('')}`;
      }
    }
  } catch (_) {}

  const template = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return template.replace(/[xy]/g, ch => {
    const rnd = Math.random() * 16 | 0;
    const val = ch === 'x' ? rnd : (rnd & 0x3) | 0x8;
    return val.toString(16);
  });
}

async function loadDropOffTeamsCloud() {
  const tableName = getTableName('teams');
  if (isKnownMissingTable('teams')) return [];

  let { data, error } = await supabaseClient
    .from(tableName)
    .select('*')
    .order('created_at', { ascending: true });

  if (error && isMissingColumnError(error)) {
    const second = await supabaseClient.from(tableName).select('*');
    data = second.data;
    error = second.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('teams', error);
      return [];
    }
    throw error;
  }

  const rows = (Array.isArray(data) ? data : [])
    .map(normalizeTeamRow)
    .filter(row => row.id);

  rows.forEach(row => cacheDropOffTeamMeta(row));

  return rows.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'ja', { sensitivity: 'base' }));
}

async function getDropOffTeamMeta(teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return { data: null, error: null };

  const tableName = getTableName('teams');
  if (isKnownMissingTable('teams')) return { data: null, error: null };

  const { data, error } = await supabaseClient
    .from(tableName)
    .select('*')
    .eq('id', safeTeamId)
    .maybeSingle();

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('teams', error);
    }
    const cached = getCachedDropOffTeamMeta(safeTeamId);
    return { data: cached, error };
  }

  const normalized = data ? normalizeTeamRow(data) : getCachedDropOffTeamMeta(safeTeamId);
  if (normalized) cacheDropOffTeamMeta(normalized);
  return { data: normalized, error: null };
}

async function updateDropOffTeamCurrentOriginSlot(teamId, slotNo) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return { data: null, error: null };

  const normalizedSlot = slotNo === null || slotNo === undefined || slotNo === '' ? null : Number(slotNo);
  if (normalizedSlot !== null && (!Number.isInteger(normalizedSlot) || normalizedSlot < 1 || normalizedSlot > 5)) {
    return { data: null, error: new Error('current_origin_slot is invalid') };
  }

  const tableName = getTableName('teams');
  if (isKnownMissingTable('teams')) return { data: null, error: null };

  let payload = {
    current_origin_slot: normalizedSlot,
    updated_at: new Date().toISOString()
  };

  let { data, error } = await supabaseClient
    .from(tableName)
    .update(payload)
    .eq('id', safeTeamId)
    .select('*')
    .maybeSingle();

  while (error && isMissingColumnError(error)) {
    const missingColumn = getMissingColumnName(error);
    if (!missingColumn || !(missingColumn in payload)) break;
    payload = omitKeys(payload, missingColumn);
    const retry = await supabaseClient
      .from(tableName)
      .update(payload)
      .eq('id', safeTeamId)
      .select('*')
      .maybeSingle();
    data = retry.data;
    error = retry.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('teams', error);
    }
    const cached = getCachedDropOffTeamMeta(safeTeamId) || normalizeTeamRow({ id: safeTeamId, current_origin_slot: normalizedSlot });
    if (cached) cacheDropOffTeamMeta({ ...cached, current_origin_slot: normalizedSlot });
    return { data: cached, error };
  }

  const normalized = data ? normalizeTeamRow(data) : normalizeTeamRow({ id: safeTeamId, current_origin_slot: normalizedSlot });
  if (normalized) cacheDropOffTeamMeta(normalized);
  return { data: normalized, error: null };
}

async function loadInvitationsCloud(teamId = null) {
  const tableName = getTableName('invitations');
  if (isKnownMissingTable('invitations')) return [];
  const workspaceTeamId = String(teamId || await ensureDropOffWorkspaceId() || '').trim() || null;

  let query = supabaseClient
    .from(tableName)
    .select('*')
    .order('created_at', { ascending: false });
  if (workspaceTeamId) query = query.eq('team_id', workspaceTeamId);

  let { data, error } = await query;

  if (error && isMissingColumnError(error)) {
    const second = await supabaseClient.from(tableName).select('*');
    data = second.data;
    error = second.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('invitations', error);
      return [];
    }
    throw error;
  }

  let rows = (Array.isArray(data) ? data : []).map(normalizeInvitationRow);
  if (workspaceTeamId) rows = rows.filter(row => String(row?.team_id || '').trim() === workspaceTeamId);
  return rows;
}

async function findProfileByEmailForInvitation(email) {
  const tableName = getTableName('profiles');
  if (isKnownMissingTable('profiles') || !email) return null;

  let { data, error } = await supabaseClient
    .from(tableName)
    .select('*')
    .eq('email', email)
    .limit(1)
    .maybeSingle();

  if (error && isMissingColumnError(error)) {
    const second = await supabaseClient.from(tableName).select('*').eq('email', email);
    data = Array.isArray(second.data) ? second.data[0] : second.data;
    error = second.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('profiles', error);
      return null;
    }
    throw error;
  }

  return data ? normalizeProfileRow(data) : null;
}


async function findLatestInvitationForEmail(email) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!normalizedEmail) return null;
  const tableName = getTableName('invitations');
  if (isKnownMissingTable('invitations')) return null;

  let data = null;
  let error = null;
  ({ data, error } = await supabaseClient
    .from(tableName)
    .select('*')
    .ilike('email', normalizedEmail)
    .order('created_at', { ascending: false }));

  if (error && isMissingColumnError(error)) {
    const second = await supabaseClient.from(tableName).select('*');
    data = second.data;
    error = second.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('invitations', error);
      return null;
    }
    throw error;
  }

  const rows = (Array.isArray(data) ? data : [])
    .map(normalizeInvitationRow)
    .filter(row => String(row.email || '').trim().toLowerCase() === normalizedEmail)
    .filter(row => row.status === 'pending' || row.status === 'accepted');

  if (!rows.length) return null;
  rows.sort((a, b) => {
    const aw = a.status === 'pending' ? 0 : 1;
    const bw = b.status === 'pending' ? 0 : 1;
    if (aw != bw) return aw - bw;
    return String(b.created_at || '').localeCompare(String(a.created_at || ''));
  });
  return rows[0] || null;
}

async function findTeamMemberRow(teamId, userId, email) {
  const safeTeamId = String(teamId || '').trim();
  const safeUserId = String(userId || '').trim();
  const safeEmail = String(email || '').trim().toLowerCase();
  if (!safeTeamId) return null;
  const teamMembersTable = getTableName('team_members');

  if (safeUserId) {
    const { data, error } = await supabaseClient
      .from(teamMembersTable)
      .select('*')
      .eq('team_id', safeTeamId)
      .eq('user_id', safeUserId)
      .limit(1);
    if (error) {
      if (isMissingTableError(error)) {
        warnMissingTableOnce('team_members', error);
        return null;
      }
      throw error;
    }
    const row = Array.isArray(data) ? (data[0] || null) : null;
    if (row) return row;
  }

  if (safeEmail) {
    const { data, error } = await supabaseClient
      .from(teamMembersTable)
      .select('*')
      .eq('team_id', safeTeamId)
      .eq('member_email', safeEmail)
      .limit(1);
    if (error) {
      if (isMissingTableError(error)) {
        warnMissingTableOnce('team_members', error);
        return null;
      }
      throw error;
    }
    return Array.isArray(data) ? (data[0] || null) : null;
  }

  return null;
}

async function ensureTeamMembershipForInvitation(user, invitation) {
  const safeTeamId = String(invitation?.team_id || '').trim();
  const safeUserId = String(user?.id || '').trim();
  const safeEmail = String(user?.email || invitation?.email || '').trim().toLowerCase();
  if (!safeTeamId || !safeUserId || !safeEmail) return null;

  const teamMembersTable = getTableName('team_members');
  const normalizedRole = normalizeProfileRole(invitation?.invited_role || 'user');
  const now = new Date().toISOString();
  const existingRow = await findTeamMemberRow(safeTeamId, safeUserId, safeEmail);

  if (existingRow) {
    const preservedRole = normalizeProfileRole(existingRow?.role || normalizedRole);
    let updatePayload = {
      user_id: safeUserId,
      member_email: safeEmail,
      role: preservedRole,
      updated_at: now
    };
    const runUpdate = async (payload) => {
      let query = supabaseClient.from(teamMembersTable).update(payload).eq('team_id', safeTeamId);
      if (existingRow?.user_id) query = query.eq('user_id', existingRow.user_id);
      else if (existingRow?.member_email) query = query.eq('member_email', existingRow.member_email);
      return await query.select('*').maybeSingle();
    };
    let result = await runUpdate(updatePayload);
    while (result?.error && isMissingColumnError(result.error)) {
      const missingColumn = getMissingColumnName(result.error);
      if (!missingColumn || !(missingColumn in updatePayload)) break;
      updatePayload = omitKeys(updatePayload, missingColumn);
      result = await runUpdate(updatePayload);
    }
    if (result?.error) throw result.error;
    return result?.data || { ...(existingRow || {}), ...(updatePayload || {}) };
  }

  const insertPayload = {
    team_id: safeTeamId,
    user_id: safeUserId,
    member_email: safeEmail,
    role: normalizedRole,
    created_at: now,
    updated_at: now
  };
  const insertResult = await insertSelectSingleWithColumnFallback(teamMembersTable, insertPayload);
  if (insertResult?.error) throw insertResult.error;
  return insertResult?.data || insertResult?.payload || insertPayload;
}

async function markInvitationAcceptedById(invitationId) {
  const safeInvitationId = String(invitationId || '').trim();
  if (!safeInvitationId) return null;
  const tableName = getTableName('invitations');
  const payload = {
    status: 'accepted',
    accepted_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };

  let { data, error } = await supabaseClient
    .from(tableName)
    .update(payload)
    .eq('id', safeInvitationId)
    .select('*')
    .maybeSingle();

  if (error && isMissingColumnError(error)) {
    let workingPayload = { ...payload };
    let result = { data, error };
    while (result?.error && isMissingColumnError(result.error)) {
      const missingColumn = getMissingColumnName(result.error);
      if (!missingColumn || !(missingColumn in workingPayload)) break;
      workingPayload = omitKeys(workingPayload, missingColumn);
      result = await supabaseClient
        .from(tableName)
        .update(workingPayload)
        .eq('id', safeInvitationId)
        .select('*')
        .maybeSingle();
    }
    data = result?.data;
    error = result?.error;
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('invitations', error);
      return null;
    }
    throw error;
  }
  return normalizeInvitationRow(data || { id: safeInvitationId, ...payload });
}

async function ensureInvitedAccessForUser(user) {
  if (!user?.id) return { allowed: false, reason: 'ユーザー情報を取得できません。' };

  const safeEmail = String(user?.email || '').trim().toLowerCase();
  const profile = await ensureCurrentUserProfileCloud(user);

  let invitation = null;
  try {
    invitation = await findLatestInvitationForEmail(safeEmail);
  } catch (error) {
    console.error('findLatestInvitationForEmail failed:', error);
    throw error;
  }

  if (invitation?.team_id) {
    const membership = await ensureTeamMembershipForInvitation(user, invitation);
    const acceptedInvitation = invitation.status === 'accepted'
      ? invitation
      : await markInvitationAcceptedById(invitation.id) || invitation;

    const teamId = String(acceptedInvitation?.team_id || invitation?.team_id || '').trim();
    const invitedRole = normalizeProfileRole(membership?.role || acceptedInvitation?.invited_role || profile?.role);
    let linkedProfile = profile || null;

    if (linkedProfile && teamId) {
      const profilePayload = {
        id: linkedProfile.id || linkedProfile.user_id || user.id,
        user_id: user.id,
        email: String(user?.email || linkedProfile.email || '').trim().toLowerCase(),
        display_name: String(acceptedInvitation?.display_name || linkedProfile.display_name || formatDisplayNameFromEmail(user?.email)).trim() || formatDisplayNameFromEmail(user?.email),
        role: invitedRole,
        is_active: linkedProfile.is_active !== false,
        updated_at: new Date().toISOString(),
        team_id: teamId
      };
      if (!linkedProfile.created_at) profilePayload.created_at = new Date().toISOString();
      const profileResult = await upsertProfileWithFallback(profilePayload);
      linkedProfile = normalizeProfileRow(profileResult?.data || { ...(linkedProfile || {}), ...(profileResult?.payload || profilePayload) }, user);
    }

    if (teamId) {
      persistResolvedWorkspaceTeamId(user.id, teamId);
      try { window.currentWorkspaceTeamId = teamId; } catch (_) {}
      try {
        const meta = await getDropOffTeamMeta(teamId);
        if (meta?.data) {
          const safeTeamLabel = String(meta.data.name || '').trim();
          window.currentWorkspaceInfo = {
            ...(window.currentWorkspaceInfo || {}),
            id: teamId,
            name: safeTeamLabel || String(window.currentWorkspaceInfo?.name || '').trim(),
            team_name: safeTeamLabel || String(window.currentWorkspaceInfo?.team_name || '').trim(),
            current_origin_slot: meta.data.current_origin_slot ?? null
          };
          if (safeTeamLabel) window.__DROP_OFF_CURRENT_TEAM_LABEL__ = safeTeamLabel;
        }
      } catch (_) {}
      if (linkedProfile && typeof linkedProfile === 'object') {
        linkedProfile.team_id = teamId;
        linkedProfile.current_dropoff_team_id = teamId;
        linkedProfile.role = invitedRole;
      }
    }

    return {
      allowed: true,
      source: 'invitation',
      team_id: teamId || null,
      invitation: acceptedInvitation,
      membership,
      profile: linkedProfile || null
    };
  }

  return {
    allowed: true,
    source: 'legacy',
    profile: profile || null
  };
}

async function createInvitationCloud(payload = {}) {
  const tableName = getTableName('invitations');
  const email = String(payload?.email || '').trim().toLowerCase();
  const displayName = String(payload?.display_name || payload?.name || formatDisplayNameFromEmail(email)).trim() || formatDisplayNameFromEmail(email);
  const invitedRole = normalizeInvitationRole(payload?.invited_role || payload?.role);
  const teamId = payload?.team_id || await ensureDropOffWorkspaceId();
  const invitedByUserId = payload?.invited_by_user_id || getCurrentUserIdSafe() || null;

  if (!email) throw new Error('メールアドレスを入力してください。');
  if (!teamId) throw new Error('所属チームを選択してください。');

  const { data: teamPlan } = await getTeamPlan(teamId);
  const normalizedPlan = normalizeTeamPlanPayload({ ...(teamPlan || {}), id: teamId });
  const memberLimit = Number(normalizedPlan?.limits?.members);
  if (String(normalizedPlan?.plan_type || 'free') !== 'paid' && Number.isFinite(memberLimit) && memberLimit > 0) {
    const { data: memberUsage } = await getTeamPlanMemberUsage(teamId);
    const seatsUsed = Number(memberUsage?.seats_used || 0);
    if (seatsUsed >= memberLimit) {
      throw new Error(`freeでは利用人数は${memberLimit}名までです。不要なメンバーや招待を整理してから追加してください。`);
    }
  }

  const existingProfile = await findProfileByEmailForInvitation(email);
  if (existingProfile && existingProfile.is_active !== false) {
    throw new Error('このメールアドレスはすでに登録済みです。');
  }

  const existingInvitations = await loadInvitationsCloud();
  const pending = existingInvitations.find(row => row.email === email && row.status === 'pending');
  if (pending) {
    throw new Error('このメールアドレスには、まだ有効な招待が残っています。');
  }

  const now = new Date();
  const expiresAt = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
  const basePayload = {
    id: payload?.id || generateUuidV4Safe(),
    email,
    display_name: displayName,
    invited_role: invitedRole,
    team_id: teamId,
    invited_by_user_id: invitedByUserId,
    status: 'pending',
    expires_at: expiresAt,
    created_at: now.toISOString(),
    updated_at: now.toISOString()
  };

  const result = await insertSelectSingleWithColumnFallback(tableName, basePayload);
  if (result?.error) {
    if (isMissingTableError(result.error)) warnMissingTableOnce('invitations', result.error);
    throw result.error;
  }
  return normalizeInvitationRow(result.data || result.payload || basePayload);
}

async function updateInvitationStatusCloud(invitationId, nextStatus) {
  const tableName = getTableName('invitations');
  const status = normalizeInvitationStatus(nextStatus);
  const payload = {
    status,
    updated_at: new Date().toISOString()
  };
  if (status === 'revoked') payload.accepted_at = null;

  const { data, error } = await supabaseClient
    .from(tableName)
    .update(payload)
    .eq('id', invitationId)
    .select('*')
    .maybeSingle();

  if (error) {
    if (isMissingTableError(error)) warnMissingTableOnce('invitations', error);
    throw error;
  }
  return normalizeInvitationRow(data || { id: invitationId, ...payload });
}

async function revokeInvitationCloud(invitationId) {
  return updateInvitationStatusCloud(invitationId, 'revoked');
}

window.getTableName = getTableName;
window.getRelationName = getRelationName;
window.remapRelationSelect = remapRelationSelect;
window.isMissingTableError = isMissingTableError;
window.warnMissingTableOnce = warnMissingTableOnce;
window.ensureDropOffWorkspaceId = ensureDropOffWorkspaceId;
window.isKnownMissingTable = isKnownMissingTable;
window.isMissingColumnError = isMissingColumnError;
window.normalizeProfileRole = normalizeProfileRole;
window.ensureCurrentUserProfileCloud = ensureCurrentUserProfileCloud;
window.loadUserProfilesCloud = loadUserProfilesCloud;
window.saveUserProfileCloud = saveUserProfileCloud;
window.deleteUserProfileCloud = deleteUserProfileCloud;
window.loadDropOffTeamsCloud = loadDropOffTeamsCloud;
window.getDropOffTeamMeta = getDropOffTeamMeta;
window.updateDropOffTeamCurrentOriginSlot = updateDropOffTeamCurrentOriginSlot;
window.normalizeInvitationRole = normalizeInvitationRole;
window.normalizeInvitationStatus = normalizeInvitationStatus;
window.loadInvitationsCloud = loadInvitationsCloud;
window.createInvitationCloud = createInvitationCloud;
window.ensureInvitedAccessForUser = ensureInvitedAccessForUser;
window.revokeInvitationCloud = revokeInvitationCloud;

// save/load/import service layer

async function ensureAuth() {
  return true;
}

function getCurrentUserIdSafe() {
  return currentUser?.id || window.currentUser?.id || null;
}

function formatLocalDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function todayStr() {
  return formatLocalDate(new Date());
}

function getMonthStartStr(dateStr) {
  const d = new Date(dateStr || todayStr());
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth(), 1));
}

function getMonthEndStr(dateStr) {
  const d = new Date(dateStr || todayStr());
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth() + 1, 0));
}

async function saveCast() {
  const name = els.castName?.value.trim();
  const address = els.castAddress?.value.trim();

  if (!name) {
    alert("氏名を入力してください");
    return;
  }

  const normalizedEditingCastId = typeof normalizeDispatchEntityId === "function"
    ? normalizeDispatchEntityId(editingCastId)
    : String(editingCastId || "").trim() || null;
  const castLimitCheck = typeof canAddCast === "function"
    ? canAddCast({ isEditingExisting: Boolean(normalizedEditingCastId) })
    : { allowed: true, reason: "" };
  if (!castLimitCheck.allowed) {
    alert(castLimitCheck.reason || "このプランではこれ以上キャストを追加できません。");
    return;
  }

  const duplicate = typeof isDuplicateCast === "function" ? isDuplicateCast(name, address) : null;
  if (duplicate) {
    alert("このキャストは既に登録されています");
    return;
  }

  let lat = typeof toNullableNumber === "function" ? toNullableNumber(els.castLat?.value) : null;
  let lng = typeof toNullableNumber === "function" ? toNullableNumber(els.castLng?.value) : null;

  const addressKey = typeof normalizeGeocodeAddressKey === "function"
    ? normalizeGeocodeAddressKey(address)
    : String(address || "").trim().toLowerCase();

  if (address && typeof isValidLatLng === "function" && (!isValidLatLng(lat, lng) || addressKey !== lastCastGeocodeKey)) {
    const geocoded = typeof fillCastLatLngFromAddress === "function"
      ? await fillCastLatLngFromAddress({ silent: true, force: addressKey !== lastCastGeocodeKey })
      : null;
    lat = geocoded?.lat ?? (typeof toNullableNumber === "function" ? toNullableNumber(els.castLat?.value) : null);
    lng = geocoded?.lng ?? (typeof toNullableNumber === "function" ? toNullableNumber(els.castLng?.value) : null);
  }

  const workspaceTeamId = await ensureDropOffWorkspaceId();
  const manualArea = els.castArea?.value.trim() || "";
  const autoArea = typeof guessArea === "function" ? guessArea(lat, lng, address) : "";
  const normalizedArea = typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(manualArea || autoArea || "") : (manualArea || autoArea || "");
  const dynamicMetrics = typeof getCastOriginMetrics === "function"
    ? getCastOriginMetrics({ latitude: lat, longitude: lng, area: normalizedArea }, address)
    : null;
  const autoDistance = typeof resolveDistanceKmFromOrigin === "function"
    ? await resolveDistanceKmFromOrigin(address, lat, lng)
    : null;

  const payload = buildCastTablePayload({
    team_id: workspaceTeamId || getCurrentWorkspaceTeamIdSync?.() || null,
    name,
    phone: els.castPhone?.value.trim() || "",
    address,
    area: normalizedArea,
    distance_km: (typeof toNullableNumber === "function" ? toNullableNumber(els.castDistanceKm?.value) : null) ?? dynamicMetrics?.distance_km ?? autoDistance,
    travel_minutes: (typeof getStoredTravelMinutes === "function" ? (getStoredTravelMinutes(els.castTravelMinutes?.value) || null) : null) ?? dynamicMetrics?.travel_minutes ?? null,
    latitude: lat,
    longitude: lng,
    memo: els.castMemo?.value.trim() || "",
    is_active: true
  });

  let error;
  if (editingCastId) {
    ({ error } = await insertOrUpdateWithColumnFallback(getTableName("casts"), "update", payload, "id", editingCastId));
  } else {
    payload.created_by = getCurrentUserIdSafe();
    ({ error } = await insertOrUpdateWithColumnFallback(getTableName("casts"), "insert", payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, editingCastId ? "update_cast" : "create_cast", editingCastId ? "キャストを更新" : "キャストを作成");
  if (typeof resetCastForm === "function") resetCastForm();
  await loadCasts();
}

async function cleanupInactiveCastsForTeam(teamId, options = {}) {
  const safeTeamId = String(teamId || "").trim();
  if (!safeTeamId) return { deleted: 0, error: null };

  try {
    const { error } = await supabaseClient
      .from(getTableName("casts"))
      .delete()
      .eq("team_id", safeTeamId)
      .eq("is_active", false);

    if (!error) {
      return { deleted: 0, error: null };
    }
    if (isMissingColumnError(error) && /is_active/i.test(String(error?.message || ""))) {
      return { deleted: 0, error: null };
    }
    if (!options.silent) console.error("cleanupInactiveCastsForTeam error:", error);
    return { deleted: 0, error };
  } catch (error) {
    if (!options.silent) console.error("cleanupInactiveCastsForTeam exception:", error);
    return { deleted: 0, error };
  }
}

async function deleteCast(castId) {
  if (!window.confirm("このキャストを削除しますか？")) return;

  const workspaceTeamId = await ensureDropOffWorkspaceId();
  let query = supabaseClient
    .from(getTableName("casts"))
    .delete()
    .eq("id", castId);
  if (workspaceTeamId) {
    query = query.eq("team_id", workspaceTeamId);
  }

  let { error } = await query;
  if (error) {
    alert(error.message);
    return;
  }

  if (workspaceTeamId) {
    await cleanupInactiveCastsForTeam(workspaceTeamId, { silent: true });
  }

  await addHistory(null, null, "delete_cast", `キャストID ${castId} を削除`);
  await loadCasts();
}

function canUseCsvFeatureForCurrentPlan() {
  const plan = typeof getCurrentPlanRecord === "function" ? getCurrentPlanRecord() : null;
  const isPaid = String(plan?.plan_type || "free").trim() === "paid";
  const csvFlag = plan?.feature_flags?.csv;
  return isPaid && csvFlag !== false;
}

function isImportCountLimitBypassedForCurrentPlan() {
  if (window.isPlatformAdminUser === true) return true;
  const plan = typeof getCurrentPlanRecord === "function" ? getCurrentPlanRecord() : null;
  return String(plan?.plan_type || "free").trim() === "paid";
}

function ensureCsvFeatureAccessForCurrentPlan(actionLabel = "CSV機能") {
  if (canUseCsvFeatureForCurrentPlan()) return true;
  const plan = typeof getCurrentPlanRecord === "function" ? getCurrentPlanRecord() : null;
  const planLabel = typeof getPlanTypeLabel === "function" ? getPlanTypeLabel(plan?.plan_type) : "現在のプラン";
  alert(`${planLabel}では${actionLabel}は利用できません。Paidで利用できます。`);
  return false;
}

async function importCastCsvFile() {
  if (!ensureCsvFeatureAccessForCurrentPlan("キャストCSVインポート")) {
    if (els.csvFileInput) els.csvFileInput.value = "";
    return;
  }
  const file = els.csvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const workspaceTeamId = await ensureDropOffWorkspaceId();
    if (!workspaceTeamId) {
      alert("現在のワークスペース(team)を特定できないため、キャストCSVを取り込めません。いったんログアウトして再ログインしてください。");
      return;
    }

    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const normalizeImportKey = (name, address) => `${String(name || "").trim().toLowerCase()}__${String(address || "").trim().toLowerCase()}`;
    const uniqueMap = new Map();

    for (const row of rows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const key = normalizeImportKey(name, address);
      if (!uniqueMap.has(key)) uniqueMap.set(key, row);
    }

    const mergedRows = [...uniqueMap.values()];
    const payloads = [];

    for (const row of mergedRows) {
      const name = String(row.name || "").trim();
      const address = String(row.address || "").trim();
      if (!name || !address) continue;

      const lat = typeof toNullableNumber === "function" ? toNullableNumber(row.latitude) : null;
      const lng = typeof toNullableNumber === "function" ? toNullableNumber(row.longitude) : null;
      const autoArea = typeof guessArea === "function" ? guessArea(lat, lng, address) : "";

      payloads.push(buildCastTablePayload({
        team_id: workspaceTeamId,
        name,
        phone: String(row.phone || "").trim(),
        address,
        area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.area || "").trim() || autoArea || "") : (String(row.area || "").trim() || autoArea || ""),
        distance_km:
          (typeof toNullableNumber === "function" ? toNullableNumber(row.distance_km) : null) ??
          ((typeof isValidLatLng === "function" && isValidLatLng(lat, lng) && typeof estimateRoadKmFromStation === "function") ? estimateRoadKmFromStation(lat, lng) : null),
        travel_minutes: typeof getStoredTravelMinutes === "function" ? (getStoredTravelMinutes(row.travel_minutes) || null) : null,
        latitude: lat,
        longitude: lng,
        memo: String(row.memo || "").trim(),
        is_active: true,
        created_by: getCurrentUserIdSafe()
      }));
    }

    if (!payloads.length) {
      alert("取り込めるデータがありません");
      els.csvFileInput.value = "";
      return;
    }

    const existingResult = await selectRowsClientSideSafe(
      getTableName("casts"),
      "casts",
      [{ column: "id", ascending: true }],
      { teamId: workspaceTeamId }
    );

    if (existingResult?.error) {
      console.error("CSV import existing cast load error:", existingResult.error);
      alert("既存キャストの確認に失敗しました: " + existingResult.error.message);
      return;
    }

    const existingByKey = new Map();
    for (const row of (existingResult?.data || [])) {
      const key = normalizeImportKey(row?.name, row?.address);
      if (!key || existingByKey.has(key)) continue;
      existingByKey.set(key, row);
    }

    const currentCastCount = Array.isArray(allCastsCache) ? allCastsCache.length : ((existingResult?.data || []).length);
    const newCastCount = payloads.reduce((count, payload) => {
      const key = normalizeImportKey(payload.name, payload.address);
      return count + (existingByKey.has(key) ? 0 : 1);
    }, 0);
    const castLimit = typeof getPlanLimit === "function" ? Number(getPlanLimit("casts")) : NaN;
    if (!isImportCountLimitBypassedForCurrentPlan() && Number.isFinite(castLimit) && castLimit > 0 && currentCastCount + newCastCount > castLimit) {
      const planLabel = typeof getPlanTypeLabel === "function" ? getPlanTypeLabel(getCurrentPlanRecord?.()?.plan_type) : "現在のプラン";
      alert(`${planLabel}ではキャストは${castLimit}件までです。CSV取込前に不要なデータを削除してください。`);
      els.csvFileInput.value = "";
      return;
    }

    const tableName = getTableName("casts");
    let savedCount = 0;

    for (const payload of payloads) {
      const key = normalizeImportKey(payload.name, payload.address);
      const existing = existingByKey.get(key);
      let result = null;

      if (existing?.id) {
        const updatePayload = omitKeys(payload, ["created_by"]);
        result = await insertOrUpdateWithColumnFallback(tableName, "update", updatePayload, "id", existing.id);
      } else {
        result = await insertSelectSingleWithColumnFallback(tableName, payload);
      }

      if (result?.error) {
        console.error("CSV import supabase error:", result.error, payload);
        alert(`CSV取込エラー: ${payload.name} / ${payload.address} -> ${result.error.message}`);
        return;
      }

      savedCount += 1;
      const savedId = result?.data?.id || existing?.id || null;
      if (savedId && !existingByKey.has(key)) {
        existingByKey.set(key, { id: savedId, name: payload.name, address: payload.address, team_id: workspaceTeamId });
      }
    }

    els.csvFileInput.value = "";
    await addHistory(null, null, "import_csv", `${savedCount}件のキャストをCSV取込/更新`);
    alert(`${savedCount}件のキャストをCSV取込/更新しました`);
    await loadCasts();
  } catch (error) {
    console.error("importCastCsvFile error:", error);
    alert("CSV取込中にエラーが発生しました");
  }
}

function parseNullableCoordinate(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

async function saveVehicle() {
  const plateNumber = els.vehiclePlateNumber?.value.trim();
  if (!plateNumber) {
    alert("車両IDを入力してください");
    return;
  }

  const editingVehicleKey = String(editingVehicleId || "").trim() || null;
  const vehicleLimitCheck = typeof canAddVehicle === "function"
    ? canAddVehicle({ isEditingExisting: Boolean(editingVehicleKey) })
    : { allowed: true, reason: "" };
  if (!vehicleLimitCheck.allowed) {
    alert(vehicleLimitCheck.reason || "このプランではこれ以上車両を追加できません。");
    return;
  }

  const duplicate = typeof isDuplicateVehicle === "function" ? isDuplicateVehicle(plateNumber) : null;
  if (duplicate) {
    alert("この車両IDは既に登録されています");
    return;
  }

  const detailPayload = normalizeVehicleMetaPayload({
    plate_number: plateNumber,
    vehicle_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(els.vehicleArea?.value.trim() || "") : (els.vehicleArea?.value.trim() || ""),
    home_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(els.vehicleHomeArea?.value.trim() || "") : (els.vehicleHomeArea?.value.trim() || ""),
    home_lat: parseNullableCoordinate(els.vehicleHomeLat?.value),
    home_lng: parseNullableCoordinate(els.vehicleHomeLng?.value),
    seat_capacity: Number(els.vehicleSeatCapacity?.value || 4),
    driver_name: els.vehicleDriverName?.value.trim() || "",
    line_id: els.vehicleLineId?.value.trim() || "",
    status: els.vehicleStatus?.value || "waiting",
    memo: els.vehicleMemo?.value.trim() || ""
  });

  const tableName = getTableName("vehicles");
  const editingVehicle = Array.isArray(allVehiclesCache)
    ? allVehiclesCache.find(v => String(v.id) === String(editingVehicleId))
    : null;
  const cloudVehicleId = editingVehicle?.cloud_row_id || editingVehicle?.db_id || null;
  const workspaceTeamId = await ensureDropOffWorkspaceId();
  if (!workspaceTeamId) {
    alert("現在のワークスペース(team)を特定できないため、車両を保存できません。いったんログアウトして再ログインしてください。");
    return;
  }
  const cloudPayload = buildVehicleTablePayload({ plate_number: plateNumber, team_id: workspaceTeamId, ...detailPayload });

  let savedRow = null;
  let error = null;

  if (cloudVehicleId) {
    const updateResult = await insertOrUpdateWithColumnFallback(tableName, "update", cloudPayload, "id", cloudVehicleId);
    error = updateResult.error;
    if (!error) {
      const reloadResult = await supabaseClient
        .from(tableName)
        .select("*")
        .eq("id", cloudVehicleId)
        .single();
      error = reloadResult.error;
      savedRow = reloadResult.data || null;
    }
  } else {
    const insertResult = await insertSelectSingleWithColumnFallback(tableName, cloudPayload);
    error = insertResult.error;
    savedRow = insertResult.data || null;
  }

  if (error) {
    alert(error.message);
    return;
  }

  const store = readVehicleMetaStore();
  const mergedSource = {
    ...detailPayload,
    ...(savedRow || {}),
    plate_number: plateNumber
  };
  upsertVehicleMetaEntry(store, mergedSource, {
    cloudId: savedRow?.id || cloudVehicleId || null,
    name: plateNumber,
    appId: editingVehicle?.id || null
  });
  writeVehicleMetaStore(store);

  await addHistory(null, null, editingVehicleId ? "update_vehicle" : "create_vehicle", editingVehicleId ? "車両を更新" : "車両を登録");
  if (typeof resetVehicleForm === "function") resetVehicleForm();
  await loadVehicles();
}

async function deleteVehicle(vehicleId) {
  if (!window.confirm("この車両を削除しますか？")) return;

  const targetVehicle = Array.isArray(allVehiclesCache)
    ? allVehiclesCache.find(v => String(v.id) === String(vehicleId))
    : null;
  const cloudVehicleId = targetVehicle?.cloud_row_id || targetVehicle?.db_id || null;

  let error = null;
  if (cloudVehicleId) {
    ({ error } = await supabaseClient
      .from(getTableName("vehicles"))
      .delete()
      .eq("id", cloudVehicleId));
  }

  if (error) {
    alert(error.message);
    return;
  }

  const store = readVehicleMetaStore();
  deleteVehicleMetaEntry(store, {
    cloudId: cloudVehicleId,
    name: targetVehicle?.plate_number || "",
    appId: vehicleId
  });
  writeVehicleMetaStore(store);

  await addHistory(null, null, "delete_vehicle", `車両ID ${vehicleId} を削除`);
  await loadVehicles();
}

async function importVehicleCsvFile() {
  if (!ensureCsvFeatureAccessForCurrentPlan("車両CSVインポート")) {
    if (els.vehicleCsvFileInput) els.vehicleCsvFileInput.value = "";
    return;
  }
  const file = els.vehicleCsvFileInput?.files?.[0];
  if (!file) {
    alert("CSVファイルを選択してください");
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVデータが空です");
      return;
    }

    const inserts = [];

    for (const row of rows) {
      const plateNumber = String(row.plate_number || "").trim();
      if (!plateNumber) continue;

      const exists = Array.isArray(allVehiclesCache) ? allVehiclesCache.find(
        v => String(v.plate_number || "").trim() === plateNumber
      ) : null;
      if (exists) {
        console.log("車両重複スキップ:", plateNumber);
        continue;
      }

      const detail = normalizeVehicleMetaPayload({
        plate_number: plateNumber,
        vehicle_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.vehicle_area || "").trim() || "") : String(row.vehicle_area || "").trim(),
        home_area: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel(String(row.home_area || "").trim() || "") : String(row.home_area || "").trim(),
        home_lat: parseNullableCoordinate(row.home_lat),
        home_lng: parseNullableCoordinate(row.home_lng),
        seat_capacity: Number(row.seat_capacity || 4),
        driver_name: String(row.driver_name || "").trim(),
        line_id: String(row.line_id || "").trim(),
        status: String(row.status || "waiting").trim() || "waiting",
        memo: String(row.memo || "").trim()
      });

      inserts.push({
        cloud: buildVehicleTablePayload({ plate_number: plateNumber, ...detail }),
        detail,
        plate_number: plateNumber
      });
    }

    if (!inserts.length) {
      alert("新規車両はありません");
      els.vehicleCsvFileInput.value = "";
      return;
    }

    const currentVehicleCount = Array.isArray(allVehiclesCache) ? allVehiclesCache.length : 0;
    const vehicleLimit = typeof getPlanLimit === "function" ? Number(getPlanLimit("vehicles")) : NaN;
    if (!isImportCountLimitBypassedForCurrentPlan() && Number.isFinite(vehicleLimit) && vehicleLimit > 0 && currentVehicleCount + inserts.length > vehicleLimit) {
      const planLabel = typeof getPlanTypeLabel === "function" ? getPlanTypeLabel(getCurrentPlanRecord?.()?.plan_type) : "現在のプラン";
      alert(`${planLabel}では車両は${vehicleLimit}件までです。CSV取込前に不要なデータを削除してください。`);
      els.vehicleCsvFileInput.value = "";
      return;
    }

    const store = readVehicleMetaStore();
    let savedCount = 0;

    for (const row of inserts) {
      const result = await insertSelectSingleWithColumnFallback(getTableName("vehicles"), row.cloud);
      if (result.error) {
        console.error("Vehicle CSV import error:", result.error);
        alert("車両CSV取込エラー: " + result.error.message);
        return;
      }
      savedCount += 1;
      upsertVehicleMetaEntry(store, {
        ...row.detail,
        ...(result.data || {}),
        plate_number: row.plate_number
      }, {
        cloudId: result.data?.id || null,
        name: row.plate_number,
        appId: null
      });
    }

    writeVehicleMetaStore(store);

    els.vehicleCsvFileInput.value = "";
    await addHistory(null, null, "import_vehicle_csv", `${savedCount}件の車両をCSV取込`);
    alert(`${savedCount}件の車両を取り込みました`);
    await loadVehicles();
  } catch (error) {
    console.error("importVehicleCsvFile error:", error);
    alert("車両CSV取込中にエラーが発生しました");
  }
}

function getVehicleDailyRunsTableNameForService() {
  return String(window?.DROP_OFF_TABLES?.vehicle_daily_runs || "dropoff_vehicle_daily_runs");
}

async function resolveMileageWorkspaceTeamIdSafe() {
  try {
    if (typeof resolveWorkspaceTeamIdForDailyRuns === "function") {
      const resolved = await resolveWorkspaceTeamIdForDailyRuns();
      if (String(resolved || "").trim()) return String(resolved).trim();
    }
  } catch (_) {}

  try {
    if (typeof ensureDropOffWorkspaceId === "function") {
      const resolved = await ensureDropOffWorkspaceId();
      if (String(resolved || "").trim()) return String(resolved).trim();
    }
  } catch (_) {}

  return String(window?.currentWorkspaceTeamId || "").trim() || null;
}

function resolveMileageVehicleRecord(rawVehicleId, fallbackDriverName = "") {
  const direct = String(rawVehicleId || "").trim();
  const safeDriver = String(fallbackDriverName || "").trim();

  const localVehicleId = typeof resolveVehicleLocalNumericId === "function"
    ? Number(resolveVehicleLocalNumericId(rawVehicleId) || 0)
    : Number(rawVehicleId || 0);

  let matchedVehicle = null;
  if (localVehicleId > 0 && Array.isArray(allVehiclesCache)) {
    matchedVehicle = allVehiclesCache.find(vehicle => Number(vehicle?.id || 0) === localVehicleId) || null;
  }

  if (!matchedVehicle && direct && Array.isArray(allVehiclesCache)) {
    matchedVehicle = allVehiclesCache.find(vehicle => {
      const localId = String(vehicle?.id || "").trim();
      const cloudId = String(vehicle?.cloud_row_id || vehicle?.db_id || "").trim();
      return direct === localId || (cloudId && cloudId === direct);
    }) || null;
  }

  if (!matchedVehicle && safeDriver && Array.isArray(allVehiclesCache)) {
    matchedVehicle = allVehiclesCache.find(vehicle => String(vehicle?.driver_name || "").trim() === safeDriver) || null;
  }

  return {
    localVehicleId: Number(matchedVehicle?.id || localVehicleId || 0),
    plateNumber: String(matchedVehicle?.plate_number || "").trim() || "-",
    driverName: String(matchedVehicle?.driver_name || safeDriver || "").trim() || "-"
  };
}

function normalizeDailyRunRowsToLegacyMileageRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row, index) => {
    const vehicleMeta = resolveMileageVehicleRecord(row?.vehicle_id, row?.driver_name);
    return {
      id: row?.id || `dailyrun:${index}`,
      team_id: row?.team_id || null,
      vehicle_id: Number(vehicleMeta.localVehicleId || 0),
      raw_vehicle_id: String(row?.vehicle_id || "").trim() || null,
      report_date: row?.run_date || row?.report_date || "",
      distance_km: Number(row?.reference_distance_km ?? row?.distance_km ?? 0),
      worked_flag: Number(row?.is_workday === false ? 0 : 1),
      note: row?.note || "",
      driver_name: vehicleMeta.driverName,
      plate_number: vehicleMeta.plateNumber,
      vehicles: {
        id: Number(vehicleMeta.localVehicleId || 0) || null,
        plate_number: vehicleMeta.plateNumber,
        driver_name: vehicleMeta.driverName
      }
    };
  });
}

async function fetchDriverMileageRowsFromDailyRuns(startDate, endDate) {
  const tableName = getVehicleDailyRunsTableNameForService();
  const workspaceTeamId = await resolveMileageWorkspaceTeamIdSafe();

  const baseSelectColumns = [
    "id",
    "team_id",
    "vehicle_id",
    "run_date",
    "reference_distance_km",
    "is_workday"
  ];

  let query = supabaseClient
    .from(tableName)
    .select(baseSelectColumns.join(", "))
    .gte("run_date", startDate)
    .lte("run_date", endDate)
    .order("run_date", { ascending: true })
    .order("vehicle_id", { ascending: true })
    .order("id", { ascending: true });

  if (workspaceTeamId) {
    query = query.eq("team_id", workspaceTeamId);
  }

  let { data, error } = await query;

  if (error && workspaceTeamId && typeof isMissingColumnError === "function" && isMissingColumnError(error) && /team_id/i.test(String(error?.message || ""))) {
    ({ data, error } = await supabaseClient
      .from(tableName)
      .select(baseSelectColumns.filter(col => col !== "team_id").join(", "))
      .gte("run_date", startDate)
      .lte("run_date", endDate)
      .order("run_date", { ascending: true })
      .order("vehicle_id", { ascending: true })
      .order("id", { ascending: true }));
  }

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("vehicle_daily_runs", error);
      return { rows: [], usedDailyRuns: false, missingTable: true };
    }
    console.error(error);
    alert("走行実績の取得に失敗しました: " + error.message);
    return { rows: [], usedDailyRuns: true, missingTable: false, fatal: true };
  }

  return {
    rows: normalizeDailyRunRowsToLegacyMileageRows(data || []),
    usedDailyRuns: true,
    missingTable: false,
    fatal: false
  };
}

async function fetchDriverMileageRows(startDate, endDate) {
  const dailyRunsResult = await fetchDriverMileageRowsFromDailyRuns(startDate, endDate);
  if (dailyRunsResult.usedDailyRuns || dailyRunsResult.fatal) {
    return dailyRunsResult.rows || [];
  }

  if (isKnownMissingTable("vehicle_daily_reports")) {
    return [];
  }

  const { data, error } = await supabaseClient
    .from(getTableName("vehicle_daily_reports"))
    .select(remapRelationSelect(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `))
    .gte("report_date", startDate)
    .lte("report_date", endDate)
    .order("report_date", { ascending: true })
    .order("vehicle_id", { ascending: true });

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("vehicle_daily_reports", error);
      return [];
    }
    console.error(error);
    alert("走行実績の取得に失敗しました: " + error.message);
    return [];
  }

  return data || [];
}

async function loadCasts() {
  const workspaceTeamId = await ensureDropOffWorkspaceId();
  if (!workspaceTeamId) {
    allCastsCache = [];
    if (typeof renderCastsTable === "function") renderCastsTable();
    if (typeof renderCastSearchResults === "function") renderCastSearchResults();
    if (typeof renderCastSelects === "function") renderCastSelects();
    if (typeof renderHomeSummary === "function") renderHomeSummary();
    if (typeof renderPlanInfo === "function") renderPlanInfo();
    return;
  }

  if (!window.__DROP_OFF_CAST_CLEANUP_DONE__) window.__DROP_OFF_CAST_CLEANUP_DONE__ = {};
  if (!window.__DROP_OFF_CAST_CLEANUP_DONE__[workspaceTeamId]) {
    const cleanupResult = await cleanupInactiveCastsForTeam(workspaceTeamId, { silent: true });
    if (!cleanupResult?.error) {
      window.__DROP_OFF_CAST_CLEANUP_DONE__[workspaceTeamId] = true;
    }
  }

  const { data, error } = await selectRowsClientSideSafe(
    getTableName("casts"),
    "casts",
    [
      { column: "name", ascending: true },
      { column: "id", ascending: true }
    ],
    { teamId: workspaceTeamId }
  );

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("casts", error);
      allCastsCache = [];
      if (typeof renderCastsTable === "function") renderCastsTable();
      if (typeof renderCastSearchResults === "function") renderCastSearchResults();
      if (typeof renderCastSelects === "function") renderCastSelects();
      if (typeof renderHomeSummary === "function") renderHomeSummary();
      return;
    }
    console.error(error);
    return;
  }

  allCastsCache = (data || []).map(normalizeCastRecordForApp);
  if (typeof refreshCastMetricsForCurrentOrigin === "function") {
    try {
      await refreshCastMetricsForCurrentOrigin({ render: false });
    } catch (error) {
      console.error('refreshCastMetricsForCurrentOrigin error:', error);
    }
  }
  if (typeof renderCastsTable === "function") renderCastsTable();
  if (typeof renderCastSearchResults === "function") renderCastSearchResults();
  if (typeof renderCastSelects === "function") renderCastSelects();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
  if (typeof renderPlanInfo === "function") renderPlanInfo();
}

async function loadVehicles() {
  const workspaceTeamId = await ensureDropOffWorkspaceId();
  if (!workspaceTeamId) {
    allVehiclesCache = [];
    if (typeof renderVehiclesTable === "function") renderVehiclesTable();
    if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
    if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
    if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
    if (typeof renderHomeSummary === "function") renderHomeSummary();
    if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
    else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
    if (typeof renderPlanInfo === "function") renderPlanInfo();
    return;
  }

  const { data, error } = await supabaseClient
    .from(getTableName("vehicles"))
    .select('*')
    .eq('team_id', workspaceTeamId)
    .order('name', { ascending: true });

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("vehicles", error);
      allVehiclesCache = [];
      if (typeof renderVehiclesTable === "function") renderVehiclesTable();
      if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
      if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
      if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
      if (typeof renderHomeSummary === "function") renderHomeSummary();
      if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
      else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
      if (typeof renderPlanInfo === "function") renderPlanInfo();
      return;
    }
    console.error(error);
    return;
  }

  const store = readVehicleMetaStore();
  allVehiclesCache = (data || []).map(row => {
    const cloudName = String(row?.name || row?.vehicle_id || row?.plate_number || "").trim();
    const meta = upsertVehicleMetaEntry(store, row, {
      cloudId: row?.id || null,
      name: cloudName
    });
    return normalizeVehicleRecordForApp(row, meta);
  });
  writeVehicleMetaStore(store);

  const validIds = new Set(allVehiclesCache.map(v => Number(v.id || 0)).filter(Boolean));
  activeVehicleIdsForToday = new Set(
    [...(activeVehicleIdsForToday || new Set())].filter(id => validIds.has(Number(id || 0)))
  );

  if (typeof renderVehiclesTable === "function") renderVehiclesTable();
  if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
  if (typeof renderPlanInfo === "function") renderPlanInfo();
}

async function loadDailyReports(dateStr) {
  const start = getMonthStartStr(dateStr || todayStr());
  const end = getMonthEndStr(dateStr || todayStr());

  const dailyRunsResult = await fetchDriverMileageRowsFromDailyRuns(start, end);
  if (dailyRunsResult.usedDailyRuns || dailyRunsResult.fatal) {
    currentDailyReportsCache = dailyRunsResult.rows || [];
    if (typeof renderVehiclesTable === "function") renderVehiclesTable();
    if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
    else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
    if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
    if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
    return;
  }

  if (isKnownMissingTable("vehicle_daily_reports")) {
    currentDailyReportsCache = [];
    if (typeof renderVehiclesTable === "function") renderVehiclesTable();
    if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
    else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
    if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
    if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
    return;
  }

  const { data, error } = await supabaseClient
    .from(getTableName("vehicle_daily_reports"))
    .select(remapRelationSelect(`
      *,
      vehicles (
        id,
        plate_number,
        driver_name
      )
    `))
    .gte("report_date", start)
    .lte("report_date", end)
    .order("report_date", { ascending: true })
    .order("vehicle_id", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("vehicle_daily_reports", error);
      currentDailyReportsCache = [];
      if (typeof renderVehiclesTable === "function") renderVehiclesTable();
      if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
      else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
      if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
      if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
      return;
    }
    console.error(error);
    return;
  }

  currentDailyReportsCache = data || [];
  if (typeof renderVehiclesTable === "function") renderVehiclesTable();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
}

async function loadHistory() {
  if (isKnownMissingTable("dispatch_history")) {
      if (els?.historyList) els.historyList.innerHTML = `<div class="muted">履歴テーブル未設定</div>`;
      return;
  }
  const { data, error } = await supabaseClient
    .from(getTableName("dispatch_history"))
    .select("*")
    .order("created_at", { ascending: false })
    .limit(100);

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("dispatch_history", error);
      if (els?.historyList) els.historyList.innerHTML = `<div class="muted">履歴テーブル未設定</div>`;
      return;
    }
    console.error(error);
    return;
  }

  if (!els?.historyList) return;
  els.historyList.innerHTML = "";

  if (!data?.length) {
    els.historyList.innerHTML = `<div class="muted">履歴はありません</div>`;
    return;
  }

  data.forEach(row => {
    const div = document.createElement("div");
    div.className = "history-item";
    div.innerHTML = `
      <h4>${escapeHtml(row.action)}</h4>
      <p>${escapeHtml(row.message || "")}</p>
      <p class="muted">${escapeHtml(formatDateTimeJa(row.created_at))}</p>
    `;
    els.historyList.appendChild(div);
  });
}

async function loadHomeAndAll() {
  const dateStr = els?.dispatchDate?.value || todayStr();

  if (els?.dispatchDate) els.dispatchDate.value = dateStr;
  if (els?.planDate) els.planDate.value = dateStr;
  if (els?.actualDate) els.actualDate.value = dateStr;
  if (typeof syncMileageReportRange === "function") {
    syncMileageReportRange(dateStr, true);
  } else {
    if (els?.mileageReportStartDate) els.mileageReportStartDate.value = getMonthStartStr(dateStr);
    if (els?.mileageReportEndDate) els.mileageReportEndDate.value = dateStr;
  }

  await loadCasts();
  await loadVehicles();
  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  await loadHistory();

  if (typeof renderDailyVehicleChecklist === "function") renderDailyVehicleChecklist();
  if (typeof renderDailyMileageInputs === "function") renderDailyMileageInputs();
  if (typeof renderDailyDispatchResult === "function") renderDailyDispatchResult();
  if (typeof renderHomeSummary === "function") renderHomeSummary();
  if (typeof refreshHomeMonthlyVehicleList === "function") await refreshHomeMonthlyVehicleList();
  else if (typeof renderHomeMonthlyVehicleList === "function") renderHomeMonthlyVehicleList();
}

async function loadAllData() {
  return loadHomeAndAll();
}



async function loadPlansByDate(dateStr) {
  if (isKnownMissingTable("dispatch_plans")) {
      currentPlansCache = [];
      renderPlanGroupedTable();
      renderPlansTimeAreaMatrix();
      renderPlanSelect();
      renderPlanCastSelect();
      renderHomeSummary();
      renderOperationAndSimulationUI();
      return;
  }
  const { data, error } = await supabaseClient
    .from(getTableName("dispatch_plans"))
    .select(remapRelationSelect(`
      *,
      casts (
        id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude
      )
    `))
    .eq("plan_date", dateStr)
    .order("plan_hour", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("dispatch_plans", error);
      currentPlansCache = [];
      renderPlanGroupedTable();
      renderPlansTimeAreaMatrix();
      renderPlanSelect();
      renderPlanCastSelect();
      renderHomeSummary();
      renderOperationAndSimulationUI();
      return;
    }
    console.error(error);
    return;
  }

  currentPlansCache = data || [];
  renderPlanGroupedTable();
  renderPlansTimeAreaMatrix();
  renderPlanSelect();
  renderPlanCastSelect();
  renderHomeSummary();
  renderOperationAndSimulationUI();
}

async function loadActualsByDate(dateStr) {
  if (isKnownMissingTable("dispatches")) {
      currentDispatchId = null;
      currentActualsCache = [];
      renderActualTable();
      renderActualTimeAreaMatrix();
      renderHomeSummary();
      renderCastSelects();
      renderManualLastVehicleInfo();
      return;
  }
  const { data: dispatches, error: dispatchError } = await supabaseClient
    .from(getTableName("dispatches"))
    .select("*")
    .eq("dispatch_date", dateStr)
    .order("id", { ascending: false })
    .limit(1);

  if (dispatchError) {
    if (isMissingTableError(dispatchError)) {
      warnMissingTableOnce("dispatches", dispatchError);
      currentDispatchId = null;
      currentActualsCache = [];
      renderActualTable();
      renderActualTimeAreaMatrix();
      renderHomeSummary();
      renderCastSelects();
      renderManualLastVehicleInfo();
      return;
    }
    console.error(dispatchError);
    return;
  }

  if (dispatches?.length) {
    currentDispatchId = dispatches[0].id;
  } else {
    const { data: inserted, error: createError } = await supabaseClient
      .from(getTableName("dispatches"))
      .insert({
        dispatch_date: dateStr,
        status: "draft",
        created_by: getCurrentUserIdSafe()
      })
      .select()
      .single();

    if (createError) {
      if (isMissingTableError(createError)) {
        warnMissingTableOnce("dispatches", createError);
        currentDispatchId = null;
        currentActualsCache = [];
        renderActualTable();
        renderActualTimeAreaMatrix();
        renderHomeSummary();
        renderCastSelects();
        renderManualLastVehicleInfo();
        return;
      }
      console.error(createError);
      return;
    }
    currentDispatchId = inserted.id;
  }

  const { data, error } = await supabaseClient
    .from(getTableName("dispatch_items"))
    .select(remapRelationSelect(`
      *,
      casts (
        id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude
      )
    `))
    .eq("dispatch_id", currentDispatchId)
    .order("actual_hour", { ascending: true })
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce("dispatch_items", error);
      currentActualsCache = [];
      renderActualTable();
      renderActualTimeAreaMatrix();
      renderHomeSummary();
      renderCastSelects();
      renderManualLastVehicleInfo();
      return;
    }
    console.error(error);
    return;
  }

  currentActualsCache = data || [];
  renderActualTable();
  renderActualTimeAreaMatrix();
  renderHomeSummary();
  renderCastSelects();
  renderManualLastVehicleInfo();
}

async function savePlan() {
  const cast = findCastByInputValue(els.planCastSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const planDate = els.planDate?.value || todayStr();
  const hour = Number(els.planHour?.value || 0);
  const address = els.planAddress?.value.trim() || "";
  let distanceKm = toNullableNumber(els.planDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.planDistanceKm) els.planDistanceKm.value = String(distanceKm);
  }
  const area = els.planArea?.value.trim() || "";
  const note = els.planNote?.value.trim() || "";

  const payload = {
    plan_date: planDate,
    plan_hour: hour,
    cast_id: castId,
    destination_address: address,
    planned_area: normalizeAreaLabel(area || "無し"),
    distance_km: distanceKm,
    note,
    status: "planned"
  };

  let error;
  if (editingPlanId) {
    ({ error } = await supabaseClient.from(getTableName("dispatch_plans")).update(payload).eq("id", editingPlanId));
  } else {
    payload.created_by = getCurrentUserIdSafe();
    ({ error } = await supabaseClient.from(getTableName("dispatch_plans")).insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    null,
    null,
    editingPlanId ? "update_plan" : "create_plan",
    editingPlanId ? "予定を更新" : "予定を作成"
  );

  resetPlanForm();
  await loadPlansByDate(planDate);
}

async function deletePlan(planId) {
  if (!window.confirm("この予定を削除しますか？")) return;

  const { error } = await supabaseClient.from(getTableName("dispatch_plans")).delete().eq("id", planId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "delete_plan", `予定ID ${planId} を削除`);
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function clearAllPlans() {
  if (!window.confirm("この日の予定を全消去しますか？")) return;

  const planDate = els.planDate?.value || todayStr();
  const { error } = await supabaseClient
    .from(getTableName("dispatch_plans"))
    .delete()
    .eq("plan_date", planDate);

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "clear_plans", `${planDate} の予定を全削除`);
  await loadPlansByDate(planDate);
}

async function saveActual() {
  const cast = findCastByInputValue(els.castSelect?.value || "");
  const castId = Number(cast?.id || 0);
  if (!castId) {
    alert("キャストを選択または入力してください");
    return;
  }

  const dateStr = els.actualDate?.value || todayStr();
  const hour = Number(els.actualHour?.value || 0);
  const address = els.actualAddress?.value.trim() || "";
  const area = normalizeAreaLabel(els.actualArea?.value.trim() || "無し");
  let distanceKm = toNullableNumber(els.actualDistanceKm?.value);
  if (distanceKm === null) {
    distanceKm = await resolveDistanceKmForCastRecord(cast, address);
    if (distanceKm !== null && els.actualDistanceKm) els.actualDistanceKm.value = String(distanceKm);
  }
  const status = els.actualStatus?.value || "pending";
  const note = els.actualNote?.value.trim() || "";

  const existingActual = editingActualId
    ? currentActualsCache.find(x => Number(x.id) === Number(editingActualId))
    : null;

  const stopOrder = existingActual
    ? Number(existingActual.stop_order || 1)
    : currentActualsCache.filter(
        x =>
          Number(x.actual_hour) === hour &&
          Number(x.id) !== Number(editingActualId || 0)
      ).length + 1;

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: castId,
    actual_hour: hour,
    stop_order: stopOrder,
    pickup_label: ORIGIN_LABEL,
    destination_address: address,
    destination_area: area,
    distance_km: distanceKm,
    status,
    note,
    plan_date: dateStr
  };

  let error;
  if (editingActualId) {
    ({ error } = await supabaseClient.from(getTableName("dispatch_items")).update(payload).eq("id", editingActualId));
  } else {
    ({ error } = await supabaseClient.from(getTableName("dispatch_items")).insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(
    currentDispatchId,
    editingActualId || null,
    editingActualId ? "update_actual" : "create_actual",
    editingActualId ? "実際の送りを更新" : "実際の送りを追加"
  );

  resetActualForm();
  await loadActualsByDate(dateStr);
  if (!editingActualId) {
    try {
      await assignUnassignedActualsForToday();
      await loadActualsByDate(dateStr);
    } catch (assignError) {
      console.error("assignUnassignedActualsForToday error:", assignError);
    }
  }
  await loadPlansByDate(els.planDate?.value || dateStr);
}

async function deleteActual(itemId) {
  if (!window.confirm("このActualを削除しますか？")) return;

  const { error } = await supabaseClient.from(getTableName("dispatch_items")).delete().eq("id", itemId);
  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, itemId, "delete_actual", `Actual ID ${itemId} を削除`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function updateActualStatus(itemId, status) {
  const item = currentActualsCache.find(x => Number(x.id) === Number(itemId));
  if (!item) {
    alert("対象のActualが見つかりません");
    return;
  }

  const { error } = await supabaseClient
    .from(getTableName("dispatch_items"))
    .update({ status })
    .eq("id", itemId);

  if (error) {
    alert(error.message);
    return;
  }

  const targetPlan = currentPlansCache.find(
    plan =>
      Number(plan.cast_id) === Number(item.cast_id) &&
      plan.plan_date === (els.actualDate?.value || todayStr()) &&
      Number(plan.plan_hour) === Number(item.actual_hour ?? -1)
  );

  if (targetPlan) {
    let nextPlanStatus = targetPlan.status;
    if (status === "done") nextPlanStatus = "done";
    else if (status === "cancel") nextPlanStatus = "cancel";
    else if (status === "pending") nextPlanStatus = "assigned";

    const { error: planError } = await supabaseClient
      .from(getTableName("dispatch_plans"))
      .update({ status: nextPlanStatus })
      .eq("id", targetPlan.id);

    if (planError) console.error(planError);
  }

  await addHistory(currentDispatchId, itemId, "update_actual_status", `Actual状態を ${status} に変更`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function addPlanToActual() {
  const planId = String(els.planSelect?.value || "").trim();
  if (!planId) {
    alert("予定を選択してください");
    return;
  }

  const plan = currentPlansCache.find(x => sameDispatchEntityId(x.id, planId));
  if (!plan) {
    alert("予定が見つかりません");
    return;
  }

  if (isPlanAlreadyAddedToActual(plan)) {
    alert("その予定はすでにActualへ追加されています");
    renderPlanSelect();
    return;
  }

  if (currentActualsCache.some(x => sameDispatchEntityId(x.cast_id, plan.cast_id) && normalizeStatus(x.status) !== "cancel")) {
    alert("そのキャストはすでにActualにあります");
    renderPlanSelect();
    return;
  }

  const doneCastIds = getDoneCastIdsInActuals();
  if (doneCastIds.has(normalizeDispatchEntityId(plan.cast_id))) {
    alert("このキャストはすでに送り完了です");
    renderPlanSelect();
    return;
  }

  const payload = {
    dispatch_id: currentDispatchId,
    cast_id: plan.cast_id,
    actual_hour: Number(plan.plan_hour || 0),
    stop_order:
      currentActualsCache.filter(x => Number(x.actual_hour) === Number(plan.plan_hour || 0)).length + 1,
    pickup_label: ORIGIN_LABEL,
    destination_address: plan.destination_address || plan.casts?.address || "",
    destination_area: normalizeAreaLabel(plan.planned_area || "無し"),
    distance_km: plan.distance_km ?? plan.casts?.distance_km ?? null,
    status: "pending",
    note: plan.note || "",
    plan_date: plan.plan_date
  };

  const { error } = await supabaseClient.from(getTableName("dispatch_items")).insert(payload);
  if (error) {
    alert(error.message);
    return;
  }

  await supabaseClient
    .from(getTableName("dispatch_plans"))
    .update({ status: "assigned" })
    .eq("id", plan.id);

  await addHistory(currentDispatchId, null, "add_plan_to_actual", `予定ID ${plan.id} をActualへ追加`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  if (els.planSelect) els.planSelect.value = "";
  renderPlanSelect();
}


function getDefaultTeamPlanPayload(overrides = {}) {
  return {
    plan_type: "free",
    limits: {
      members: 3,
      origins: 1,
      vehicles: 4,
      casts: 50
    },
    feature_flags: {
      csv: false,
      line: true,
      monthly_full: true,
      plan_to_actual_add: true,
      multi_user: true,
      backup: false,
      restore: false,
      google_api_dispatch: false
    },
    billing_status: "inactive",
    billing_current_period_end: null,
    billing_customer_id: null,
    billing_subscription_id: null,
    billing_price_id: null,
    billing_checkout_session_id: null,
    billing_cancel_at_period_end: false,
    ...(overrides && typeof overrides === "object" ? overrides : {})
  };
}

function normalizeTeamPlanPayload(row = {}) {
  const base = getDefaultTeamPlanPayload();
  let limits = row?.limits;
  if (typeof limits === "string") {
    try { limits = JSON.parse(limits); } catch (_) { limits = null; }
  }
  let featureFlags = row?.feature_flags;
  if (typeof featureFlags === "string") {
    try { featureFlags = JSON.parse(featureFlags); } catch (_) { featureFlags = null; }
  }
  return {
    ...base,
    ...(row && typeof row === "object" ? row : {}),
    plan_type: String(row?.plan_type || base.plan_type).trim() === "paid" ? "paid" : "free",
    limits: {
      ...base.limits,
      ...(limits && typeof limits === "object" ? limits : {})
    },
    feature_flags: {
      ...base.feature_flags,
      ...(featureFlags && typeof featureFlags === "object" ? featureFlags : {})
    },
    billing_cancel_at_period_end: Boolean(row?.billing_cancel_at_period_end)
  };
}

async function getTeamPlan(teamId) {
  const safeTeamId = String(teamId || "").trim();
  if (!safeTeamId) {
    return { data: normalizeTeamPlanPayload(), error: null };
  }

  const teamsTable = getTableName("teams");

  try {
    const { data, error } = await supabaseClient
      .from(teamsTable)
      .select("id, plan_type, limits, feature_flags, billing_status, billing_current_period_end, billing_customer_id, billing_subscription_id, billing_price_id, billing_checkout_session_id, billing_cancel_at_period_end")
      .eq("id", safeTeamId)
      .maybeSingle();

    if (error) {
      if (isMissingColumnError(error)) {
        console.warn("team plan columns are not available yet. fallback to default free plan.", error);
        return { data: normalizeTeamPlanPayload({ id: safeTeamId }), error: null };
      }
      return { data: normalizeTeamPlanPayload({ id: safeTeamId }), error };
    }

    return { data: normalizeTeamPlanPayload({ ...(data || {}), id: safeTeamId }), error: null };
  } catch (error) {
    return { data: normalizeTeamPlanPayload({ id: safeTeamId }), error };
  }
}


function normalizeTeamPlanUsagePayload(payload = {}) {
  const members = Number(payload?.members || 0);
  const pendingInvitations = Number(payload?.pending_invitations || payload?.pendingInvitations || 0);
  const seatsUsedRaw = Number(payload?.seats_used);
  const safeMembers = Number.isFinite(members) && members > 0 ? members : 0;
  const safePending = Number.isFinite(pendingInvitations) && pendingInvitations > 0 ? pendingInvitations : 0;
  const safeSeatsUsed = Number.isFinite(seatsUsedRaw) && seatsUsedRaw >= 0 ? seatsUsedRaw : safeMembers + safePending;
  return {
    members: safeMembers,
    pending_invitations: safePending,
    seats_used: safeSeatsUsed
  };
}

async function getTeamPlanMemberUsage(teamId) {
  const safeTeamId = String(teamId || "").trim();
  if (!safeTeamId) {
    return { data: normalizeTeamPlanUsagePayload(), error: null };
  }

  let members = 0;
  let pendingInvitations = 0;
  let firstError = null;

  try {
    const { count, error } = await supabaseClient
      .from(getTableName("team_members"))
      .select("*", { count: "exact", head: true })
      .eq("team_id", safeTeamId);
    if (error) {
      if (isMissingTableError(error)) {
        warnMissingTableOnce("team_members", error);
      } else {
        firstError = firstError || error;
      }
    } else {
      members = Number(count || 0);
    }
  } catch (error) {
    firstError = firstError || error;
  }

  if (!isKnownMissingTable("invitations")) {
    try {
      let query = supabaseClient
        .from(getTableName("invitations"))
        .select("*", { count: "exact", head: true })
        .eq("team_id", safeTeamId)
        .eq("status", "pending");
      let { count, error } = await query;

      if (error && isMissingColumnError(error)) {
        const retry = await supabaseClient
          .from(getTableName("invitations"))
          .select("*")
          .eq("team_id", safeTeamId);
        count = Array.isArray(retry.data)
          ? retry.data.map(normalizeInvitationRow).filter(row => row?.status === "pending").length
          : 0;
        error = retry.error;
      }

      if (error) {
        if (isMissingTableError(error)) {
          warnMissingTableOnce("invitations", error);
        } else {
          firstError = firstError || error;
        }
      } else {
        pendingInvitations = Number(count || 0);
      }
    } catch (error) {
      firstError = firstError || error;
    }
  }

  return {
    data: normalizeTeamPlanUsagePayload({
      members,
      pending_invitations: pendingInvitations,
      seats_used: members + pendingInvitations
    }),
    error: firstError
  };
}

// platform admin helpers
function normalizeAdminTeamName(row = {}) {
  return String(
    row.team_name ?? row.name ?? row.workspace_name ?? row.team_label ?? row.label ?? row.title ?? ""
  ).trim();
}

function buildAdminTeamStatusPayload(status, reason = "") {
  const normalized = String(status || "active").trim() === "suspended" ? "suspended" : "active";
  return normalized === "suspended"
    ? {
        status: "suspended",
        suspended_at: new Date().toISOString(),
        suspended_reason: String(reason || "").trim() || null
      }
    : {
        status: "active",
        suspended_at: null,
        suspended_reason: null
      };
}

async function getAllTeamsForAdmin() {
  const tableName = getTableName("teams");
  const { data, error } = await supabaseClient
    .from(tableName)
    .select("*");

  if (error) {
    return { data: [], error };
  }

  const rows = Array.isArray(data) ? data : [];
  const ownerIds = [...new Set(rows.map(r => String(r?.owner_user_id || "").trim()).filter(Boolean))];
  let profilesById = {};

  if (ownerIds.length) {
    try {
      const profileTable = getTableName("profiles");
      const { data: profileRows } = await supabaseClient
        .from(profileTable)
        .select("*")
        .in("id", ownerIds);
      profilesById = Object.fromEntries((Array.isArray(profileRows) ? profileRows : []).map(row => [String(row?.id || ""), row]));
    } catch (_) {}
  }

  const normalized = rows
    .map(row => {
      const owner = profilesById[String(row?.owner_user_id || "").trim()] || null;
      return {
        ...row,
        team_name: normalizeAdminTeamName(row),
        owner_email: String(owner?.email || "").trim(),
        owner_display_name: String(owner?.display_name || owner?.name || "").trim(),
        status: String(row?.status || "active").trim() || "active"
      };
    })
    .sort((a, b) => String(b?.created_at || "").localeCompare(String(a?.created_at || ""), "ja"));

  return { data: normalized, error: null };
}

async function getTeamMembersForAdmin(teamId) {
  const memberTable = getTableName("team_members");
  const { data, error } = await supabaseClient
    .from(memberTable)
    .select("*")
    .eq("team_id", teamId);

  if (error) {
    return { data: [], error };
  }

  const memberRows = Array.isArray(data) ? data : [];
  const userIds = [...new Set(memberRows.map(r => String(r?.user_id || "").trim()).filter(Boolean))];
  let profilesById = {};

  if (userIds.length) {
    try {
      const profileTable = getTableName("profiles");
      const { data: profileRows } = await supabaseClient
        .from(profileTable)
        .select("*")
        .in("id", userIds);
      profilesById = Object.fromEntries((Array.isArray(profileRows) ? profileRows : []).map(row => [String(row?.id || ""), row]));
    } catch (_) {}
  }

  const normalized = memberRows
    .map(row => {
      const profile = profilesById[String(row?.user_id || "").trim()] || null;
      return {
        ...row,
        role: String(row?.role || "user").trim() || "user",
        email: String(profile?.email || "").trim(),
        display_name: String(profile?.display_name || profile?.name || "").trim(),
        is_active: profile?.is_active !== false
      };
    })
    .sort((a, b) => String(a?.created_at || "").localeCompare(String(b?.created_at || ""), "ja"));

  return { data: normalized, error: null };
}

async function getTeamSummaryForAdmin(teamId) {
  const teamsTable = getTableName("teams");
  const { data: team, error: teamError } = await supabaseClient
    .from(teamsTable)
    .select("*")
    .eq("id", teamId)
    .maybeSingle();

  if (teamError) {
    return { data: null, error: teamError };
  }

  const countFor = async (logicalName) => {
    try {
      let query = supabaseClient
        .from(getTableName(logicalName))
        .select("*", { count: "exact", head: true })
        .eq("team_id", teamId);
      if (logicalName === "casts") {
        query = query.eq("is_active", true);
      }
      const { count, error } = await query;
      if (!error) return Number(count || 0);
      if (logicalName === "casts" && isMissingColumnError(error) && /is_active/i.test(String(error?.message || ""))) {
        const fallback = await supabaseClient
          .from(getTableName(logicalName))
          .select("*", { count: "exact", head: true })
          .eq("team_id", teamId);
        return Number(fallback.count || 0);
      }
      return 0;
    } catch (_) {
      return 0;
    }
  };

  const [membersCount, vehiclesCount, castsCount, originsCount, dispatchesCount] = await Promise.all([
    countFor("team_members"),
    countFor("vehicles"),
    countFor("casts"),
    countFor("origins"),
    countFor("dispatches")
  ]);

  const { data: members } = await getTeamMembersForAdmin(teamId);
  return {
    data: {
      ...(team || {}),
      team_name: normalizeAdminTeamName(team || {}),
      status: String(team?.status || "active").trim() || "active",
      members_count: membersCount,
      vehicles_count: vehiclesCount,
      casts_count: castsCount,
      origins_count: originsCount,
      dispatches_count: dispatchesCount,
      members: Array.isArray(members) ? members : []
    },
    error: null
  };
}

async function setTeamStatusForAdmin(teamId, status, reason = "") {
  const teamsTable = getTableName("teams");
  const payload = buildAdminTeamStatusPayload(status, reason);
  const { data, error } = await supabaseClient
    .from(teamsTable)
    .update(payload)
    .eq("id", teamId)
    .select("*")
    .maybeSingle();

  return { data, error };
}

async function updateTeamNameForAdmin(teamId, nextName) {
  const safeTeamId = String(teamId || '').trim();
  const safeName = String(nextName || '').trim().replace(/\s+/g, ' ');
  if (!safeTeamId) return { data: null, error: new Error('teamId is required') };
  if (!safeName) return { data: null, error: new Error('team_name is required') };

  const teamsTable = getTableName('teams');
  let payload = {
    team_name: safeName,
    updated_at: new Date().toISOString()
  };

  let { data, error } = await supabaseClient
    .from(teamsTable)
    .update(payload)
    .eq('id', safeTeamId)
    .select('*')
    .maybeSingle();

  while (error && isMissingColumnError(error)) {
    const missingColumn = getMissingColumnName(error);
    if (!missingColumn || !(missingColumn in payload)) break;
    payload = omitKeys(payload, missingColumn);
    const retry = await supabaseClient
      .from(teamsTable)
      .update(payload)
      .eq('id', safeTeamId)
      .select('*')
      .maybeSingle();
    data = retry.data;
    error = retry.error;
  }

  if (error) {
    return { data: null, error };
  }

  const normalized = normalizeTeamRow({ ...(data || {}), id: safeTeamId, team_name: safeName, name: safeName });
  cacheDropOffTeamMeta(normalized);
  return { data: normalized, error: null };
}



async function updateTeamPlanForAdmin(teamId, nextPlanType) {
  const safeTeamId = String(teamId || '').trim();
  const planType = String(nextPlanType || '').trim() === 'paid' ? 'paid' : 'free';
  if (!safeTeamId) return { data: null, error: new Error('teamId is required') };

  const teamsTable = getTableName('teams');
  const planPayload = planType === 'paid'
    ? {
        plan_type: 'paid',
        limits: { members: null, origins: 5, vehicles: null, casts: null },
        feature_flags: {
          csv: true,
          line: true,
          monthly_full: true,
          plan_to_actual_add: true,
          multi_user: true,
          backup: true,
          restore: true,
          google_api_dispatch: true
        }
      }
    : {
        plan_type: 'free',
        limits: { members: 3, origins: 1, vehicles: 4, casts: 50 },
        feature_flags: {
          csv: false,
          line: true,
          monthly_full: true,
          plan_to_actual_add: true,
          multi_user: true,
          backup: false,
          restore: false,
          google_api_dispatch: false
        }
      };

  let payload = {
    ...planPayload,
    updated_at: new Date().toISOString(),
    plan_updated_at: new Date().toISOString()
  };

  let { data, error } = await supabaseClient
    .from(teamsTable)
    .update(payload)
    .eq('id', safeTeamId)
    .select('*')
    .maybeSingle();

  while (error && isMissingColumnError(error)) {
    const missingColumn = getMissingColumnName(error);
    if (!missingColumn || !(missingColumn in payload)) break;
    payload = omitKeys(payload, missingColumn);
    const retry = await supabaseClient
      .from(teamsTable)
      .update(payload)
      .eq('id', safeTeamId)
      .select('*')
      .maybeSingle();
    data = retry.data;
    error = retry.error;
  }

  if (error) {
    return { data: null, error };
  }

  const normalized = normalizeTeamPlanPayload({ ...(data || {}), id: safeTeamId, ...planPayload });
  const cachedTeamMeta = getCachedDropOffTeamMeta(safeTeamId);
  if (cachedTeamMeta && typeof cachedTeamMeta === 'object') {
    cacheDropOffTeamMeta({
      ...cachedTeamMeta,
      id: safeTeamId,
      plan_type: normalized.plan_type,
      limits: normalized.limits,
      feature_flags: normalized.feature_flags,
      updated_at: data?.updated_at || payload.updated_at || cachedTeamMeta.updated_at || null
    });
  }

  return { data: normalized, error: null };
}

window.getTeamPlan = getTeamPlan;
window.getTeamPlanMemberUsage = getTeamPlanMemberUsage;
window.getCachedDropOffTeamMeta = getCachedDropOffTeamMeta;
window.cacheDropOffTeamMeta = cacheDropOffTeamMeta;
window.getAllTeamsForAdmin = getAllTeamsForAdmin;
window.getTeamMembersForAdmin = getTeamMembersForAdmin;
window.getTeamSummaryForAdmin = getTeamSummaryForAdmin;
window.setTeamStatusForAdmin = setTeamStatusForAdmin;
window.updateTeamNameForAdmin = updateTeamNameForAdmin;
window.updateTeamPlanForAdmin = updateTeamPlanForAdmin;
