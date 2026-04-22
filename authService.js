// authService
// 認証責務を dashboard.js から分離

const DROPOFF_RECOVERY_FLAG_KEY = "dropoff_password_recovery_pending";

function markRecoveryPending() {
  try {
    sessionStorage.setItem(DROPOFF_RECOVERY_FLAG_KEY, "1");
  } catch (e) {}
}

function hasRecoveryPending() {
  try {
    return sessionStorage.getItem(DROPOFF_RECOVERY_FLAG_KEY) === "1";
  } catch (e) {
    return false;
  }
}

function isRecoveryNavigation() {
  try {
    const hash = new URLSearchParams(String(window.location.hash || '').replace(/^#/, ''));
    const query = new URLSearchParams(String(window.location.search || ''));
    const type = hash.get('type') || query.get('type') || '';
    const code = hash.get('code') || query.get('code') || '';
    const tokenHash = hash.get('token_hash') || query.get('token_hash') || '';
    const hasToken = Boolean(hash.get('access_token') || hash.get('refresh_token') || query.get('access_token') || query.get('refresh_token'));
    return type === 'recovery' || Boolean(code || tokenHash || hasToken) || hasRecoveryPending();
  } catch (e) {
    return hasRecoveryPending();
  }
}

function redirectRecoveryToLogin() {
  markRecoveryPending();
  const target = `index.html${window.location.search || ''}${window.location.hash || ''}`;
  window.location.replace(target);
}

function getCurrentUserIdSafe() {
  try {
    if (typeof currentUser !== "undefined" && currentUser?.id) return currentUser.id;
  } catch (e) {}
  return window.currentUser?.id || null;
}

async function getCurrentUserIdSafeAsync() {
  const syncId = getCurrentUserIdSafe();
  if (syncId) return syncId;
  try {
    const { data, error } = await supabaseClient.auth.getUser();
    if (error) {
      console.error(error);
      return null;
    }
    const user = data?.user || null;
    if (user) {
      window.currentUser = user;
      try {
        if (typeof setCurrentUserState === "function") setCurrentUserState(user);
        else if (typeof currentUser !== "undefined") currentUser = user;
      } catch (e) {}
      return user.id || null;
    }
  } catch (error) {
    console.error(error);
  }
  return null;
}

async function isPlatformAdmin() {
  try {
    const userId = await getCurrentUserIdSafeAsync();
    if (!userId) {
      window.isPlatformAdminUser = false;
      return false;
    }

    const { data, error } = await supabaseClient
      .from("dropoff_platform_admins")
      .select("user_id")
      .eq("user_id", userId)
      .maybeSingle();

    if (error) {
      console.warn("isPlatformAdmin check failed:", error);
      window.isPlatformAdminUser = false;
      return false;
    }

    const ok = !!data?.user_id;
    window.isPlatformAdminUser = ok;
    return ok;
  } catch (error) {
    console.warn("isPlatformAdmin exception:", error);
    window.isPlatformAdminUser = false;
    return false;
  }
}

window.isPlatformAdmin = isPlatformAdmin;


function persistWorkspaceTeamId(teamId) {
  const value = String(teamId || '').trim();
  if (!value) return;
  try { window.localStorage.setItem('dropoff_workspace_team_id', value); } catch (e) {}
  try { window.localStorage.setItem('current_dropoff_team_id', value); } catch (e) {}
  try { window.localStorage.setItem('workspaceTeamId', value); } catch (e) {}
}

function getStoredWorkspaceTeamId() {
  try {
    return String(
      window.localStorage.getItem('dropoff_workspace_team_id') ||
      window.localStorage.getItem('current_dropoff_team_id') ||
      window.localStorage.getItem('workspaceTeamId') ||
      ''
    ).trim() || null;
  } catch (e) {
    return null;
  }
}

function getPerUserWorkspaceTeamId(userId) {
  const safeUserId = String(userId || '').trim();
  if (!safeUserId) return null;
  try {
    return String(window.localStorage.getItem(`dropoff_workspace_team_id_v1_${safeUserId}`) || '').trim() || null;
  } catch (e) {
    return null;
  }
}

function clearWorkspaceTeamCaches() {
  try { window.localStorage.removeItem('dropoff_workspace_team_id'); } catch (e) {}
  try { window.localStorage.removeItem('current_dropoff_team_id'); } catch (e) {}
  try { window.localStorage.removeItem('workspaceTeamId'); } catch (e) {}
  try { window.localStorage.removeItem('__DROP_OFF_LAST_TEAM_ID__'); } catch (e) {}
  try {
    const userId = String(window.currentUser?.id || '').trim();
    if (userId) window.localStorage.removeItem(`dropoff_workspace_team_id_v1_${userId}`);
  } catch (e) {}
  try {
    const cache = window.__DROP_OFF_WORKSPACE_CACHE__ || {};
    const userId = String(window.currentUser?.id || '').trim();
    if (userId && cache && typeof cache === 'object') delete cache[userId];
  } catch (e) {}
}

async function doesWorkspaceTeamExist(teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return false;
  const teamsTable = typeof getTableName === 'function' ? getTableName('teams') : 'dropoff_teams';
  try {
    const { data, error } = await supabaseClient
      .from(teamsTable)
      .select('id')
      .eq('id', safeTeamId)
      .maybeSingle();
    return !error && !!data?.id;
  } catch (e) {
    return false;
  }
}

async function hasWorkspaceMembership(teamId, identityIds = []) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return false;
  const membersTable = typeof getTableName === 'function' ? getTableName('team_members') : 'dropoff_team_members';
  const ids = Array.from(new Set((Array.isArray(identityIds) ? identityIds : [])
    .map(v => String(v || '').trim())
    .filter(Boolean)));
  for (const identityId of ids) {
    try {
      const { data, error } = await supabaseClient
        .from(membersTable)
        .select('team_id')
        .eq('team_id', safeTeamId)
        .eq('user_id', identityId)
        .limit(1);
      if (!error && Array.isArray(data) && data[0]?.team_id) {
        return true;
      }
      if (error && typeof isMissingTableError === 'function' && isMissingTableError(error)) {
        return false;
      }
    } catch (e) {}
  }
  return false;
}


async function getWorkspaceMemberRoleForAuth(teamId, identityIds = [], email = '') {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return null;
  const membersTable = typeof getTableName === 'function' ? getTableName('team_members') : 'dropoff_team_members';
  const ids = Array.from(new Set((Array.isArray(identityIds) ? identityIds : [])
    .map(v => String(v || '').trim())
    .filter(Boolean)));

  for (const identityId of ids) {
    try {
      const { data, error } = await supabaseClient
        .from(membersTable)
        .select('role')
        .eq('team_id', safeTeamId)
        .eq('user_id', identityId)
        .limit(1);
      if (!error && Array.isArray(data) && data[0]?.role) {
        return typeof normalizeProfileRole === 'function' ? normalizeProfileRole(data[0].role) : String(data[0].role || 'user');
      }
      if (error && typeof isMissingTableError === 'function' && isMissingTableError(error)) {
        return null;
      }
    } catch (e) {}
  }

  const safeEmail = String(email || '').trim();
  if (safeEmail) {
    try {
      const { data, error } = await supabaseClient
        .from(membersTable)
        .select('role')
        .eq('team_id', safeTeamId)
        .eq('member_email', safeEmail)
        .limit(1);
      if (!error && Array.isArray(data) && data[0]?.role) {
        return typeof normalizeProfileRole === 'function' ? normalizeProfileRole(data[0].role) : String(data[0].role || 'user');
      }
    } catch (e) {}
  }

  return null;
}

function getAdminForcedTeamId() {
  try {
    return String(window.localStorage.getItem('admin_force_team_id') || '').trim() || null;
  } catch (e) {
    return null;
  }
}

async function resolveWorkspaceTeamIdForAuth(user, profile) {
  const explicit = (() => {
    try {
      const query = new URLSearchParams(String(window.location.search || ''));
      return String(query.get('team_id') || '').trim() || null;
    } catch (e) {
      return null;
    }
  })();

  const isAdmin = !!window.isPlatformAdminUser;
  const forcedTeamId = isAdmin ? getAdminForcedTeamId() : null;
  const identityIds = Array.from(new Set([
    user?.id,
    profile?.id,
    profile?.user_id
  ].map(v => String(v || '').trim()).filter(Boolean)));

  if (forcedTeamId && await doesWorkspaceTeamExist(forcedTeamId)) {
    window.currentWorkspaceTeamId = forcedTeamId;
    persistWorkspaceTeamId(forcedTeamId);
    return forcedTeamId;
  }

  const candidateList = [
    getPerUserWorkspaceTeamId(user?.id),
    profile?.current_dropoff_team_id,
    window.currentWorkspaceTeamId,
    isAdmin ? explicit : null,
    getStoredWorkspaceTeamId()
  ].map(v => String(v || '').trim()).filter(Boolean);

  for (const candidate of candidateList) {
    try {
      const exists = await doesWorkspaceTeamExist(candidate);
      if (!exists) continue;
      const allowed = await hasWorkspaceMembership(candidate, identityIds);
      if (!allowed) continue;
      window.currentWorkspaceTeamId = candidate;
      persistWorkspaceTeamId(candidate);
      return candidate;
    } catch (e) {}
  }

  const membersTable = typeof getTableName === 'function' ? getTableName('team_members') : 'dropoff_team_members';
  try {
    const allRows = [];
    for (const identityId of identityIds) {
      const { data, error } = await supabaseClient
        .from(membersTable)
        .select('team_id, role, created_at, user_id')
        .eq('user_id', identityId)
        .order('created_at', { ascending: true });
      if (error) continue;
      if (Array.isArray(data)) allRows.push(...data);
    }
    if (allRows.length) {
      const uniqueRows = [];
      const seen = new Set();
      for (const row of allRows) {
        const teamId = String(row?.team_id || '').trim();
        if (!teamId || seen.has(teamId)) continue;
        seen.add(teamId);
        uniqueRows.push(row);
      }
      const prioritized = uniqueRows.slice().sort((a, b) => {
        const ar = String(a?.role || 'user');
        const br = String(b?.role || 'user');
        const aw = ar === 'owner' ? 0 : ar === 'admin' ? 1 : 2;
        const bw = br === 'owner' ? 0 : br === 'admin' ? 1 : 2;
        if (aw !== bw) return aw - bw;
        return String(a?.created_at || '').localeCompare(String(b?.created_at || ''));
      })[0];
      const teamId = String(prioritized?.team_id || '').trim() || null;
      if (teamId) {
        window.currentWorkspaceTeamId = teamId;
        persistWorkspaceTeamId(teamId);
        return teamId;
      }
    }
  } catch (e) {
    console.warn('resolveWorkspaceTeamIdForAuth failed:', e);
  }
  return null;
}

async function getWorkspaceStatusForAuth(teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) return null;
  const teamsTable = typeof getTableName === 'function' ? getTableName('teams') : 'dropoff_teams';
  try {
    const { data, error } = await supabaseClient
      .from(teamsTable)
      .select('*')
      .eq('id', safeTeamId)
      .maybeSingle();
    if (error) {
      return null;
    }
    return data || null;
  } catch (e) {
    return null;
  }
}

async function upsertProfileSafely(user) {
  if (typeof ensureCurrentUserProfileCloud === "function") {
    try {
      return await ensureCurrentUserProfileCloud(user);
    } catch (error) {
      console.error(error);
    }
  }

  const tableName = typeof getTableName === "function" ? getTableName("profiles") : "dropoff_profiles";
  if (typeof isKnownMissingTable === "function" && isKnownMissingTable("profiles")) return null;

  const payloads = [
    { id: user.id, user_id: user.id, email: user.email, display_name: user.email, role: "owner", is_active: true },
    { id: user.id, email: user.email, display_name: user.email },
    { id: user.id, email: user.email }
  ];

  let lastError = null;
  for (const payload of payloads) {
    const { data, error } = await supabaseClient.from(tableName).upsert(payload).select('*').maybeSingle();
    if (!error) {
      return data || payload;
    }
    lastError = error;
    if (typeof isMissingTableError === "function" && isMissingTableError(error)) {
      if (typeof warnMissingTableOnce === "function") warnMissingTableOnce("profiles", error);
      return null;
    }
    const message = String(error?.message || "");
    const code = String(error?.code || "");
    const missingColumn = code === "PGRST204" || code === "42703" || /Could not find the '.+?' column/i.test(message) || /column .+ does not exist/i.test(message);
    if (!missingColumn) break;
  }

  if (lastError) {
    console.error(lastError);
  }
  return null;
}

async function ensureAuth() {
  let recoveryRedirected = false;

  try {
    supabaseClient.auth.onAuthStateChange((event) => {
      if (event === 'PASSWORD_RECOVERY') {
        markRecoveryPending();
        if (!recoveryRedirected) {
          recoveryRedirected = true;
          redirectRecoveryToLogin();
        }
      }
    });
  } catch (e) {
    console.error(e);
  }

  if (isRecoveryNavigation()) {
    redirectRecoveryToLogin();
    return false;
  }

  const { data, error } = await supabaseClient.auth.getUser();

  if (error) {
    alert("ユーザー情報の取得に失敗しました");
    window.location.href = "index.html";
    return false;
  }

  const user = data?.user || null;
  window.currentUser = user;
  try {
    if (typeof setCurrentUserState === "function") setCurrentUserState(user);
    else if (typeof currentUser !== "undefined") currentUser = user;
  } catch (e) {}

  if (!user) {
    window.location.href = "index.html";
    return false;
  }

  let accessInfo = { allowed: true, source: "legacy" };
  if (typeof ensureInvitedAccessForUser === "function") {
    try {
      accessInfo = await ensureInvitedAccessForUser(user);
    } catch (error) {
      console.error(error);
      alert("招待情報の確認に失敗しました。");
      await supabaseClient.auth.signOut();
      window.location.href = "index.html";
      return false;
    }
  }

  if (accessInfo?.allowed === false) {
    alert(accessInfo.reason || "このアカウントでは参加できません。招待を確認してください。");
    await supabaseClient.auth.signOut();
    window.location.href = "index.html";
    return false;
  }

  let profile = accessInfo?.profile || await upsertProfileSafely(user);
  window.currentUserProfile = profile || null;
  try {
    if (typeof setCurrentUserProfileState === "function") setCurrentUserProfileState(profile || null);
  } catch (e) {}

  if ((profile?.is_active) === false) {
    alert("このユーザーは無効化されています。オーナーへ確認してください。");
    await supabaseClient.auth.signOut();
    window.location.href = "index.html";
    return false;
  }

  if (typeof els !== "undefined" && els?.userEmail) {
    els.userEmail.value = user.email || "";
  }

  try {
    await isPlatformAdmin();
  } catch (e) {
    window.isPlatformAdminUser = false;
  }

  try {
    const teamId = await resolveWorkspaceTeamIdForAuth(user, profile || null);
    const identityIds = Array.from(new Set([
      user?.id,
      profile?.id,
      profile?.user_id
    ].map(v => String(v || '').trim()).filter(Boolean)));
    const effectiveRole = await getWorkspaceMemberRoleForAuth(teamId, identityIds, user?.email || profile?.email || '');
    if (effectiveRole) {
      profile = { ...(profile || {}), role: effectiveRole };
      window.currentUserProfile = profile;
      try {
        if (typeof setCurrentUserProfileState === "function") setCurrentUserProfileState(profile || null);
      } catch (e) {}
    }
    const workspace = await getWorkspaceStatusForAuth(teamId);
    window.currentWorkspaceTeamId = teamId || null;
    window.currentWorkspaceInfo = workspace || null;
    window.currentWorkspaceSuspended = !window.isPlatformAdminUser && String(workspace?.status || 'active').trim() === 'suspended';
  } catch (e) {
    window.currentWorkspaceSuspended = false;
    window.currentWorkspaceInfo = null;
  }

  if (typeof els !== "undefined" && els?.userRoleText) {
    const role = typeof normalizeProfileRole === "function" ? normalizeProfileRole(profile?.role) : String(profile?.role || "user");
    const label = role === "owner" ? "オーナー" : role === "admin" ? "管理者" : "利用者";
    els.userRoleText.value = label;
  }

  return true;
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.currentUser = null;
  window.currentUserProfile = null;
  window.isPlatformAdminUser = false;
  window.currentWorkspaceSuspended = false;
  window.currentWorkspaceInfo = null;
  clearWorkspaceTeamCaches();
  try { window.localStorage.removeItem('admin_force_team_id'); } catch (e) {}
  try { window.localStorage.removeItem('admin_force_team_name'); } catch (e) {}
  try { window.localStorage.removeItem('admin_force_prev_team_id'); } catch (e) {}
  try { window.localStorage.removeItem('admin_force_prev_team_name'); } catch (e) {}
  try {
    if (typeof setCurrentUserState === "function") setCurrentUserState(null);
    else if (typeof currentUser !== "undefined") currentUser = null;
    if (typeof setCurrentUserProfileState === "function") setCurrentUserProfileState(null);
  } catch (e) {}
  window.location.href = "index.html";
}

window.ensureAuth = ensureAuth;
window.logout = logout;
window.getCurrentUserIdSafe = getCurrentUserIdSafe;
window.getCurrentUserIdSafeAsync = getCurrentUserIdSafeAsync;
