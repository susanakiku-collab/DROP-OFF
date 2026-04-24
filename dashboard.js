
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
  ORIGIN_LABEL: DEFAULT_ORIGIN_LABEL,
  ORIGIN_LAT: DEFAULT_ORIGIN_LAT,
  ORIGIN_LNG: DEFAULT_ORIGIN_LNG
} = window.APP_CONFIG;

let ORIGIN_LABEL = DEFAULT_ORIGIN_LABEL;
let ORIGIN_LAT = Number(DEFAULT_ORIGIN_LAT);
let ORIGIN_LNG = Number(DEFAULT_ORIGIN_LNG);

const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_ANON_KEY
);
window.supabaseClient = supabaseClient;

let currentUser = null;
let currentUserProfile = null;
let allUserProfilesCache = [];
let allInvitationRowsCache = [];
let invitationTeamOptionsCache = [];
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
let allOriginsCache = [];
let editingOriginSlotNo = null;
let activeOriginSlotNo = null;
const ORIGIN_SLOT_LIMIT = 5;
const ACTIVE_ORIGIN_SLOT_STORAGE_KEY = "dropoff_active_origin_slot_v1";
const ORIGIN_LOCAL_BACKUP_STORAGE_KEY = "dropoff_origin_slots_backup_v1";

let platformAdminTeamsCache = [];
let selectedPlatformAdminTeamId = null;
let suspendedWorkspaceViewRendered = false;

const ADMIN_FORCE_TEAM_ID_KEY = 'admin_force_team_id';
const ADMIN_FORCE_TEAM_NAME_KEY = 'admin_force_team_name';
const ADMIN_FORCE_PREV_TEAM_ID_KEY = 'admin_force_prev_team_id';
const ADMIN_FORCE_PREV_TEAM_NAME_KEY = 'admin_force_prev_team_name';

function getAdminForcedTeamId() {
  try {
    return String(window.localStorage.getItem(ADMIN_FORCE_TEAM_ID_KEY) || '').trim() || null;
  } catch (_) {
    return null;
  }
}

function getAdminForcedTeamName() {
  try {
    return String(window.localStorage.getItem(ADMIN_FORCE_TEAM_NAME_KEY) || '').trim() || null;
  } catch (_) {
    return null;
  }
}

function getCurrentWorkspaceTeamIdSync() {
  try {
    const uid = String(currentUser?.id || window.currentUser?.id || '').trim();
    const perUserCached = uid ? String(window.localStorage.getItem(`dropoff_workspace_team_id_v1_${uid}`) || '').trim() || null : null;
    const forcedTeamId = window.isPlatformAdminUser ? String(window.localStorage.getItem(ADMIN_FORCE_TEAM_ID_KEY) || '').trim() || null : null;
    return (
      forcedTeamId ||
      String(window.currentWorkspaceTeamId || '').trim() ||
      perUserCached ||
      window.localStorage.getItem('dropoff_workspace_team_id') ||
      window.localStorage.getItem('current_dropoff_team_id') ||
      window.localStorage.getItem('workspaceTeamId') ||
      null
    );
  } catch (_) {
    return null;
  }
}

function normalizeWorkspaceTeamId(teamId) {
  return String(teamId || "").trim() || null;
}

async function resolveCurrentWorkspaceTeamId() {
  try {
    const resolved = typeof ensureDropOffWorkspaceId === "function"
      ? await ensureDropOffWorkspaceId()
      : null;
    return normalizeWorkspaceTeamId(resolved || getCurrentWorkspaceTeamIdSync());
  } catch (_) {
    return normalizeWorkspaceTeamId(getCurrentWorkspaceTeamIdSync());
  }
}

function getCurrentWorkspaceCastIdSet() {
  const set = new Set();
  for (const cast of Array.isArray(allCastsCache) ? allCastsCache : []) {
    const castId = normalizeDispatchEntityId(cast?.id || cast?.cast_id);
    if (castId) set.add(castId);
  }
  return set;
}

function filterDispatchRowsByWorkspace(rows, teamId) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const safeTeamId = normalizeWorkspaceTeamId(teamId || getCurrentWorkspaceTeamIdSync());
  const workspaceCastIds = getCurrentWorkspaceCastIdSet();

  return safeRows.filter(row => {
    const rowTeamId = normalizeWorkspaceTeamId(row?.team_id);
    if (safeTeamId && rowTeamId) return rowTeamId === safeTeamId;

    const relatedCastTeamId = normalizeWorkspaceTeamId(row?.casts?.team_id);
    if (safeTeamId && relatedCastTeamId) return relatedCastTeamId === safeTeamId;

    const rowCastId = normalizeDispatchEntityId(row?.cast_id || row?.person_id || row?.casts?.id);
    if (rowCastId && workspaceCastIds.size) return workspaceCastIds.has(rowCastId);

    return !safeTeamId;
  });
}

function persistWorkspaceTeamIdLocally(teamId) {
  const value = String(teamId || '').trim();
  if (!value) return;
  try { window.localStorage.setItem('dropoff_workspace_team_id', value); } catch (_) {}
  try { window.localStorage.setItem('current_dropoff_team_id', value); } catch (_) {}
  try { window.localStorage.setItem('workspaceTeamId', value); } catch (_) {}
  try { window.localStorage.setItem('__DROP_OFF_LAST_TEAM_ID__', value); } catch (_) {}
  try {
    const uid = String(currentUser?.id || window.currentUser?.id || '').trim();
    if (uid) window.localStorage.setItem(`dropoff_workspace_team_id_v1_${uid}`, value);
  } catch (_) {}
  try {
    window.__DROP_OFF_WORKSPACE_CACHE__ = window.__DROP_OFF_WORKSPACE_CACHE__ || {};
    const uid = String(currentUser?.id || window.currentUser?.id || '').trim();
    if (uid) window.__DROP_OFF_WORKSPACE_CACHE__[uid] = value;
  } catch (_) {}
}

function renderAdminForceModeBanner() {
  const wrap = els.platformAdminForceBanner;
  const textNode = els.platformAdminForceBannerText;
  if (!wrap || !textNode) return;
  const forcedTeamId = getAdminForcedTeamId();
  if (!window.isPlatformAdminUser || !forcedTeamId) {
    wrap.classList.add('hidden');
    return;
  }
  const teamName = getAdminForcedTeamName() || String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || 'このチーム');
  textNode.textContent = `現在、運営者として ${teamName} チームを表示中です。`;
  wrap.classList.remove('hidden');
}

function clearAdminForceTeamStorage() {
  const prevTeamId = (() => { try { return String(window.localStorage.getItem(ADMIN_FORCE_PREV_TEAM_ID_KEY) || '').trim() || null; } catch (_) { return null; } })();
  try { window.localStorage.removeItem(ADMIN_FORCE_TEAM_ID_KEY); } catch (_) {}
  try { window.localStorage.removeItem(ADMIN_FORCE_TEAM_NAME_KEY); } catch (_) {}
  try { window.localStorage.removeItem(ADMIN_FORCE_PREV_TEAM_ID_KEY); } catch (_) {}
  try { window.localStorage.removeItem(ADMIN_FORCE_PREV_TEAM_NAME_KEY); } catch (_) {}
  if (prevTeamId) {
    persistWorkspaceTeamIdLocally(prevTeamId);
    window.currentWorkspaceTeamId = prevTeamId;
  }
}

function clearWorkspaceScopedRuntimeCaches() {
  try { currentPlansCache = []; } catch (_) {}
  try { currentActualsCache = []; } catch (_) {}
  try { currentDailyReportsCache = []; } catch (_) {}
  try { currentMileageExportRows = []; } catch (_) {}
  try { activeVehicleIdsForToday = new Set(); } catch (_) {}
  try { allOriginsCache = []; } catch (_) {}
  try { allCastsCache = []; } catch (_) {}
  try { allVehiclesCache = []; } catch (_) {}
  try { window.__DROP_OFF_WORKSPACE_CACHE__ = {}; } catch (_) {}
}

function clearForceTeamMode() {
  clearWorkspaceScopedRuntimeCaches();
  clearAdminForceTeamStorage();
  window.location.href = 'dashboard.html?platform_admin=1';
}

function isSameTeamId(a, b) {
  return String(a || '').trim() !== '' && String(a || '').trim() === String(b || '').trim();
}

function sanitizeDownloadLabel(value, fallback = 'team') {
  const raw = String(value || '').trim() || fallback;
  const sanitized = raw.replace(/[\/:*?"<>|\s]+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '');
  return sanitized || fallback;
}

function getPlatformAdminTeamRow(teamId) {
  return (Array.isArray(platformAdminTeamsCache) ? platformAdminTeamsCache : []).find(row => isSameTeamId(row?.id, teamId)) || null;
}

function getPlatformAdminTeamLabel(teamId, fallback = 'このチーム') {
  const row = getPlatformAdminTeamRow(teamId);
  const name = String(row?.team_name || row?.name || row?.workspace_name || row?.team_label || '').trim();
  if (name) return name;
  if (isSameTeamId(getAdminForcedTeamId(), teamId)) {
    const forcedName = String(getAdminForcedTeamName() || '').trim();
    if (forcedName) return forcedName;
  }
  return String(fallback || 'このチーム');
}

function setPlatformAdminActionStatus(message, isError = false) {
  if (!els.platformAdminActionStatusText) return;
  els.platformAdminActionStatusText.textContent = String(message || '').trim() || '選択中チームに対して、強制切替 / バックアップ / 復元 / チーム削除を実行できます。';
  els.platformAdminActionStatusText.style.color = isError ? '#ff8f8f' : '';
}

function setPlatformAdminTeamNameStatus(message, isError = false) {
  if (!els.platformAdminTeamNameStatusText) return;
  els.platformAdminTeamNameStatusText.textContent = String(message || '').trim() || '運営者だけが、このチームの表示名を変更できます。';
  els.platformAdminTeamNameStatusText.style.color = isError ? '#ff8f8f' : '';
}

function updatePlatformAdminTeamNameCaches(teamId, nextName) {
  const safeTeamId = String(teamId || '').trim();
  const safeName = String(nextName || '').trim();
  if (!safeTeamId || !safeName) return;

  try {
    platformAdminTeamsCache = (Array.isArray(platformAdminTeamsCache) ? platformAdminTeamsCache : []).map(row =>
      isSameTeamId(row?.id, safeTeamId)
        ? { ...(row || {}), team_name: safeName, name: safeName, updated_at: new Date().toISOString() }
        : row
    );
  } catch (_) {}

  try {
    if (typeof window.cacheDropOffTeamMeta === 'function') {
      const cached = typeof window.getCachedDropOffTeamMeta === 'function' ? window.getCachedDropOffTeamMeta(safeTeamId) : null;
      window.cacheDropOffTeamMeta({ ...(cached || {}), id: safeTeamId, team_name: safeName, name: safeName });
    }
  } catch (_) {}

  try {
    if (isSameTeamId(window.currentWorkspaceTeamId, safeTeamId) || isSameTeamId(window.currentWorkspaceInfo?.id, safeTeamId) || isSameTeamId(getAdminForcedTeamId(), safeTeamId)) {
      setCurrentWorkspaceMetaState({ ...(window.currentWorkspaceInfo || {}), id: safeTeamId, team_name: safeName, name: safeName });
      renderCurrentWorkspaceInfo();
      if (typeof loadCurrentUserProfileForDisplay === 'function') {
        Promise.resolve(loadCurrentUserProfileForDisplay()).catch(() => {});
      }
    }
  } catch (_) {}
}

function reportImportRemovedColumns(scopeLabel, removedColumns) {
  const list = Array.isArray(removedColumns) ? removedColumns.map(v => String(v || '').trim()).filter(Boolean) : [];
  if (!list.length) return;
  const noiseOnly = list.every(name => ['created_by', 'updated_by', 'created_at', 'updated_at'].includes(name));
  if (noiseOnly) return;
  console.warn(`${scopeLabel} import removed columns:`, list);
}

async function stabilizePlatformAdminForcedContext() {
  if (!window.isPlatformAdminUser) return;
  const forcedTeamId = String(getAdminForcedTeamId() || '').trim();
  if (!forcedTeamId) return;

  persistWorkspaceTeamIdLocally(forcedTeamId);
  window.currentWorkspaceTeamId = forcedTeamId;

  const summaryFn = window.getTeamSummaryForAdmin;
  if (typeof summaryFn !== 'function') return;

  try {
    const { data, error } = await summaryFn(forcedTeamId);
    if (error || !data) {
      clearAdminForceTeamStorage();
      setPlatformAdminActionStatus('強制切替先が見つからなかったため、運営者表示を通常モードへ戻しました。', true);
      return;
    }
    window.currentWorkspaceInfo = { ...(window.currentWorkspaceInfo || {}), ...(data || {}) };
    try { window.localStorage.setItem(ADMIN_FORCE_TEAM_NAME_KEY, String(data.team_name || '').trim()); } catch (_) {}
  } catch (_) {}
}

function forceSwitchToTeam(teamId, teamName) {
  const safeTeamId = String(teamId || '').trim();
  if (!window.isPlatformAdminUser || !safeTeamId) return;
  const currentTeamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || '').trim();
  const alreadyForced = getAdminForcedTeamId();
  if (!alreadyForced && currentTeamId && currentTeamId !== safeTeamId) {
    try { window.localStorage.setItem(ADMIN_FORCE_PREV_TEAM_ID_KEY, currentTeamId); } catch (_) {}
    try { window.localStorage.setItem(ADMIN_FORCE_PREV_TEAM_NAME_KEY, String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || '')); } catch (_) {}
  }
  try { window.localStorage.setItem(ADMIN_FORCE_TEAM_ID_KEY, safeTeamId); } catch (_) {}
  try { window.localStorage.setItem(ADMIN_FORCE_TEAM_NAME_KEY, String(teamName || '').trim()); } catch (_) {}
  clearWorkspaceScopedRuntimeCaches();
  persistWorkspaceTeamIdLocally(safeTeamId);
  window.currentWorkspaceTeamId = safeTeamId;
  window.location.href = 'dashboard.html?platform_admin=1&admin_view=1&team_id=' + encodeURIComponent(safeTeamId);
}

function getOriginLocalBackupStorageKey() {
  const teamId = getCurrentWorkspaceTeamIdSync();
  return teamId ? `${ORIGIN_LOCAL_BACKUP_STORAGE_KEY}:${teamId}` : ORIGIN_LOCAL_BACKUP_STORAGE_KEY;
}


function renderSuspendedWorkspaceMode() {
  if (suspendedWorkspaceViewRendered) return;
  suspendedWorkspaceViewRendered = true;

  document.querySelectorAll('.main-tab').forEach(btn => {
    const tabId = String(btn?.dataset?.tab || '');
    if (!tabId) {
      btn.classList.add('hidden');
      return;
    }
    btn.classList.add('hidden');
  });
  if (els.platformAdminTabBtn && window.isPlatformAdminUser) {
    els.platformAdminTabBtn.classList.remove('hidden');
  }
  if (els.adminHeaderBtn) els.adminHeaderBtn.classList.add('hidden');
  if (els.openManualBtn) els.openManualBtn.classList.add('hidden');

  const pageBody = document.querySelector('.page-body');
  if (!pageBody) return;
  const workspace = window.currentWorkspaceInfo || {};
  const teamName = String(workspace?.name || workspace?.team_name || 'このワークスペース');
  const reason = String(workspace?.suspended_reason || '').trim();
  const suspendedAt = workspace?.suspended_at ? formatDateTimeLabel(workspace.suspended_at) : '';

  pageBody.innerHTML = `
    <section class="page-panel active" id="workspaceSuspendedTab">
      <section class="panel-card">
        <h2>このワークスペースは停止中です</h2>
        <p class="soft-text">${escapeHtml(teamName)} は現在、運営者により停止されています。</p>
        ${reason ? `<p class="soft-text">理由: ${escapeHtml(reason)}</p>` : ''}
        ${suspendedAt ? `<p class="soft-text">停止日時: ${escapeHtml(suspendedAt)}</p>` : ''}
        <div class="action-row" style="margin-top:16px;">
          <button id="suspendedLogoutBtn" class="btn danger">ログアウト</button>
        </div>
      </section>
    </section>
  `;
  const logoutBtn = document.getElementById('suspendedLogoutBtn');
  logoutBtn?.addEventListener('click', async () => {
    try {
      if (typeof logout === 'function') await logout();
      else {
        await supabaseClient.auth.signOut();
        window.location.href = 'index.html';
      }
    } catch (e) {
      window.location.href = 'index.html';
    }
  });
}

function isUuidLike(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}
let lastAutoDispatchRunAtMinutes = null;
let simulationSlotHour = null;
let lastSimulationResult = null;
let isRefreshingHybridUI = false;
let suppressSimulationSlotChange = false;
const ENABLE_DISTANCE_REBALANCE = false;
const ENFORCE_AREA_PRIORITY_STRICT = true;
const ENABLE_DISPLAY_GROUP_FORCE_BRANCH = true;

const els = {
  plansTimeAreaMatrix: document.getElementById("plansTimeAreaMatrix"),
  userEmail: document.getElementById("userEmail"),
  userRoleText: document.getElementById("userRoleText"),
  currentTeamText: document.getElementById("currentTeamText"),
  adminHeaderBtn: document.getElementById("adminHeaderBtn"),
  platformAdminTabBtn: document.getElementById("platformAdminTabBtn"),
  platformAdminForceBanner: document.getElementById("platformAdminForceBanner"),
  platformAdminForceBannerText: document.getElementById("platformAdminForceBannerText"),
  platformAdminBackBtn: document.getElementById("platformAdminBackBtn"),
  refreshPlatformAdminTeamsBtn: document.getElementById("refreshPlatformAdminTeamsBtn"),
  platformAdminStatusText: document.getElementById("platformAdminStatusText"),
  platformAdminTeamsTableBody: document.getElementById("platformAdminTeamsTableBody"),
  platformAdminDetailWrap: document.getElementById("platformAdminDetailWrap"),
  platformAdminDetailEmpty: document.getElementById("platformAdminDetailEmpty"),
  platformAdminDetailContent: document.getElementById("platformAdminDetailContent"),
  platformAdminDetailTeamName: document.getElementById("platformAdminDetailTeamName"),
  platformAdminDetailStatus: document.getElementById("platformAdminDetailStatus"),
  platformAdminDetailMembersCount: document.getElementById("platformAdminDetailMembersCount"),
  platformAdminDetailVehiclesCount: document.getElementById("platformAdminDetailVehiclesCount"),
  platformAdminDetailCastsCount: document.getElementById("platformAdminDetailCastsCount"),
  platformAdminDetailOriginsCount: document.getElementById("platformAdminDetailOriginsCount"),
  platformAdminDetailDispatchesCount: document.getElementById("platformAdminDetailDispatchesCount"),
  platformAdminDetailMeta: document.getElementById("platformAdminDetailMeta"),
  platformAdminTeamNameInput: document.getElementById("platformAdminTeamNameInput"),
  savePlatformAdminTeamNameBtn: document.getElementById("savePlatformAdminTeamNameBtn"),
  platformAdminTeamNameStatusText: document.getElementById("platformAdminTeamNameStatusText"),
  platformAdminDetailPlanType: document.getElementById("platformAdminDetailPlanType"),
  platformAdminDetailPlanLimits: document.getElementById("platformAdminDetailPlanLimits"),
  platformAdminDetailPlanFeatures: document.getElementById("platformAdminDetailPlanFeatures"),
  platformAdminSetFreePlanBtn: document.getElementById("platformAdminSetFreePlanBtn"),
  platformAdminSetPaidPlanBtn: document.getElementById("platformAdminSetPaidPlanBtn"),
  platformAdminPlanStatusText: document.getElementById("platformAdminPlanStatusText"),
  platformAdminDetailFlags: document.getElementById("platformAdminDetailFlags"),
  platformAdminMembersTableBody: document.getElementById("platformAdminMembersTableBody"),
  switchPlatformTeamBtn: document.getElementById("switchPlatformTeamBtn"),
  suspendPlatformTeamBtn: document.getElementById("suspendPlatformTeamBtn"),
  resumePlatformTeamBtn: document.getElementById("resumePlatformTeamBtn"),
  exportPlatformTeamBackupBtn: document.getElementById("exportPlatformTeamBackupBtn"),
  importPlatformTeamBackupBtn: document.getElementById("importPlatformTeamBackupBtn"),
  importPlatformTeamBackupInput: document.getElementById("importPlatformTeamBackupInput"),
  deletePlatformTeamBtn: document.getElementById("deletePlatformTeamBtn"),
  platformAdminActionStatusText: document.getElementById("platformAdminActionStatusText"),
  refreshPlatformAdminAnalyticsBtn: document.getElementById("refreshPlatformAdminAnalyticsBtn"),
  platformAdminAnalyticsStatusText: document.getElementById("platformAdminAnalyticsStatusText"),
  platformAdminHomeTodayCount: document.getElementById("platformAdminHomeTodayCount"),
  platformAdminDashboardTodayCount: document.getElementById("platformAdminDashboardTodayCount"),
  platformAdminHome30dCount: document.getElementById("platformAdminHome30dCount"),
  platformAdminDashboard30dCount: document.getElementById("platformAdminDashboard30dCount"),
  platformAdminHomeChart: document.getElementById("platformAdminHomeChart"),
  platformAdminDashboardChart: document.getElementById("platformAdminDashboardChart"),
  platformAdminAnalyticsDailyBody: document.getElementById("platformAdminAnalyticsDailyBody"),
  userManagementSection: document.getElementById("userManagementSection"),
  refreshProfilesBtn: document.getElementById("refreshProfilesBtn"),
  profilesTableWrap: document.getElementById("profilesTableWrap"),
  invitationManagementSection: document.getElementById("invitationManagementSection"),
  inviteEmailInput: document.getElementById("inviteEmailInput"),
  inviteDisplayNameInput: document.getElementById("inviteDisplayNameInput"),
  inviteRoleSelect: document.getElementById("inviteRoleSelect"),
  inviteTeamSelect: document.getElementById("inviteTeamSelect"),
  sendInvitationBtn: document.getElementById("sendInvitationBtn"),
  refreshInvitationsBtn: document.getElementById("refreshInvitationsBtn"),
  inviteStatusText: document.getElementById("inviteStatusText"),
  invitationsTableWrap: document.getElementById("invitationsTableWrap"),
  originManagementSection: document.getElementById("originManagementSection"),
  dataManagementSection: document.getElementById("dataManagementSection"),
  accountManagementSection: document.getElementById("accountManagementSection"),
  ownerDangerZone: document.getElementById("ownerDangerZone"),
  historySection: document.getElementById("historySection"),
  sendLineBtn: document.getElementById("sendLineBtn"),
  originLabelText: document.getElementById("originLabelText"),
  planTypeText: document.getElementById("planTypeText"),
  planMembersText: document.getElementById("planMembersText"),
  planOriginsText: document.getElementById("planOriginsText"),
  planVehiclesText: document.getElementById("planVehiclesText"),
  planCastsText: document.getElementById("planCastsText"),
  planGoogleApiText: document.getElementById("planGoogleApiText"),
  planBackupText: document.getElementById("planBackupText"),
  planCsvText: document.getElementById("planCsvText"),
  planBillingStatusText: document.getElementById("planBillingStatusText"),
  planBillingPeriodText: document.getElementById("planBillingPeriodText"),
  startPaidCheckoutBtn: document.getElementById("startPaidCheckoutBtn"),
  openBillingPortalBtn: document.getElementById("openBillingPortalBtn"),
  refreshPlanInfoBtn: document.getElementById("refreshPlanInfoBtn"),
  planCheckoutStatusText: document.getElementById("planCheckoutStatusText"),
  castEditorWrap: document.getElementById("castEditorWrap"),
  vehicleEditorWrap: document.getElementById("vehicleEditorWrap"),
  originSlotSelect: document.getElementById("originSlotSelect"),
  originNameInput: document.getElementById("originNameInput"),
  originAddressInput: document.getElementById("originAddressInput"),
  originLatLngInput: document.getElementById("originLatLngInput"),
  fetchOriginLatLngBtn: document.getElementById("fetchOriginLatLngBtn"),
  openOriginGoogleMapBtn: document.getElementById("openOriginGoogleMapBtn"),
  saveOriginBtn: document.getElementById("saveOriginBtn"),
  useOriginDraftBtn: document.getElementById("useOriginDraftBtn"),
  cancelOriginEditBtn: document.getElementById("cancelOriginEditBtn"),
  originStatusText: document.getElementById("originStatusText"),
  originSlotsWrap: document.getElementById("originSlotsWrap"),
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
  castDistanceHint: document.getElementById("castDistanceHint"),
  castTravelMinutes: document.getElementById("castTravelMinutes"),
  fetchCastTravelMinutesBtn: document.getElementById("fetchCastTravelMinutesBtn"),
  castAddress: document.getElementById("castAddress"),
  castArea: document.getElementById("castArea"),
  castMemo: document.getElementById("castMemo"),
  castLatLngText: document.getElementById("castLatLngText"),
  castPhone: document.getElementById("castPhone"),
  castGeoStatus: document.getElementById("castGeoStatus"),
  castApiQuotaText: document.getElementById("castApiQuotaText"),
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
  vehicleHomeLatLngText: document.getElementById("vehicleHomeLatLngText"),
  vehicleGeoStatus: document.getElementById("vehicleGeoStatus"),
  vehicleHomeLat: document.getElementById("vehicleHomeLat"),
  vehicleHomeLng: document.getElementById("vehicleHomeLng"),
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
  addFromPlansDialog: document.getElementById("addFromPlansDialog"),
  addFromPlansDateLabel: document.getElementById("addFromPlansDateLabel"),
  addFromPlansCloseBtn: document.getElementById("addFromPlansCloseBtn"),
  addFromPlansEmpty: document.getElementById("addFromPlansEmpty"),
  addFromPlansList: document.getElementById("addFromPlansList"),
  addFromPlansCancelBtn: document.getElementById("addFromPlansCancelBtn"),
  addFromPlansConfirmBtn: document.getElementById("addFromPlansConfirmBtn"),

  historyList: document.getElementById("historyList")
};

const AREA_DIRECTION_MAP = {
  "松戸近郊": "CENTER",
  "葛飾方面": "W",
  "足立方面": "W",
  "江戸川方面": "SW",
  "市川方面": "S",
  "船橋方面": "SE",
  "鎌ヶ谷方面": "SE",
  "我孫子方面": "NE",
  "取手方面": "NE",
  "藤代方面": "E",
  "守谷方面": "E",
  "柏方面": "E",
  "柏の葉方面": "NE",
  "流山方面": "N",
  "野田方面": "N",
  "三郷方面": "NW",
  "八潮方面": "W",
  "草加方面": "NW",
  "吉川方面": "N",
  "越谷方面": "NW",
  "千葉方面": "SE"
};

const DIRECTION_RING = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"];

const HOME_ALLOWED_AREA_MAP = {
  "葛飾方面": ["葛飾方面", "松戸近郊", "足立方面", "江戸川方面", "市川方面", "八潮方面", "三郷方面"],
  "足立方面": ["足立方面", "葛飾方面", "八潮方面", "草加方面", "松戸近郊"],
  "江戸川方面": ["江戸川方面", "葛飾方面", "市川方面", "船橋方面"],
  "市川方面": ["市川方面", "江戸川方面", "船橋方面", "鎌ヶ谷方面", "松戸近郊"],
  "船橋方面": ["船橋方面", "鎌ヶ谷方面", "市川方面", "千葉方面"],
  "鎌ヶ谷方面": ["鎌ヶ谷方面", "船橋方面", "市川方面", "柏方面", "松戸近郊"],
  "我孫子方面": ["我孫子方面", "取手方面", "藤代方面", "守谷方面", "柏方面"],
  "取手方面": ["取手方面", "藤代方面", "我孫子方面", "守谷方面", "柏方面"],
  "藤代方面": ["藤代方面", "取手方面", "守谷方面", "我孫子方面"],
  "守谷方面": ["守谷方面", "取手方面", "藤代方面", "我孫子方面", "柏方面"],
  "柏方面": ["柏方面", "柏の葉方面", "流山方面", "我孫子方面", "鎌ヶ谷方面", "松戸近郊"],
  "柏の葉方面": ["柏の葉方面", "柏方面", "流山方面", "野田方面"],
  "流山方面": ["流山方面", "柏方面", "野田方面", "三郷方面", "吉川方面", "松戸近郊"],
  "野田方面": ["野田方面", "流山方面", "柏方面", "吉川方面", "柏の葉方面"],
  "三郷方面": ["三郷方面", "八潮方面", "松戸近郊", "吉川方面", "流山方面"],
  "八潮方面": ["八潮方面", "三郷方面", "足立方面", "葛飾方面", "草加方面"],
  "草加方面": ["草加方面", "足立方面", "八潮方面", "越谷方面"],
  "吉川方面": ["吉川方面", "流山方面", "野田方面", "三郷方面", "越谷方面"],
  "越谷方面": ["越谷方面", "草加方面", "吉川方面"],
  "松戸近郊": ["松戸近郊", "葛飾方面", "市川方面", "柏方面", "三郷方面", "足立方面", "流山方面"],
  "千葉方面": ["千葉方面", "船橋方面"]
};

const ROUTE_FLOW_MAP = {
  "松戸近郊": ["流山方面", "吉川方面", "三郷方面", "柏方面", "我孫子方面", "市川方面", "葛飾方面"],
  "流山方面": ["松戸近郊", "吉川方面", "野田方面", "柏方面", "柏の葉方面"],
  "吉川方面": ["松戸近郊", "流山方面", "三郷方面", "野田方面"],
  "三郷方面": ["松戸近郊", "吉川方面", "八潮方面", "流山方面"],
  "柏方面": ["松戸近郊", "我孫子方面", "流山方面", "柏の葉方面", "取手方面"],
  "柏の葉方面": ["柏方面", "流山方面", "野田方面"],
  "我孫子方面": ["松戸近郊", "柏方面", "取手方面", "藤代方面", "守谷方面"],
  "取手方面": ["我孫子方面", "藤代方面", "守谷方面", "柏方面"],
  "藤代方面": ["我孫子方面", "取手方面", "守谷方面"],
  "守谷方面": ["我孫子方面", "取手方面", "藤代方面", "柏方面"],
  "市川方面": ["松戸近郊", "鎌ヶ谷方面", "船橋方面", "葛飾方面", "江戸川方面"],
  "鎌ヶ谷方面": ["市川方面", "船橋方面", "柏方面", "松戸近郊"],
  "船橋方面": ["市川方面", "鎌ヶ谷方面", "千葉方面"],
  "葛飾方面": ["松戸近郊", "足立方面", "江戸川方面", "市川方面", "三郷方面"],
  "足立方面": ["葛飾方面", "松戸近郊", "八潮方面"],
  "江戸川方面": ["葛飾方面", "市川方面", "船橋方面"],
  "野田方面": ["流山方面", "吉川方面", "柏方面", "柏の葉方面"],
  "八潮方面": ["葛飾方面", "足立方面", "三郷方面", "草加方面"],
  "草加方面": ["八潮方面", "足立方面", "越谷方面"],
  "越谷方面": ["草加方面", "吉川方面"],
  "千葉方面": ["船橋方面", "市川方面"]
};

const ROUTE_CONTINUITY_HINTS = [
  { dir: "N", patterns: ["新松戸", "馬橋", "北松戸", "北小金", "小金", "幸谷", "新八柱", "八柱", "常盤平", "みのり台"] },
  { dir: "NE", patterns: ["南流山", "流山", "柏", "柏の葉", "我孫子", "取手", "藤代", "守谷"] },
  { dir: "W", patterns: ["葛飾", "足立", "綾瀬", "亀有", "金町", "八潮"] },
  { dir: "SW", patterns: ["江戸川"] },
  { dir: "S", patterns: ["市川", "本八幡", "妙典", "行徳"] },
  { dir: "SE", patterns: ["船橋", "習志野", "鎌ヶ谷", "鎌ケ谷"] },
  { dir: "NW", patterns: ["三郷", "吉川", "越谷", "草加"] }
];

const HARD_ROUTE_MIX_GROUPS = {
  NE: ["取手方面", "藤代方面", "守谷方面", "我孫子方面", "柏方面", "柏の葉方面"],
  N: ["流山方面", "野田方面", "吉川方面"],
  NW: ["三郷方面", "八潮方面", "草加方面", "越谷方面", "足立方面", "葛飾方面"],
  S: ["市川方面", "江戸川方面"],
  SE: ["船橋方面", "鎌ヶ谷方面", "千葉方面"],
  CENTER: ["松戸近郊"]
};

function getHardRouteMixGroup(area) {
  const canonical = getCanonicalArea(area);
  for (const [group, areas] of Object.entries(HARD_ROUTE_MIX_GROUPS)) {
    if (areas.includes(canonical)) return group;
  }
  return "";
}

function isHardReverseMixForRoute(areaA, areaB) {
  const a = getCanonicalArea(areaA);
  const b = getCanonicalArea(areaB);
  if (!a || !b || a === b) return false;

  const groupA = getHardRouteMixGroup(a);
  const groupB = getHardRouteMixGroup(b);
  const dirA = getAreaTravelDirection(a);
  const dirB = getAreaTravelDirection(b);
  const dirDistance = getDirectionDistanceByKey(dirA, dirB);
  const affinity = getAreaAffinityScore(a, b);

  const neLike = new Set(["NE", "N"]);
  const southLike = new Set(["S", "SE", "SW"]);
  const westLike = new Set(["W", "NW"]);

  if ((neLike.has(groupA) && southLike.has(groupB)) || (neLike.has(groupB) && southLike.has(groupA))) {
    return true;
  }
  if ((groupA === "NE" && westLike.has(groupB)) || (groupB === "NE" && westLike.has(groupA))) {
    return affinity <= 28;
  }
  if ((groupA === "N" && southLike.has(groupB)) || (groupB === "N" && southLike.has(groupA))) {
    return affinity <= 36;
  }
  if (dirDistance >= 4 && affinity <= 28) return true;
  if (dirDistance >= 3 && affinity <= 18 && ![a, b].includes("松戸近郊")) return true;
  return false;
}

function getAreaTravelDirection(area) {
  const raw = normalizeAreaLabel(area);
  if (!raw || raw === "無し") return "";

  for (const hint of ROUTE_CONTINUITY_HINTS) {
    if (hint.patterns.some(pattern => raw.includes(pattern))) return hint.dir;
  }

  return getAreaDirectionCluster(raw);
}

function getDirectionDistanceByKey(dirA, dirB) {
  if (!dirA || !dirB) return 99;
  if (dirA === dirB) return 0;
  if (dirA === "CENTER" || dirB === "CENTER") return 1;
  const indexA = DIRECTION_RING.indexOf(dirA);
  const indexB = DIRECTION_RING.indexOf(dirB);
  if (indexA < 0 || indexB < 0) return 99;
  const raw = Math.abs(indexA - indexB);
  return Math.min(raw, DIRECTION_RING.length - raw);
}

function isGatewayNearArea(area) {
  const raw = normalizeAreaLabel(area);
  if (!raw || raw === "無し") return false;
  return ["松戸", "新松戸", "馬橋", "北松戸", "北小金", "小金", "八柱", "常盤平", "みのり台"].some(keyword => raw.includes(keyword));
}


function getOnTheWayCompatibility(areaA, areaB) {
  const a = normalizeAreaLabel(areaA);
  const b = normalizeAreaLabel(areaB);
  if (!a || !b || a === "無し" || b === "無し") return 0;

  const corridor = ["柏", "我孫子", "取手", "藤代", "牛久", "ひたち野うしく"];
  const idxA = corridor.findIndex(k => a.includes(k));
  const idxB = corridor.findIndex(k => b.includes(k));

  if (idxA === -1 || idxB === -1) return 0;

  const diff = Math.abs(idxA - idxB);
  if (diff === 0) return 80;
  if (diff === 1) return 100;
  if (diff === 2) return 60;
  return 0;
}

function getPairRouteContinuityPenalty(areaA, areaB) {
  const dirA = getAreaTravelDirection(areaA);
  const dirB = getAreaTravelDirection(areaB);
  const distance = getDirectionDistanceByKey(dirA, dirB);
  const routeFlow = getRouteFlowCompatibilityBetweenAreas(areaA, areaB);
  const affinity = getAreaAffinityScore(areaA, areaB);
  const groupA = getHardRouteMixGroup(areaA);
  const groupB = getHardRouteMixGroup(areaB);
  const sameBroadGroup = groupA && groupA === groupB;

  let penalty = 0;

  if (distance === 0) penalty += 0;
  else if (distance === 1) penalty += sameBroadGroup ? 6 : 18;
  else if (distance === 2) penalty += sameBroadGroup ? 28 : 120;
  else if (distance === 3) penalty += sameBroadGroup ? 88 : 300;
  else if (distance >= 4 && distance < 99) penalty += sameBroadGroup ? 160 : 560;

  if (routeFlow <= 0) penalty += sameBroadGroup ? 8 : 42;
  else if (routeFlow < 40) penalty += sameBroadGroup ? 4 : 22;

  if (affinity >= 88) penalty -= 72;
  else if (affinity >= 72) penalty -= 40;
  else if (affinity >= 54) penalty -= 14;

  const gatewayA = isGatewayNearArea(areaA);
  const gatewayB = isGatewayNearArea(areaB);
  if (gatewayA || gatewayB) {
    if (distance >= 2) penalty += 140;
  }

  const pair = [getCanonicalArea(areaA), getCanonicalArea(areaB)].filter(Boolean).sort().join("__");
  const strongPairs = new Set([
    "吉川方面__松戸近郊",
    "三郷方面__松戸近郊",
    "我孫子方面__松戸近郊",
    "柏方面__松戸近郊",
    "流山方面__松戸近郊"
  ]);
  if (strongPairs.has(pair)) penalty -= 48;

  if ((gatewayA || gatewayB) && ((dirA === "N" && ["W", "SW", "S"].includes(dirB)) || (dirB === "N" && ["W", "SW", "S"].includes(dirA)))) {
    penalty += 140;
  }

  if (isHardReverseMixForRoute(areaA, areaB)) {
    penalty += 900;
  } else if (sameBroadGroup && affinity >= 72) {
    penalty -= 36;
  }

  const onTheWay = getOnTheWayCompatibility(areaA, areaB);
  if (onTheWay >= 80) penalty -= 120;
  else if (onTheWay >= 50) penalty -= 70;

  return Math.max(0, penalty);
}

function isStrictSameDirectionArea(targetArea, existingAreas = []) {
  const target = normalizeAreaLabel(targetArea);
  const areas = Array.isArray(existingAreas) ? existingAreas.map(normalizeAreaLabel).filter(Boolean) : [];
  if (!target || target === "無し" || !areas.length) return true;
  const targetGroup = getAreaDisplayGroup(target);
  return areas.every(area => {
    const group = getAreaDisplayGroup(area);
    if (group && targetGroup && group !== targetGroup) return false;
    return !hasHardReverseMix(target, [area]);
  });
}

function isLastRunHardAreaConstraintSatisfied(targetArea, existingAreas = [], homeArea = "") {
  const target = normalizeAreaLabel(targetArea);
  const home = normalizeAreaLabel(homeArea || "");
  if (!target || target === "無し") return false;
  if (home && isHardReverseForHome(target, home)) return false;
  if (!isStrictSameDirectionArea(target, existingAreas)) return false;
  const strict = getStrictHomeCompatibilityScore(target, home);
  const direction = getDirectionAffinityScore(target, home);
  // Last run should prioritize going toward driver's home direction.
  return strict >= 52 || direction >= 24;
}

function getRouteFlowSortWeight(area) {
  const canonical = getCanonicalArea(area);
  if (["守谷方面", "藤代方面", "取手方面", "我孫子方面", "千葉方面"].includes(canonical)) return 100;
  if (["吉川方面", "船橋方面", "野田方面", "柏の葉方面"].includes(canonical)) return 85;
  if (["流山方面", "柏方面", "市川方面", "鎌ヶ谷方面", "三郷方面", "足立方面", "葛飾方面"].includes(canonical)) return 70;
  if (canonical === "松戸近郊") return 40;
  return 55;
}

function sortClustersForRouteFlow(clusters) {
  return [...clusters].sort((a, b) => {
    if (a.hour !== b.hour) return a.hour - b.hour;
    const aw = getRouteFlowSortWeight(a.area);
    const bw = getRouteFlowSortWeight(b.area);
    if (bw !== aw) return bw - aw;
    if (b.count !== a.count) return b.count - a.count;
    return b.totalDistance - a.totalDistance;
  });
}

function getAssignmentAreasByVehicleHour(assignments, itemsById, vehicleId, hour, excludeItemId = null) {
  return assignments
    .filter(a => Number(a.vehicle_id) === Number(vehicleId) && Number(a.actual_hour) === Number(hour) && Number(a.item_id) !== Number(excludeItemId || -1))
    .map(a => normalizeAreaLabel(itemsById.get(Number(a.item_id))?.destination_area || ""))
    .filter(Boolean);
}

function getAreaAffinityScore(areaA, areaB) {
  const a = getCanonicalArea(areaA);
  const b = getCanonicalArea(areaB);
  if (!a || !b) return 0;
  if (a === b) return 100;
  return Number(AREA_AFFINITY_MAP[a]?.[b] || AREA_AFFINITY_MAP[b]?.[a] || 0);
}

function getAreaDirectionCluster(area) {
  const canonical = getCanonicalArea(area);
  return AREA_DIRECTION_MAP[canonical] || "";
}

function getDirectionDistance(areaA, areaB) {
  const dirA = getAreaDirectionCluster(areaA);
  const dirB = getAreaDirectionCluster(areaB);
  if (!dirA || !dirB) return 99;
  if (dirA === "CENTER" || dirB === "CENTER") return 1;
  const indexA = DIRECTION_RING.indexOf(dirA);
  const indexB = DIRECTION_RING.indexOf(dirB);
  if (indexA < 0 || indexB < 0) return 99;
  const raw = Math.abs(indexA - indexB);
  return Math.min(raw, DIRECTION_RING.length - raw);
}

function getDirectionAffinityScore(areaA, areaB) {
  const distance = getDirectionDistance(areaA, areaB);
  if (distance === 99) return 0;
  if (distance === 0) return 100;
  if (distance === 1) return 72;
  if (distance === 2) return 28;
  if (distance === 3) return -38;
  return -95;
}

function getStrictHomeCompatibilityScore(clusterArea, homeArea) {
  const cluster = getCanonicalArea(clusterArea);
  const home = getCanonicalArea(homeArea);
  if (!cluster || !home) return 0;
  if (cluster === home) return 100;
  const allowed = HOME_ALLOWED_AREA_MAP[home] || [];
  if (allowed.includes(cluster)) return 78;
  const directionScore = getDirectionAffinityScore(cluster, home);
  if (directionScore >= 72) return 52;
  if (directionScore >= 28) return 18;
  return 0;
}

function isHardReverseForHome(clusterArea, homeArea) {
  const affinity = getAreaAffinityScore(clusterArea, homeArea);
  const directionScore = getDirectionAffinityScore(clusterArea, homeArea);
  const strictScore = getStrictHomeCompatibilityScore(clusterArea, homeArea);
  if (directionScore <= -95) return true;
  if (directionScore <= -38 && strictScore === 0) return true;
  if (affinity <= 25 && strictScore === 0) return true;
  return false;
}

function getLastTripHomePriorityWeight(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster) {
  const affinity = getAreaAffinityScore(clusterArea, homeArea);
  const directionScore = getDirectionAffinityScore(clusterArea, homeArea);
  const strictScore = getStrictHomeCompatibilityScore(clusterArea, homeArea);

  let weight = affinity * 1.1 + Math.max(directionScore, 0) * 1.15 + strictScore * 1.35;

  if (directionScore < 0) {
    weight += directionScore * 2.4;
  }

  if (isHardReverseForHome(clusterArea, homeArea)) {
    weight -= isLastRun ? 520 : (isDefaultLastHourCluster ? 320 : 90);
  }

  if (isLastRun) return weight * 2.8;
  if (isDefaultLastHourCluster) return weight * 2.2;
  return weight * 0.45;
}

function getVehicleMonthlyStatsMap(reportRows, targetMonth) {
  const map = new Map();

  function resolveVehicleIdFromReport(row) {
    const directId = Number(row?.vehicle_id || row?.vehicles?.id || 0);
    if (directId > 0) return directId;

    const reportDriver = String(row?.driver_name || row?.vehicles?.driver_name || "").trim();
    const reportPlate = String(row?.plate_number || row?.vehicles?.plate_number || "").trim();

    const matchedVehicle = (Array.isArray(allVehiclesCache) ? allVehiclesCache : []).find(vehicle => {
      const vehicleDriver = String(vehicle?.driver_name || "").trim();
      const vehiclePlate = String(vehicle?.plate_number || "").trim();
      return (reportDriver && vehicleDriver && reportDriver === vehicleDriver) ||
             (reportPlate && vehiclePlate && reportPlate === vehiclePlate);
    });

    return Number(matchedVehicle?.id || 0);
  }

  (Array.isArray(reportRows) ? reportRows : []).forEach(row => {
    if (getMonthKey(row?.report_date) !== targetMonth) return;

    const vehicleId = resolveVehicleIdFromReport(row);
    if (!vehicleId) return;

    const prev = map.get(vehicleId) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0,
      byDate: new Map()
    };

    const dateKey = String(row?.report_date || "").trim();
    const distance = Number(row?.distance_km || 0);
    prev.totalDistance += distance;
    if (dateKey) {
      prev.byDate.set(dateKey, Number((prev.byDate.get(dateKey) || 0) + distance));
    }
    prev.workedDays = [...prev.byDate.values()].filter(value => Number(value || 0) > 0).length;
    prev.avgDistance = prev.workedDays > 0 ? prev.totalDistance / prev.workedDays : 0;
    map.set(vehicleId, prev);
  });

  map.forEach(stats => {
    stats.totalDistance = Number(Number(stats.totalDistance || 0).toFixed(1));
    stats.avgDistance = Number(Number(stats.avgDistance || 0).toFixed(1));
    delete stats.byDate;
  });

  return map;
}



function getUnifiedMonthlyUiStatsMap(reportRows = currentDailyReportsCache, baseDate) {
  const targetDate = baseDate || (els?.dispatchDate?.value || todayStr());
  const startDate = getMonthStartStr(targetDate);
  const endDate = getMonthEndStr(targetDate);
  const normalizedRows = normalizeMileageExportRows(
    (Array.isArray(reportRows) ? reportRows : []).filter(row => {
      const d = String(row?.report_date || "");
      return d && d >= startDate && d <= endDate;
    })
  );

  const calendar = buildMileageCalendarRows(normalizedRows, startDate, endDate);
  const map = new Map();

  const vehicles = Array.isArray(allVehiclesCache) ? allVehiclesCache : [];
  const normalizeText = value => String(value || "").trim();
  const findVehicleId = entry => {
    const directId = Number(entry?.vehicle_id || 0);
    if (directId > 0) return directId;

    const driver = normalizeText(entry?.driver_name);
    const plate = normalizeText(entry?.plate_number);

    const exact = vehicles.find(vehicle => {
      const vehicleDriver = normalizeText(vehicle?.driver_name);
      const vehiclePlate = normalizeText(vehicle?.plate_number);
      return (plate && vehiclePlate && plate === vehiclePlate) ||
             (driver && vehicleDriver && driver === vehicleDriver);
    });
    return Number(exact?.id || 0);
  };

  (calendar?.drivers || []).forEach(entry => {
    const vehicleId = findVehicleId(entry);
    if (!vehicleId) return;
    map.set(vehicleId, {
      totalDistance: Number(Number(entry.total_distance_km || 0).toFixed(1)),
      workedDays: Number(entry.worked_days || 0),
      avgDistance: Number(Number(entry.avg_distance_km || 0).toFixed(1))
    });
  });

  return map;
}

function buildMonthlyDistanceMapForCurrentMonth() {
  try {
    const targetDate = typeof getSelectedDispatchDate === "function"
      ? getSelectedDispatchDate()
      : (new Date()).toISOString().slice(0, 10);

    if (typeof window.__dropoffGetMonthlyStatsMap === "function") {
      const latestMap = window.__dropoffGetMonthlyStatsMap(targetDate);
      if (latestMap instanceof Map && latestMap.size > 0) {
        return latestMap;
      }
    }

    return getUnifiedMonthlyUiStatsMap(
      Array.isArray(currentDailyReportsCache) ? currentDailyReportsCache : [],
      targetDate
    );
  } catch (_) {
    return new Map();
  }
}

function optimizeAssignmentsByRouteFlow(assignments, items, vehicles) {
  return Array.isArray(assignments) ? assignments : [];
}

function applyManualLastVehicleToAssignments(assignments, vehicles) {
  return Array.isArray(assignments) ? assignments : [];
}

function joinManualLastVehicleToAssignments(assignments, vehicles) {
  return applyManualLastVehicleToAssignments(assignments, vehicles);
}

function normalizeDispatchEntityId(value) {
  const raw = String(value ?? "").trim();
  if (!raw || raw === "null" || raw === "undefined") return "";
  return raw;
}

function sameDispatchEntityId(a, b) {
  const aa = normalizeDispatchEntityId(a);
  const bb = normalizeDispatchEntityId(b);
  return !!aa && !!bb && aa === bb;
}

function resolveDispatchActionId(rawId, fallbackSelectors = []) {
  const direct = normalizeDispatchEntityId(rawId);
  if (direct && direct !== "NaN") return direct;

  const selectors = Array.isArray(fallbackSelectors) ? fallbackSelectors.filter(Boolean) : [];
  const eventTarget = window?.event?.target;

  const candidates = [];
  if (eventTarget?.closest) {
    for (const selector of selectors) {
      const found = eventTarget.closest(selector);
      if (found?.dataset?.id) candidates.push(found.dataset.id);
    }
    const anyWithId = eventTarget.closest("[data-id]");
    if (anyWithId?.dataset?.id) candidates.push(anyWithId.dataset.id);
  }

  const active = document?.activeElement;
  if (active?.dataset?.id) candidates.push(active.dataset.id);

  for (const value of candidates) {
    const normalized = normalizeDispatchEntityId(value);
    if (normalized && normalized !== "NaN") return normalized;
  }

  return "";
}

function resolveVehicleCloudRowId(rawVehicleId) {
  const direct = normalizeDispatchEntityId(rawVehicleId);
  if (direct && /[a-f0-9]{8}-[a-f0-9]{4}-[1-5][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}/i.test(direct)) {
    return direct;
  }

  const numericVehicleId = Number(rawVehicleId || 0);
  if (!(numericVehicleId > 0)) return direct || "";

  const matchedVehicle = (Array.isArray(allVehiclesCache) ? allVehiclesCache : []).find(vehicle => Number(vehicle?.id || 0) === numericVehicleId);
  return normalizeDispatchEntityId(matchedVehicle?.cloud_row_id || matchedVehicle?.db_id || direct || "");
}

function resolveVehicleLocalNumericId(rawVehicleId) {
  const direct = normalizeDispatchEntityId(rawVehicleId);
  const directNumber = Number(direct || 0);
  if (directNumber > 0) return directNumber;

  const matchedVehicle = (Array.isArray(allVehiclesCache) ? allVehiclesCache : []).find(vehicle => {
    const localId = Number(vehicle?.id || 0);
    if (!(localId > 0)) return false;
    return sameDispatchEntityId(vehicle?.cloud_row_id, direct) ||
           sameDispatchEntityId(vehicle?.db_id, direct) ||
           sameDispatchEntityId(vehicle?.id, direct);
  });

  return Number(matchedVehicle?.id || 0);
}

function createDispatchVehicleBridge(vehicles = []) {
  const localToCloud = new Map();
  const cloudToLocal = new Map();
  const nameToLocal = new Map();
  const plateToLocal = new Map();

  (Array.isArray(vehicles) ? vehicles : []).forEach(vehicle => {
    const localId = Number(vehicle?.id || 0);
    if (!(localId > 0)) return;
    const cloudId = normalizeDispatchEntityId(vehicle?.cloud_row_id || vehicle?.db_id || '');
    if (cloudId) {
      localToCloud.set(String(localId), cloudId);
      cloudToLocal.set(cloudId, localId);
    }
    const driverName = String(vehicle?.driver_name || '').trim();
    if (driverName) nameToLocal.set(driverName, localId);
    const plate = String(vehicle?.plate_number || '').trim();
    if (plate) plateToLocal.set(plate, localId);
  });

  return {
    resolveCloud(rawVehicleId) {
      const direct = normalizeDispatchEntityId(rawVehicleId);
      if (direct && /[a-f0-9]{8}-[a-f0-9]{4}-[1-5][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}/i.test(direct)) {
        return direct;
      }
      const directNumber = Number(direct || 0);
      if (directNumber > 0) {
        return normalizeDispatchEntityId(localToCloud.get(String(directNumber)) || resolveVehicleCloudRowId(directNumber));
      }
      return normalizeDispatchEntityId(resolveVehicleCloudRowId(direct));
    },
    resolveLocal(rawVehicleId, fallbackName = '', fallbackPlate = '') {
      const direct = normalizeDispatchEntityId(rawVehicleId);
      const directNumber = Number(direct || 0);
      if (directNumber > 0) return directNumber;
      if (direct && cloudToLocal.has(direct)) return Number(cloudToLocal.get(direct) || 0);
      const driverName = String(fallbackName || '').trim();
      if (driverName && nameToLocal.has(driverName)) return Number(nameToLocal.get(driverName) || 0);
      const plate = String(fallbackPlate || '').trim();
      if (plate && plateToLocal.has(plate)) return Number(plateToLocal.get(plate) || 0);
      return resolveVehicleLocalNumericId(direct);
    },
    snapshot() {
      return {
        localToCloud: Object.fromEntries(localToCloud),
        cloudToLocal: Object.fromEntries(cloudToLocal)
      };
    }
  };
}

function sameVehicleAssignmentId(leftVehicleId, rightVehicleId) {
  return resolveVehicleLocalNumericId(leftVehicleId) > 0 &&
         resolveVehicleLocalNumericId(leftVehicleId) === resolveVehicleLocalNumericId(rightVehicleId);
}

function prepareActualRowsForDispatchCore(sourceRows = []) {
  const preparedRows = [];
  const sourceIdByTempId = new Map();

  (Array.isArray(sourceRows) ? sourceRows : []).forEach((row, index) => {
    const sourceId = normalizeDispatchEntityId(row?.id);
    if (!sourceId) return;

    const hour = Number(row?.actual_hour ?? row?.dispatch_hour ?? row?.plan_hour ?? 0);
    const tempId = `actual-${hour}-${index + 1}-${sourceId.slice(0, 8)}`;

    sourceIdByTempId.set(normalizeDispatchEntityId(tempId), sourceId);
    preparedRows.push({
      ...row,
      id: tempId,
      source_dispatch_id: sourceId
    });
  });

  return { rows: preparedRows, sourceIdByTempId };
}

function remapAutoDispatchAssignments(assignments = [], sourceIdByTempId = new Map(), vehicleBridge = null) {
  return (Array.isArray(assignments) ? assignments : []).map(assignment => {
    const tempItemId = normalizeDispatchEntityId(assignment?.item_id);
    const resolvedItemId = sourceIdByTempId.get(tempItemId) || tempItemId;
    const resolvedVehicleId = normalizeDispatchEntityId(
      vehicleBridge?.resolveCloud?.(assignment?.vehicle_id) || resolveVehicleCloudRowId(assignment?.vehicle_id)
    );
    const resolvedLocalVehicleId = Number(
      vehicleBridge?.resolveLocal?.(resolvedVehicleId || assignment?.vehicle_id, assignment?.driver_name, assignment?.plate_number) || 0
    );

    return {
      ...assignment,
      __core_vehicle_id: assignment?.vehicle_id,
      __resolved_local_vehicle_id: resolvedLocalVehicleId,
      item_id: normalizeDispatchEntityId(resolvedItemId),
      vehicle_id: normalizeDispatchEntityId(resolvedVehicleId || assignment?.vehicle_id)
    };
  });
}

function clearLinkedCastSelection(input) {
  if (!input?.dataset) return;
  delete input.dataset.selectedCastId;
}

function setLinkedCastSelection(input, cast) {
  if (!input || !cast) return;
  const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
  if (!castId) return;
  input.dataset.selectedCastId = castId;
  input.value = String(cast.name || '').trim();
}

function resolveLinkedCastFromInput(input, getCandidates) {
  if (!input) return null;
  const casts = typeof getCandidates === 'function' ? getCandidates() : (Array.isArray(allCastsCache) ? allCastsCache : []);
  const value = String(input.value || '').trim();
  const selectedId = normalizeDispatchEntityId(input.dataset?.selectedCastId);

  if (selectedId) {
    const selected = casts.find(c => sameDispatchEntityId(c.id || c.cast_id, selectedId));
    if (selected) {
      const lowered = value.toLowerCase();
      const selectedName = String(selected.name || '').trim().toLowerCase();
      const selectedComposite = getCastSearchText(selected).trim().toLowerCase();
      if (!value || lowered === selectedName || lowered === selectedComposite || selectedComposite.startsWith(lowered) || selectedName.startsWith(lowered)) {
        return selected;
      }
    }
  }

  const matched = findCastByInputValue(value);
  if (!matched) {
    clearLinkedCastSelection(input);
    return null;
  }

  const castId = normalizeDispatchEntityId(matched.id || matched.cast_id);
  const withinCandidates = casts.find(c => sameDispatchEntityId(c.id || c.cast_id, castId));
  if (!withinCandidates) {
    clearLinkedCastSelection(input);
    return null;
  }

  input.dataset.selectedCastId = castId;
  return withinCandidates;
}

function getDoneCastIdsInActuals() {
  const ids = new Set();
  currentActualsCache.forEach(item => {
    const castId = normalizeDispatchEntityId(item.cast_id);
    if (castId && normalizeStatus(item.status) === "done") {
      ids.add(castId);
    }
  });
  return ids;
}

function getPlannedCastIds() {
  const ids = new Set();
  currentPlansCache.forEach(plan => {
    const castId = normalizeDispatchEntityId(plan.cast_id);
    if (!castId) return;
    if (["planned", "assigned", "done", "cancel"].includes(plan.status)) {
      ids.add(castId);
    }
  });
  return ids;
}

function getRemainingPlannedCastIds(dateStr) {
  const ids = new Set();

  currentPlansCache.forEach(plan => {
    if (plan.plan_date !== dateStr) return;
    const castId = normalizeDispatchEntityId(plan.cast_id);
    if (!castId) return;
    const status = String(plan.status || "");
    if (status === "done" || status === "cancel") return;
    ids.add(castId);
  });

  currentActualsCache.forEach(item => {
    const status = normalizeStatus(item.status);
    const castId = normalizeDispatchEntityId(item.cast_id);
    if (!castId) return;
    if (status === "done") {
      ids.delete(castId);
    }
  });

  return ids;
}

function isLastClusterOfTheDay(cluster, dateStr) {
  const remainingIds = getRemainingPlannedCastIds(dateStr);
  cluster.items.forEach(item => {
    const castId = normalizeDispatchEntityId(item.cast_id);
    if (castId) remainingIds.delete(castId);
  });
  return remainingIds.size === 0;
}

function openManual() {
  window.open("manual.html", "_blank");
}


function normalizeAppRole(role) {
  const value = String(role || '').trim().toLowerCase();
  if (value === 'owner' || value === 'admin') return value;
  return 'user';
}

function getRoleLabel(role) {
  const normalized = normalizeAppRole(role);
  if (normalized === 'owner') return 'オーナー';
  if (normalized === 'admin') return '管理者';
  return '利用者';
}

function getCurrentAppRole() {
  return normalizeAppRole(currentUserProfile?.role);
}

function isOwnerUser() {
  return getCurrentAppRole() === 'owner';
}

function isManagerUser() {
  const role = getCurrentAppRole();
  return role === 'owner' || role === 'admin';
}

function isReadonlyUserRole() {
  return getCurrentAppRole() === 'user';
}

function getCurrentWorkspaceTeamIdForUi() {
  return String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
}

function setCurrentWorkspaceMetaState(meta = {}) {
  const teamId = String(meta?.id || getCurrentWorkspaceTeamIdForUi() || '').trim();
  const cachedMeta = teamId && typeof window.getCachedDropOffTeamMeta === 'function' ? window.getCachedDropOffTeamMeta(teamId) : null;
  const teamName = String(meta?.name || meta?.team_name || cachedMeta?.name || cachedMeta?.team_name || '').trim();
  const currentOriginSlot = normalizeOriginSlotNo(meta?.current_origin_slot ?? meta?.currentOriginSlot ?? meta?.active_origin_slot ?? cachedMeta?.current_origin_slot);
  try {
    window.currentWorkspaceInfo = {
      ...(window.currentWorkspaceInfo || {}),
      ...(cachedMeta && typeof cachedMeta === 'object' ? cachedMeta : {}),
      ...(meta && typeof meta === 'object' ? meta : {}),
      ...(teamId ? { id: teamId } : {}),
      ...(teamName ? { name: teamName, team_name: teamName } : {}),
      current_origin_slot: currentOriginSlot
    };
  } catch (_) {}
  if (teamId) {
    try { window.currentWorkspaceTeamId = teamId; } catch (_) {}
  }
  if (teamName) {
    try { window.__DROP_OFF_CURRENT_TEAM_LABEL__ = teamName; } catch (_) {}
  }
  try {
    if (teamId && typeof window.cacheDropOffTeamMeta === 'function') {
      window.cacheDropOffTeamMeta({ id: teamId, name: teamName || cachedMeta?.name || '', current_origin_slot: currentOriginSlot });
    }
  } catch (_) {}
  return window.currentWorkspaceInfo || {};
}

function getCurrentWorkspaceOriginSlotFromState() {
  return normalizeOriginSlotNo(window.currentWorkspaceInfo?.current_origin_slot);
}

function getCurrentWorkspaceTeamLabel() {
  const info = window.currentWorkspaceInfo || {};
  const teamId = getCurrentWorkspaceTeamIdForUi();
  const cachedMeta = typeof window.getCachedDropOffTeamMeta === 'function' ? window.getCachedDropOffTeamMeta(teamId) : null;
  const profileTeamName = String(window.currentUserProfile?.team_name || '').trim();
  const label = String(
    info.name
    || info.team_name
    || cachedMeta?.name
    || cachedMeta?.team_name
    || profileTeamName
    || window.__DROP_OFF_CURRENT_TEAM_LABEL__
    || getAdminForcedTeamName()
    || window.currentWorkspaceTeamId
    || getCurrentWorkspaceTeamIdSync()
    || ''
  ).trim();
  return label || '-';
}

async function loadCurrentWorkspaceTeamMeta(force = false) {
  const teamId = getCurrentWorkspaceTeamIdForUi();
  if (!teamId) return null;

  const cachedLabel = String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || '').trim();
  const cachedOriginSlot = getCurrentWorkspaceOriginSlotFromState();
  if (!force && cachedLabel && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(cachedLabel) && cachedOriginSlot) {
    return window.currentWorkspaceInfo || null;
  }

  try {
    if (typeof window.getDropOffTeamMeta === 'function') {
      const { data, error } = await window.getDropOffTeamMeta(teamId);
      if (error) console.warn('getDropOffTeamMeta warning:', error);
      if (data) return setCurrentWorkspaceMetaState(data);
    }
  } catch (error) {
    console.warn('loadCurrentWorkspaceTeamMeta failed:', error);
  }

  try {
    if (typeof loadDropOffTeamsCloud === 'function') {
      const rows = await loadDropOffTeamsCloud();
      const matched = Array.isArray(rows) ? rows.find(row => String(row?.id || '').trim() === teamId) : null;
      if (matched) return setCurrentWorkspaceMetaState(matched);
    }
  } catch (error) {
    console.warn('loadDropOffTeamsCloud for current team meta failed:', error);
  }

  return window.currentWorkspaceInfo || null;
}

async function resolveCurrentWorkspaceTeamLabelAsync() {
  const cached = getCurrentWorkspaceTeamLabel();
  if (cached && cached !== '-' && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(cached)) {
    return cached;
  }

  const meta = await loadCurrentWorkspaceTeamMeta(true);
  const resolved = String(meta?.name || meta?.team_name || '').trim();
  return resolved || cached || getCurrentWorkspaceTeamIdForUi() || '-';
}

function renderCurrentWorkspaceInfo() {
  if (els.currentTeamText) els.currentTeamText.value = getCurrentWorkspaceTeamLabel();
  Promise.resolve(resolveCurrentWorkspaceTeamLabelAsync())
    .then(label => {
      if (label) {
        try {
          if (!window.currentWorkspaceInfo) window.currentWorkspaceInfo = {};
          if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(label))) {
            window.currentWorkspaceInfo.name = label;
            window.currentWorkspaceInfo.team_name = label;
            window.__DROP_OFF_CURRENT_TEAM_LABEL__ = label;
          }
        } catch (_) {}
      }
      if (els.currentTeamText && label) els.currentTeamText.value = label;
    })
    .catch(error => console.warn('renderCurrentWorkspaceInfo team label refresh failed:', error));
}

function formatDateTimeLabel(value) {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || '-');
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mm = String(date.getMinutes()).padStart(2, '0');
  return `${y}/${m}/${d} ${hh}:${mm}`;
}

function getDisplayNameSeed(email) {
  const raw = String(email || '').trim();
  return raw ? (raw.split('@')[0] || raw) : 'ユーザー';
}

function setCurrentUserProfileState(profile) {
  currentUserProfile = profile || null;
  window.currentUserProfile = currentUserProfile;
  if (els.userRoleText) els.userRoleText.value = getRoleLabel(currentUserProfile?.role);
  renderCurrentWorkspaceInfo();
}

const DROP_OFF_PLAN_DEFAULTS = Object.freeze({
  plan_type: "free",
  limits: Object.freeze({
    members: 3,
    origins: 1,
    vehicles: 4,
    casts: 50
  }),
  feature_flags: Object.freeze({
    csv: false,
    line: true,
    monthly_full: true,
    plan_to_actual_add: true,
    multi_user: true,
    backup: false,
    restore: false,
    google_api_dispatch: false
  })
});

window.DROP_OFF_PLAN = window.DROP_OFF_PLAN || null;

function cloneDropOffPlanDefaults() {
  return {
    plan_type: DROP_OFF_PLAN_DEFAULTS.plan_type,
    limits: { ...DROP_OFF_PLAN_DEFAULTS.limits },
    feature_flags: { ...DROP_OFF_PLAN_DEFAULTS.feature_flags }
  };
}

function normalizeDropOffPlanRecord(raw = {}) {
  const base = cloneDropOffPlanDefaults();
  const planType = String(raw?.plan_type || base.plan_type).trim() === "paid" ? "paid" : "free";

  let rawLimits = raw?.limits;
  if (typeof rawLimits === "string") {
    try { rawLimits = JSON.parse(rawLimits); } catch (_) { rawLimits = null; }
  }
  let rawFlags = raw?.feature_flags;
  if (typeof rawFlags === "string") {
    try { rawFlags = JSON.parse(rawFlags); } catch (_) { rawFlags = null; }
  }

  const limits = {
    ...base.limits,
    ...(rawLimits && typeof rawLimits === "object" ? rawLimits : {})
  };
  const featureFlags = {
    ...base.feature_flags,
    ...(rawFlags && typeof rawFlags === "object" ? rawFlags : {})
  };

  return {
    ...base,
    ...(raw && typeof raw === "object" ? raw : {}),
    plan_type: planType,
    limits,
    feature_flags: featureFlags
  };
}

function parseBillingPeriodEndValue(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? null : date;
}

function shouldFallbackPaidPlanToFree(plan = null) {
  const normalized = normalizeDropOffPlanRecord(plan || null);
  if (String(normalized?.plan_type || 'free').trim() !== 'paid') return false;
  if (!Boolean(normalized?.billing_cancel_at_period_end)) return false;
  const periodEnd = parseBillingPeriodEndValue(normalized?.billing_current_period_end);
  if (!periodEnd) return false;
  return Date.now() >= periodEnd.getTime();
}

function applyBillingPlanFallback(plan = null) {
  const normalized = normalizeDropOffPlanRecord(plan || null);
  if (!shouldFallbackPaidPlanToFree(normalized)) return normalized;
  const base = cloneDropOffPlanDefaults();
  return {
    ...normalized,
    plan_type: 'free',
    limits: { ...base.limits },
    feature_flags: { ...base.feature_flags },
    billing_status: '',
    billing_current_period_end: null,
    billing_cancel_at_period_end: false
  };
}

function getCurrentPlanRecord() {
  return applyBillingPlanFallback(window.DROP_OFF_PLAN || null);
}

function getPlanTypeLabel(planType = "") {
  return String(planType || "").trim() === "paid" ? "Paid" : "free";
}

function getPlanLimit(key) {
  const value = getCurrentPlanRecord()?.limits?.[key];
  if (value === null || value === undefined || value === "") return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : value;
}

function isPlatformAdminPlanBypassEnabled() {
  return window.isPlatformAdminUser === true;
}

function hasPlanFeature(flag, options = {}) {
  if (options?.ignorePlatformAdmin !== true && isPlatformAdminPlanBypassEnabled()) return true;
  return Boolean(getCurrentPlanRecord()?.feature_flags?.[flag]);
}

function canUseCsvFeature(options = {}) {
  if (options?.ignorePlatformAdmin !== true && isPlatformAdminPlanBypassEnabled()) return true;
  const plan = getCurrentPlanRecord();
  const isPaid = String(plan?.plan_type || "free").trim() === "paid";
  const csvFlag = plan?.feature_flags?.csv;
  return isPaid && csvFlag !== false;
}

function getCsvFeatureBlockedMessage(actionLabel = "CSV機能") {
  const plan = getCurrentPlanRecord();
  return `${getPlanTypeLabel(plan?.plan_type)}では${actionLabel}は利用できません。Paidで利用できます。`;
}

function ensureCsvFeatureAccess(actionLabel = "CSV機能") {
  if (canUseCsvFeature()) return true;
  alert(getCsvFeatureBlockedMessage(actionLabel));
  return false;
}

function applyCsvFeatureUi() {
  const enabled = canUseCsvFeature();
  const toggleTargets = [
    els.importCsvBtn,
    els.exportCsvBtn,
    els.importVehicleCsvBtn,
    els.exportVehicleCsvBtn,
    els.exportMileageReportBtn,
    els.exportPlansCsvBtn,
    els.importPlansCsvBtn
  ];
  toggleTargets.forEach(el => {
    if (!el) return;
    el.classList.toggle("hidden", !enabled);
    el.disabled = !enabled;
  });
  if (els.csvFileInput) els.csvFileInput.disabled = !enabled;
  if (els.vehicleCsvFileInput) els.vehicleCsvFileInput.disabled = !enabled;
  if (els.plansCsvFileInput) els.plansCsvFileInput.disabled = !enabled;
}

function formatPlanLimitValue(value) {
  if (value === null || value === undefined || value === "") return "無制限";
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return String(numeric);
  return String(value || "-");
}

function formatPlanFeatureValue(enabled) {
  return enabled ? "利用可" : "利用不可";
}

const FREE_GOOGLE_API_DAILY_LIMIT = 5;
const GOOGLE_API_SHARED_FEATURE_KEY = 'google_api_shared';
const STRIPE_CHECKOUT_FUNCTION_NAME = 'quick-function';
const STRIPE_BILLING_PORTAL_FUNCTION_NAME = 'super-worker';
const BILLING_RETURN_QUERY_KEY = 'billing';
const BILLING_PERIOD_CACHE_KEY_PREFIX = 'dropoff_billing_period_end:';

function getBillingPeriodCacheKey(teamId = '') {
  const safeTeamId = String(teamId || '').trim();
  return safeTeamId ? `${BILLING_PERIOD_CACHE_KEY_PREFIX}${safeTeamId}` : BILLING_PERIOD_CACHE_KEY_PREFIX;
}

function rememberBillingPeriod(teamId, value) {
  const safeValue = String(value || '').trim();
  const key = getBillingPeriodCacheKey(teamId);
  if (!key) return;
  try {
    if (safeValue) {
      window.localStorage.setItem(key, safeValue);
    }
  } catch (_) {}
}

function getRememberedBillingPeriod(teamId) {
  const key = getBillingPeriodCacheKey(teamId);
  if (!key) return '';
  try {
    return String(window.localStorage.getItem(key) || '').trim();
  } catch (_) {
    return '';
  }
}

function resolveBillingPeriodValue(plan = null) {
  const teamId = String(plan?.id || getCurrentBillingTeamId() || '').trim();
  const directValue = String(plan?.billing_current_period_end || '').trim();
  if (directValue) {
    rememberBillingPeriod(teamId, directValue);
    return directValue;
  }
  if (Boolean(plan?.billing_cancel_at_period_end)) {
    const remembered = getRememberedBillingPeriod(teamId);
    if (remembered) return remembered;
  }
  return '';
}

function getCurrentBillingTeamId() {
  return String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync?.() || window.currentWorkspaceInfo?.id || '').trim();
}

function getBillingStatusLabel(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return '未契約';
  if (normalized === 'active') return '有効';
  if (normalized === 'inactive') return '未契約';
  if (normalized === 'trialing') return 'トライアル中';
  if (normalized === 'past_due') return '支払確認待ち';
  if (normalized === 'unpaid') return '未払い';
  if (normalized === 'canceled') return '解約済み';
  if (normalized === 'incomplete') return '決済待ち';
  if (normalized === 'incomplete_expired') return '決済期限切れ';
  return normalized;
}

function formatBillingPeriodLabel(value) {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || '-');
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

function resolveDisplayedPlanRecord(plan = null) {
  const basePlan = normalizeDropOffPlanRecord(plan || null);
  const workspaceInfo = window.currentWorkspaceInfo && typeof window.currentWorkspaceInfo === 'object'
    ? window.currentWorkspaceInfo
    : null;
  const teamId = String(basePlan?.id || window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync?.() || workspaceInfo?.id || '').trim();
  const sameWorkspaceTeam = workspaceInfo && isSameTeamId(workspaceInfo?.id, teamId);
  if (!sameWorkspaceTeam) {
    return applyBillingPlanFallback(basePlan);
  }

  const merged = {
    ...(workspaceInfo || {}),
    ...(basePlan || {}),
    id: teamId || String(workspaceInfo?.id || '').trim() || null,
    plan_type: String(basePlan?.plan_type || workspaceInfo?.plan_type || 'free').trim() === 'paid' ? 'paid' : 'free',
    limits: (basePlan?.limits && typeof basePlan.limits === 'object' ? basePlan.limits : workspaceInfo?.limits),
    feature_flags: (basePlan?.feature_flags && typeof basePlan.feature_flags === 'object' ? basePlan.feature_flags : workspaceInfo?.feature_flags),
    billing_status: String(basePlan?.billing_status || workspaceInfo?.billing_status || '').trim() || null,
    billing_current_period_end: String(basePlan?.billing_current_period_end || workspaceInfo?.billing_current_period_end || '').trim() || null,
    billing_customer_id: String(basePlan?.billing_customer_id || workspaceInfo?.billing_customer_id || '').trim() || null,
    billing_subscription_id: String(basePlan?.billing_subscription_id || workspaceInfo?.billing_subscription_id || '').trim() || null,
    billing_price_id: String(basePlan?.billing_price_id || workspaceInfo?.billing_price_id || '').trim() || null,
    billing_checkout_session_id: String(basePlan?.billing_checkout_session_id || workspaceInfo?.billing_checkout_session_id || '').trim() || null,
    billing_cancel_at_period_end: Boolean(basePlan?.billing_cancel_at_period_end || workspaceInfo?.billing_cancel_at_period_end)
  };

  return applyBillingPlanFallback(merged);
}

function setPlanCheckoutStatus(message = '', isError = false) {
  if (!els.planCheckoutStatusText) return;
  els.planCheckoutStatusText.textContent = String(message || '').trim();
  els.planCheckoutStatusText.classList.toggle('text-danger', Boolean(isError));
  els.planCheckoutStatusText.classList.toggle('text-done', !isError && Boolean(String(message || '').trim()));
}

function consumeBillingReturnState() {
  try {
    const params = new URLSearchParams(String(window.location.search || ''));
    const state = String(params.get(BILLING_RETURN_QUERY_KEY) || '').trim().toLowerCase();
    if (!state) return '';
    params.delete(BILLING_RETURN_QUERY_KEY);
    const nextQuery = params.toString();
    const nextUrl = `${window.location.pathname}${nextQuery ? `?${nextQuery}` : ''}${window.location.hash || ''}`;
    window.history.replaceState({}, '', nextUrl);
    return state;
  } catch (_) {
    return '';
  }
}

function applyBillingReturnStateMessage() {
  const state = consumeBillingReturnState();
  if (!state) return;
  if (state === 'success') {
    activateTab('settingsTab');
    setPlanCheckoutStatus('Stripeの決済完了画面から戻りました。Webhook反映後にプランを再読込してください。', false);
    return;
  }
  if (state === 'cancel') {
    activateTab('settingsTab');
    setPlanCheckoutStatus('Paid申込をキャンセルしました。決済は完了していません。', true);
    return;
  }
  if (state === 'portal') {
    activateTab('settingsTab');
    setPlanCheckoutStatus('Stripeの請求管理画面から戻りました。必要ならプランを再読込してください。', false);
  }
}

function getBillingPortalButtonLabel(plan = null) {
  const normalizedPlan = resolveDisplayedPlanRecord(plan || getCurrentPlanRecord());
  const billingStatus = String(normalizedPlan?.billing_status || '').trim().toLowerCase();
  if (billingStatus === 'past_due' || billingStatus === 'unpaid' || billingStatus === 'incomplete' || billingStatus === 'incomplete_expired') {
    return '支払管理';
  }
  if (String(normalizedPlan?.plan_type || 'free') === 'paid') {
    return '請求管理 / 解約';
  }
  return '請求管理';
}

async function openBillingPortalFromSettings() {
  if (getCurrentAppRole() !== 'owner') {
    setPlanCheckoutStatus('請求管理はオーナーのみ実行できます。', true);
    return;
  }
  const teamId = getCurrentBillingTeamId();
  if (!teamId) {
    setPlanCheckoutStatus('チーム情報を取得できませんでした。ページを再読込してください。', true);
    return;
  }

  const currentPlan = resolveDisplayedPlanRecord(getCurrentPlanRecord());
  const customerId = String(currentPlan?.billing_customer_id || '').trim();
  if (!customerId) {
    setPlanCheckoutStatus('請求管理を開くための課金情報が見つかりません。Paid申込後に再度お試しください。', true);
    return;
  }

  const sessionResult = await supabaseClient.auth.getSession();
  const accessToken = String(sessionResult?.data?.session?.access_token || '').trim();
  if (!accessToken) {
    setPlanCheckoutStatus('ログイン状態を取得できませんでした。再ログインしてください。', true);
    return;
  }

  const returnUrl = new URL(window.location.href);
  returnUrl.searchParams.set(BILLING_RETURN_QUERY_KEY, 'portal');
  returnUrl.hash = '';

  const prevText = els.openBillingPortalBtn?.textContent || '請求管理';
  if (els.openBillingPortalBtn) {
    els.openBillingPortalBtn.disabled = true;
    els.openBillingPortalBtn.textContent = 'Stripeへ移動中...';
  }
  setPlanCheckoutStatus('Stripeの請求管理画面を準備しています...', false);

  try {
    const response = await fetch(`${SUPABASE_URL}/functions/v1/${STRIPE_BILLING_PORTAL_FUNCTION_NAME}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
        apikey: SUPABASE_ANON_KEY
      },
      body: JSON.stringify({
        teamId,
        returnUrl: returnUrl.toString()
      })
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(String(payload?.error || `請求管理画面の開始に失敗しました (${response.status})`));
    }
    const portalUrl = String(payload?.portalUrl || '').trim();
    if (!portalUrl) {
      throw new Error('Stripe Billing Portal URL を取得できませんでした。');
    }
    window.location.href = portalUrl;
  } catch (error) {
    console.error('openBillingPortalFromSettings error:', error);
    setPlanCheckoutStatus(error?.message || '請求管理画面の開始に失敗しました。', true);
    if (els.openBillingPortalBtn) {
      els.openBillingPortalBtn.disabled = false;
      els.openBillingPortalBtn.textContent = prevText;
    }
  }
}

async function startPaidCheckoutFromSettings() {
  if (getCurrentAppRole() !== 'owner') {
    setPlanCheckoutStatus('Paid申込はオーナーのみ実行できます。', true);
    return;
  }
  const teamId = getCurrentBillingTeamId();
  if (!teamId) {
    setPlanCheckoutStatus('チーム情報を取得できませんでした。ページを再読込してください。', true);
    return;
  }

  const sessionResult = await supabaseClient.auth.getSession();
  const accessToken = String(sessionResult?.data?.session?.access_token || '').trim();
  if (!accessToken) {
    setPlanCheckoutStatus('ログイン状態を取得できませんでした。再ログインしてください。', true);
    return;
  }

  const currentUrl = new URL(window.location.href);
  currentUrl.searchParams.delete(BILLING_RETURN_QUERY_KEY);
  currentUrl.hash = '';
  const successUrl = new URL(currentUrl.toString());
  successUrl.searchParams.set(BILLING_RETURN_QUERY_KEY, 'success');
  const cancelUrl = new URL(currentUrl.toString());
  cancelUrl.searchParams.set(BILLING_RETURN_QUERY_KEY, 'cancel');

  const prevText = els.startPaidCheckoutBtn?.textContent || 'Paidを申し込む';
  if (els.startPaidCheckoutBtn) {
    els.startPaidCheckoutBtn.disabled = true;
    els.startPaidCheckoutBtn.textContent = 'Stripeへ移動中...';
  }
  setPlanCheckoutStatus('Stripeの決済画面を準備しています...', false);

  try {
    const response = await fetch(`${SUPABASE_URL}/functions/v1/${STRIPE_CHECKOUT_FUNCTION_NAME}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${accessToken}`,
        apikey: SUPABASE_ANON_KEY
      },
      body: JSON.stringify({
        teamId,
        successUrl: successUrl.toString(),
        cancelUrl: cancelUrl.toString()
      })
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(String(payload?.error || `Paid申込の開始に失敗しました (${response.status})`));
    }
    const checkoutUrl = String(payload?.checkoutUrl || '').trim();
    if (!checkoutUrl) {
      throw new Error('Stripe Checkout URL を取得できませんでした。');
    }
    window.location.href = checkoutUrl;
  } catch (error) {
    console.error('startPaidCheckoutFromSettings error:', error);
    setPlanCheckoutStatus(error?.message || 'Paid申込の開始に失敗しました。', true);
    if (els.startPaidCheckoutBtn) {
      els.startPaidCheckoutBtn.disabled = false;
      els.startPaidCheckoutBtn.textContent = prevText;
    }
  }
}

async function refreshPlanInfoFromSettings() {
  const prevText = els.refreshPlanInfoBtn?.textContent || 'プランを再読込';
  try {
    if (els.refreshPlanInfoBtn) {
      els.refreshPlanInfoBtn.disabled = true;
      els.refreshPlanInfoBtn.textContent = '再読込中...';
    }
    setPlanCheckoutStatus('プラン情報を再読込しています...', false);
    await loadCurrentTeamPlan();
    setPlanCheckoutStatus('プラン情報を更新しました。決済直後はWebhook反映まで数秒かかる場合があります。', false);
  } catch (error) {
    console.error('refreshPlanInfoFromSettings error:', error);
    setPlanCheckoutStatus(error?.message || 'プラン情報の再読込に失敗しました。', true);
  } finally {
    if (els.refreshPlanInfoBtn) {
      els.refreshPlanInfoBtn.disabled = false;
      els.refreshPlanInfoBtn.textContent = prevText;
    }
  }
}

function formatCastDistanceDisplay(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) return '';
  const rounded = Math.round(numeric * 10) / 10;
  return `${rounded.toLocaleString('ja-JP', { minimumFractionDigits: 0, maximumFractionDigits: 1 })} km`;
}
window.formatCastDistanceDisplay = formatCastDistanceDisplay;

function getCurrentOriginSlotLabelForUi() {
  const slotNo = normalizeOriginSlotNo(
    getCurrentWorkspaceOriginSlotFromState()
    || activeOriginSlotNo
  );
  return slotNo ? `起点${slotNo}` : '現在の起点';
}

function getSingleCurrentOriginRuntimeRecord() {
  const currentSlot = normalizeOriginSlotNo(
    getCurrentWorkspaceOriginSlotFromState()
    || activeOriginSlotNo
  );
  if (!currentSlot) return null;

  const matched = normalizeOriginRecord(getOriginRowBySlot(currentSlot) || {});
  if (
    normalizeOriginSlotNo(matched?.slot_no) !== currentSlot ||
    !Number.isFinite(Number(matched?.lat)) ||
    !Number.isFinite(Number(matched?.lng))
  ) {
    return null;
  }

  return {
    slot_no: currentSlot,
    name: getOriginDisplayLabel(matched?.name || matched?.label || ORIGIN_LABEL),
    address: String(matched?.address || '').trim(),
    lat: Number(matched?.lat),
    lng: Number(matched?.lng)
  };
}

function getStrictCurrentOriginRuntimeForLiveDisplay() {
  return getSingleCurrentOriginRuntimeRecord();
}

function getCastDistanceSourceText() {
  const runtimeOrigin = getStrictCurrentOriginRuntimeForLiveDisplay()
    || (typeof getCurrentOriginRuntime === 'function' ? getCurrentOriginRuntime() : null);
  const originLabel = String(runtimeOrigin?.name || getCurrentOriginSlotLabelForUi() || '起点').trim() || '起点';
  const slotLabel = runtimeOrigin?.slot_no ? `起点${runtimeOrigin.slot_no}` : getCurrentOriginSlotLabelForUi();
  const castAddress = String(els.castAddress?.value || '').trim() || '送り先住所';
  return `${slotLabel} / ${originLabel} → ${castAddress}`;
}

function updateCastDistanceHint(options = {}) {
  if (!els.castDistanceHint) return;
  const sourceText = getCastDistanceSourceText();
  if (options.hidden) {
    els.castDistanceHint.textContent = '';
    els.castDistanceHint.style.display = 'none';
    return;
  }
  els.castDistanceHint.style.display = '';
  if (options.invalidDistance) {
    els.castDistanceHint.textContent = `計算元: ${sourceText} | 距離が異常です。現在の起点座標を確認してください`;
    return;
  }
  if (options.distanceKm != null) {
    els.castDistanceHint.textContent = `計算元: ${sourceText}`;
    return;
  }
  els.castDistanceHint.textContent = `計算元: ${sourceText} | 住所と座標がそろうと距離を表示します`;
}

function syncCastBlankMetricsUi(options = {}) {
  const isEditing = options.forceEditing === true || !!editingCastId;
  const address = String(els.castAddress?.value || '').trim();
  const latLngText = String(els.castLatLngText?.value || '').trim();
  const lat = String(els.castLat?.value || '').trim();
  const lng = String(els.castLng?.value || '').trim();
  const hasCoords = !!(latLngText || (lat && lng));
  if (!isEditing && !address && !hasCoords) {
    if (els.castDistanceKm) els.castDistanceKm.value = '';
    if (els.castTravelMinutes) els.castTravelMinutes.value = '';
    updateCastDistanceHint({ hidden: true });
    return true;
  }
  return false;
}

async function refreshCastGoogleApiQuotaUi() {
  if (!els.fetchCastTravelMinutesBtn && !els.castApiQuotaText) return;
  const plan = getCurrentPlanRecord();
  if (isPlatformAdminPlanBypassEnabled()) {
    if (els.fetchCastTravelMinutesBtn) els.fetchCastTravelMinutesBtn.textContent = 'APIで座標取得';
    if (els.castApiQuotaText) els.castApiQuotaText.textContent = '運営者: API座標取得は制限なし';
    return;
  }
  if (String(plan?.plan_type || 'free') === 'paid') {
    if (els.fetchCastTravelMinutesBtn) els.fetchCastTravelMinutesBtn.textContent = 'APIで座標取得';
    if (els.castApiQuotaText) els.castApiQuotaText.textContent = 'Paid: API座標取得は制限なし';
    return;
  }
  let remaining = FREE_GOOGLE_API_DAILY_LIMIT;
  try {
    const snapshot = await ensureGoogleApiCoordinateLookupAccess({ actionLabel: 'API座標取得', consume: false });
    if (Number.isFinite(Number(snapshot?.remaining_count))) {
      remaining = Math.max(0, Number(snapshot.remaining_count));
    }
  } catch (_) {}
  if (els.fetchCastTravelMinutesBtn) {
    els.fetchCastTravelMinutesBtn.textContent = `APIで座標取得（残り${remaining}回）`;
  }
  if (els.castApiQuotaText) {
    els.castApiQuotaText.textContent = `free: API座標取得は1日${FREE_GOOGLE_API_DAILY_LIMIT}回まで。残り${remaining}回`;
  }
}


function getVisibleOriginSlotLimit() {
  if (isPlatformAdminPlanBypassEnabled()) return ORIGIN_SLOT_LIMIT;
  const numericLimit = Number(getPlanLimit('origins'));
  if (Number.isFinite(numericLimit) && numericLimit >= 1) {
    return Math.max(1, Math.min(ORIGIN_SLOT_LIMIT, Math.floor(numericLimit)));
  }
  return String(getCurrentPlanRecord()?.plan_type || 'free') === 'paid' ? ORIGIN_SLOT_LIMIT : 1;
}

function applyOriginSlotUi() {
  const visibleLimit = getVisibleOriginSlotLimit();
  if (els.originSlotSelect) {
    Array.from(els.originSlotSelect.options || []).forEach(option => {
      const slotNo = normalizeOriginSlotNo(option?.value);
      const visible = Boolean(slotNo) && slotNo <= visibleLimit;
      option.hidden = !visible;
      option.disabled = !visible;
    });
    const selected = normalizeOriginSlotNo(els.originSlotSelect.value || '');
    if (!selected || selected > visibleLimit) {
      els.originSlotSelect.value = String(Math.max(1, Math.min(visibleLimit, normalizeOriginSlotNo(activeOriginSlotNo) || 1)));
    }
  }
}

async function getGoogleApiUsageSnapshot() {
  const plan = getCurrentPlanRecord();
  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
  if (!teamId) {
    return { allowed: true, used_count: 0, limit_count: FREE_GOOGLE_API_DAILY_LIMIT, remaining_count: FREE_GOOGLE_API_DAILY_LIMIT, plan };
  }
  if (typeof window.getSharedGoogleApiUsageDaily !== 'function') {
    return { allowed: true, used_count: 0, limit_count: FREE_GOOGLE_API_DAILY_LIMIT, remaining_count: FREE_GOOGLE_API_DAILY_LIMIT, plan };
  }
  try {
    const { data, error } = await window.getSharedGoogleApiUsageDaily(teamId, GOOGLE_API_SHARED_FEATURE_KEY);
    if (error) console.warn('google api usage snapshot warning:', error);
    const used = Math.max(0, Number(data?.used_count || 0));
    const remaining = Math.max(0, FREE_GOOGLE_API_DAILY_LIMIT - used);
    return { allowed: remaining > 0, used_count: used, limit_count: FREE_GOOGLE_API_DAILY_LIMIT, remaining_count: remaining, plan };
  } catch (error) {
    console.warn('google api usage snapshot failed:', error);
    return { allowed: true, used_count: 0, limit_count: FREE_GOOGLE_API_DAILY_LIMIT, remaining_count: FREE_GOOGLE_API_DAILY_LIMIT, plan };
  }
}

async function ensureGoogleApiCoordinateLookupAccess(options = {}) {
  const actionLabel = String(options?.actionLabel || 'API座標取得').trim() || 'API座標取得';
  const consume = options?.consume === true;
  const plan = getCurrentPlanRecord();
  if (isPlatformAdminPlanBypassEnabled()) {
    return { allowed: true, reason: '', used_count: 0, remaining_count: null, limit_count: null, plan, bypass: 'platform_admin' };
  }
  if (String(plan?.plan_type || 'free') === 'paid') {
    return { allowed: true, reason: '', used_count: 0, remaining_count: null, limit_count: null, plan, bypass: 'paid' };
  }

  const snapshot = await getGoogleApiUsageSnapshot();
  if (!consume) {
    return {
      ...snapshot,
      allowed: snapshot.remaining_count > 0,
      reason: snapshot.remaining_count > 0
        ? `freeでは${actionLabel}を1日${FREE_GOOGLE_API_DAILY_LIMIT}回まで利用できます。残り${snapshot.remaining_count}回です。`
        : `freeでは${actionLabel}は1日${FREE_GOOGLE_API_DAILY_LIMIT}回までです。本日の上限に達しました。`
    };
  }

  if (snapshot.remaining_count <= 0) {
    return {
      ...snapshot,
      allowed: false,
      reason: `freeでは${actionLabel}は1日${FREE_GOOGLE_API_DAILY_LIMIT}回までです。本日の上限に達しました。`
    };
  }

  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
  if (!teamId || typeof window.incrementSharedGoogleApiUsageDaily !== 'function') {
    return {
      ...snapshot,
      allowed: true,
      remaining_count: Math.max(0, Number(snapshot.remaining_count || FREE_GOOGLE_API_DAILY_LIMIT) - 1),
      used_count: Number(snapshot.used_count || 0) + 1,
      reason: ''
    };
  }

  try {
    const { data, error } = await window.incrementSharedGoogleApiUsageDaily(teamId, GOOGLE_API_SHARED_FEATURE_KEY, 1);
    if (error) {
      console.warn('google api usage increment warning:', error);
      return {
        ...snapshot,
        allowed: true,
        remaining_count: Math.max(0, Number(snapshot.remaining_count || FREE_GOOGLE_API_DAILY_LIMIT) - 1),
        used_count: Number(snapshot.used_count || 0) + 1,
        reason: ''
      };
    }
    const used = Math.max(0, Number(data?.used_count || 0));
    const remaining = Math.max(0, FREE_GOOGLE_API_DAILY_LIMIT - used);
    return { allowed: true, reason: '', used_count: used, remaining_count: remaining, limit_count: FREE_GOOGLE_API_DAILY_LIMIT, plan };
  } catch (error) {
    console.warn('google api usage increment failed:', error);
    return {
      ...snapshot,
      allowed: true,
      remaining_count: Math.max(0, Number(snapshot.remaining_count || FREE_GOOGLE_API_DAILY_LIMIT) - 1),
      used_count: Number(snapshot.used_count || 0) + 1,
      reason: ''
    };
  }
}

function getGoogleApiPlanStatusText() {
  if (isPlatformAdminPlanBypassEnabled()) return '利用可（制限なし）';
  const plan = getCurrentPlanRecord();
  if (String(plan?.plan_type || 'free') === 'paid') return '利用可';
  return `API座標取得は1日${FREE_GOOGLE_API_DAILY_LIMIT}回`;
}

function getPlanUsageCounts() {
  const origins = Array.isArray(allOriginsCache)
    ? new Set(
        allOriginsCache
          .map(row => normalizeOriginSlotNo(row?.slot_no))
          .filter(Boolean)
      ).size
    : 0;
  const vehicles = Array.isArray(allVehiclesCache) ? allVehiclesCache.length : 0;
  const casts = Array.isArray(allCastsCache) ? allCastsCache.length : 0;
  return { origins, vehicles, casts };
}


function normalizePlanMemberUsage(payload = {}) {
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

function getCurrentPlanMemberUsage() {
  return normalizePlanMemberUsage(window.DROP_OFF_PLAN_USAGE || {});
}

async function loadCurrentTeamPlanUsage() {
  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || "").trim();
  if (!teamId) {
    window.DROP_OFF_PLAN_USAGE = normalizePlanMemberUsage();
    return window.DROP_OFF_PLAN_USAGE;
  }

  if (typeof window.getTeamPlanMemberUsage !== "function") {
    window.DROP_OFF_PLAN_USAGE = normalizePlanMemberUsage();
    return window.DROP_OFF_PLAN_USAGE;
  }

  try {
    const { data, error } = await window.getTeamPlanMemberUsage(teamId);
    if (error) console.warn("team plan member usage warning:", error);
    window.DROP_OFF_PLAN_USAGE = normalizePlanMemberUsage(data || {});
  } catch (error) {
    console.warn("team plan member usage failed:", error);
    window.DROP_OFF_PLAN_USAGE = normalizePlanMemberUsage();
  }
  return window.DROP_OFF_PLAN_USAGE;
}

function formatPlanUsageValue(limitKey, currentCount) {
  const limitValue = getPlanLimit(limitKey);
  const limitText = formatPlanLimitValue(limitValue);
  return `${Number(currentCount || 0)} / ${limitText}`;
}

function buildPlanLimitCheckResult({ limitKey, label, currentCount = 0, isEditingExisting = false, unit = '件', exceedMessage = '' } = {}) {
  const plan = getCurrentPlanRecord();
  if (isPlatformAdminPlanBypassEnabled()) {
    return { allowed: true, reason: '', count: Number(currentCount || 0), limit: null, plan };
  }
  const numericLimit = Number(getPlanLimit(limitKey));
  if (!Number.isFinite(numericLimit) || numericLimit < 1) {
    return { allowed: true, reason: '', count: Number(currentCount || 0), limit: null, plan };
  }
  if (isEditingExisting) {
    return { allowed: true, reason: '', count: Number(currentCount || 0), limit: numericLimit, plan };
  }
  if (String(plan?.plan_type || 'free') === 'paid') {
    return { allowed: true, reason: '', count: Number(currentCount || 0), limit: numericLimit, plan };
  }
  if (Number(currentCount || 0) < numericLimit) {
    return { allowed: true, reason: '', count: Number(currentCount || 0), limit: numericLimit, plan };
  }
  return {
    allowed: false,
    reason: exceedMessage || `${getPlanTypeLabel(plan?.plan_type)}では${label}は${numericLimit}${unit}までです。不要なデータを削除してから追加してください。`,
    count: Number(currentCount || 0),
    limit: numericLimit,
    plan
  };
}

function formatPlanMemberUsageValue() {
  const usage = getCurrentPlanMemberUsage();
  const limitText = formatPlanLimitValue(getPlanLimit("members"));
  const base = `${Number(usage.seats_used || 0)} / ${limitText}`;
  return Number(usage.pending_invitations || 0) > 0 ? `${base}（招待中${Number(usage.pending_invitations || 0)}）` : base;
}

function renderPlanInfo() {
  const plan = resolveDisplayedPlanRecord(getCurrentPlanRecord());
  const counts = getPlanUsageCounts();
  const role = getCurrentAppRole();
  const isOwner = role === 'owner';
  const billingStatus = String(plan?.billing_status || '').trim();
  const billingPeriod = resolveBillingPeriodValue(plan);
  const cancelAtPeriodEnd = Boolean(plan?.billing_cancel_at_period_end);
  const isPaidPlan = String(plan?.plan_type || 'free') === 'paid';
  renderCurrentWorkspaceInfo();
  if (els.planTypeText) els.planTypeText.value = getPlanTypeLabel(plan?.plan_type);
  if (els.planMembersText) els.planMembersText.value = formatPlanMemberUsageValue();
  if (els.planOriginsText) els.planOriginsText.value = formatPlanUsageValue("origins", counts.origins);
  if (els.planVehiclesText) els.planVehiclesText.value = formatPlanUsageValue("vehicles", counts.vehicles);
  if (els.planCastsText) els.planCastsText.value = formatPlanUsageValue("casts", counts.casts);
  if (els.planGoogleApiText) els.planGoogleApiText.value = getGoogleApiPlanStatusText();
  if (els.planBackupText) {
    const enabled = hasPlanFeature("backup") || hasPlanFeature("restore");
    els.planBackupText.value = formatPlanFeatureValue(enabled);
  }
  if (els.planCsvText) els.planCsvText.value = formatPlanFeatureValue(canUseCsvFeature());
  if (els.planBillingStatusText) {
    const statusLabel = getBillingStatusLabel(billingStatus || (isPaidPlan ? 'active' : 'inactive'));
    els.planBillingStatusText.value = isPaidPlan
      ? `${statusLabel}${cancelAtPeriodEnd ? '（期間終了でfreeへ戻る予定）' : ''} / Paid`
      : (billingStatus ? `${statusLabel} / free` : '未契約 / free');
  }
  if (els.planBillingPeriodText) {
    els.planBillingPeriodText.value = billingPeriod ? formatBillingPeriodLabel(billingPeriod) : '-';
  }
  if (els.startPaidCheckoutBtn) {
    els.startPaidCheckoutBtn.classList.toggle('hidden', !isOwner);
    els.startPaidCheckoutBtn.disabled = !isOwner || isPaidPlan;
    els.startPaidCheckoutBtn.textContent = isPaidPlan ? 'Paid適用中' : 'Paidを申し込む';
  }
  if (els.openBillingPortalBtn) {
    const canOpenPortal = isOwner && isPaidPlan && Boolean(String(plan?.billing_customer_id || '').trim());
    els.openBillingPortalBtn.classList.toggle('hidden', !canOpenPortal);
    els.openBillingPortalBtn.disabled = !canOpenPortal;
    els.openBillingPortalBtn.textContent = getBillingPortalButtonLabel(plan);
  }
  if (els.refreshPlanInfoBtn) {
    els.refreshPlanInfoBtn.classList.toggle('hidden', !isOwner);
    els.refreshPlanInfoBtn.disabled = !isOwner;
  }
  if (els.planCheckoutStatusText && !isOwner) {
    setPlanCheckoutStatus('Paid申込とプラン更新はオーナーのみ実行できます。', false);
  } else if (els.planCheckoutStatusText && !String(els.planCheckoutStatusText.textContent || '').trim()) {
    if (isPaidPlan && cancelAtPeriodEnd && billingPeriod) {
      setPlanCheckoutStatus(`Stripe側で解約予定です。${formatBillingPeriodLabel(billingPeriod)} の期間終了後にfreeへ戻ります。`, false);
    } else if (isPaidPlan && cancelAtPeriodEnd) {
      setPlanCheckoutStatus('Stripe側で解約予定です。期間終了日の反映を待っています。', false);
    } else if (billingStatus === 'past_due' || billingStatus === 'unpaid' || billingStatus === 'incomplete' || billingStatus === 'incomplete_expired') {
      setPlanCheckoutStatus('支払状態の確認が必要です。請求管理から支払方法や契約状態を確認してください。', true);
    } else if (isPaidPlan) {
      setPlanCheckoutStatus('Paid適用中です。請求管理から解約や支払方法の変更ができます。', false);
    } else {
      setPlanCheckoutStatus('課金を開始すると Stripe の決済画面へ移動します。反映後はプランを再読込してください。', false);
    }
  }
  applyCsvFeatureUi();
  applyGoogleApiFeatureUi();
  applyOriginSlotUi();
  refreshCastGoogleApiQuotaUi();
  updateCastDistanceHint({ hidden: true });
}

async function loadCurrentTeamPlan() {
  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || "").trim();
  if (!teamId) {
    window.DROP_OFF_PLAN = normalizeDropOffPlanRecord(null);
    window.DROP_OFF_PLAN_USAGE = normalizePlanMemberUsage();
    renderPlanInfo();
    return window.DROP_OFF_PLAN;
  }

  if (typeof window.getTeamPlan !== "function") {
    window.DROP_OFF_PLAN = normalizeDropOffPlanRecord({ id: teamId });
    await loadCurrentTeamPlanUsage();
    renderPlanInfo();
    return window.DROP_OFF_PLAN;
  }

  try {
    const { data, error } = await window.getTeamPlan(teamId);
    if (error) console.warn("team plan load warning:", error);
    const workspacePlanFallback = isSameTeamId(window.currentWorkspaceInfo?.id, teamId)
      ? (window.currentWorkspaceInfo || {})
      : {};
    window.DROP_OFF_PLAN = resolveDisplayedPlanRecord({
      ...(workspacePlanFallback || {}),
      ...(data || {}),
      id: teamId
    });
  } catch (error) {
    console.warn("team plan load failed:", error);
    const workspacePlanFallback = isSameTeamId(window.currentWorkspaceInfo?.id, teamId)
      ? (window.currentWorkspaceInfo || {})
      : {};
    window.DROP_OFF_PLAN = resolveDisplayedPlanRecord({ ...(workspacePlanFallback || {}), id: teamId });
  }

  await loadCurrentTeamPlanUsage();
  renderPlanInfo();
  return window.DROP_OFF_PLAN;
}

function canAddOrigin(options = {}) {
  const counts = getPlanUsageCounts();
  return buildPlanLimitCheckResult({
    limitKey: 'origins',
    label: '起点',
    currentCount: counts.origins,
    isEditingExisting: Boolean(options?.isEditingExisting)
  });
}

function canAddVehicle(options = {}) {
  const counts = getPlanUsageCounts();
  return buildPlanLimitCheckResult({
    limitKey: 'vehicles',
    label: '車両',
    currentCount: counts.vehicles,
    isEditingExisting: Boolean(options?.isEditingExisting)
  });
}

function canAddCast(options = {}) {
  const counts = getPlanUsageCounts();
  return buildPlanLimitCheckResult({
    limitKey: 'casts',
    label: '送り先',
    currentCount: counts.casts,
    isEditingExisting: Boolean(options?.isEditingExisting)
  });
}

function canAddMemberInvite() {
  const usage = getCurrentPlanMemberUsage();
  return buildPlanLimitCheckResult({
    limitKey: 'members',
    label: '利用人数',
    currentCount: usage.seats_used,
    unit: '名',
    exceedMessage: `freeでは利用人数は${Number(getPlanLimit('members') || 0)}名までです。不要なメンバーや招待を整理してから追加してください。`
  });
}

function canUseGoogleApi(options = {}) {
  const plan = getCurrentPlanRecord();
  const purpose = String(options?.purpose || '').trim();
  if (isPlatformAdminPlanBypassEnabled()) {
    return { allowed: true, reason: '', plan };
  }
  if (String(plan?.plan_type || 'free') === 'paid') {
    return { allowed: true, reason: '', plan };
  }
  if (purpose === 'coordinate_lookup') {
    return {
      allowed: true,
      reason: `freeではAPI座標取得を1日${FREE_GOOGLE_API_DAILY_LIMIT}回まで利用できます。`,
      plan,
      limited: true,
      daily_limit: FREE_GOOGLE_API_DAILY_LIMIT
    };
  }
  return {
    allowed: false,
    reason: `${getPlanTypeLabel(plan?.plan_type)}ではGoogleMap APIを使う機能は利用できません。Paidで利用できます。`,
    plan
  };
}

function applyGoogleApiFeatureUi() {
  const access = canUseGoogleApi({ purpose: 'coordinate_lookup' });
  const role = getCurrentAppRole();
  const canEdit = window.isPlatformAdminUser === true || role === 'owner' || role === 'admin';
  const enabled = canEdit && access.allowed;
  const title = enabled
    ? (isPlatformAdminPlanBypassEnabled()
        ? '住所から座標を取得します'
        : (String(getCurrentPlanRecord()?.plan_type || 'free') === 'paid'
            ? '住所から座標を取得します'
            : `freeではAPI座標取得を1日${FREE_GOOGLE_API_DAILY_LIMIT}回まで利用できます。`))
    : (access.reason || 'この操作は現在のプランでは利用できません。');
  const targets = [els.fetchOriginLatLngBtn, els.fetchCastTravelMinutesBtn];
  targets.forEach(btn => {
    if (!btn) return;
    btn.disabled = !enabled;
    btn.title = title;
  });
}

function ensureSettingsSectionExpanded(sectionEl, expanded = true) {
  if (!sectionEl) return;
  if ('open' in sectionEl) {
    sectionEl.open = Boolean(expanded);
  }
}

function getAllowedTabsForRole(role) {
  const normalized = normalizeAppRole(role);
  const base = (normalized === 'owner' || normalized === 'admin')
    ? ['homeTab', 'castsTab', 'castSearchTab', 'vehiclesTab', 'dailyTab', 'plansTab', 'actualTab', 'settingsTab']
    : ['homeTab', 'castsTab', 'castSearchTab', 'vehiclesTab', 'dailyTab', 'plansTab', 'actualTab', 'settingsTab'];
  if (window.isPlatformAdminUser) base.push('platformAdminTab');
  return new Set(base);
}

function applyRoleUi() {
  const role = getCurrentAppRole();
  const isOwner = role === 'owner';
  const isManager = role === 'owner' || role === 'admin';
  const isReadonlyUser = role === 'user';
  const allowedTabs = getAllowedTabsForRole(role);

  document.querySelectorAll('.main-tab').forEach(btn => {
    const tabId = btn.dataset.tab;
    if (!tabId) return;
    btn.classList.toggle('hidden', !allowedTabs.has(tabId));
  });
  if (els.platformAdminTabBtn) els.platformAdminTabBtn.classList.toggle('hidden', !window.isPlatformAdminUser);

  if (els.userManagementSection) els.userManagementSection.classList.toggle('hidden', !isOwner);
  if (els.invitationManagementSection) els.invitationManagementSection.classList.toggle('hidden', !isManager);
  if (els.originManagementSection) els.originManagementSection.classList.toggle('hidden', !isManager);
  if (els.dataManagementSection) els.dataManagementSection.classList.toggle('hidden', !isOwner);
  if (els.ownerDangerZone) els.ownerDangerZone.classList.toggle('hidden', !isOwner);
  if (els.castEditorWrap) els.castEditorWrap.classList.toggle('hidden', isReadonlyUser);
  if (els.vehicleEditorWrap) els.vehicleEditorWrap.classList.toggle('hidden', isReadonlyUser);

  const castControls = [els.saveCastBtn, els.cancelEditBtn, els.importCsvBtn, els.exportCsvBtn, els.guessAreaBtn, els.fetchCastTravelMinutesBtn];
  const vehicleControls = [els.saveVehicleBtn, els.cancelVehicleEditBtn, els.importVehicleCsvBtn, els.exportVehicleCsvBtn, els.exportMileageCsvBtn, els.runMileageReportBtn];
  const originControls = [els.fetchOriginLatLngBtn, els.openOriginGoogleMapBtn, els.saveOriginBtn, els.useOriginDraftBtn, els.cancelOriginEditBtn];
  const invitationControls = [els.inviteEmailInput, els.inviteDisplayNameInput, els.inviteRoleSelect, els.inviteTeamSelect, els.sendInvitationBtn, els.refreshInvitationsBtn];
  castControls.forEach(btn => btn && (btn.disabled = !isManager));
  vehicleControls.forEach(btn => btn && (btn.disabled = !isManager));
  originControls.forEach(btn => btn && (btn.disabled = !isManager));
  invitationControls.forEach(btn => btn && (btn.disabled = !isManager));
  applyCsvFeatureUi();
  applyGoogleApiFeatureUi();
  renderInvitationRoleOptions();
  if (els.exportAllBtn) els.exportAllBtn.disabled = !isOwner;
  if (els.importAllBtn) els.importAllBtn.disabled = !isOwner;
  if (els.resetCastsBtn) els.resetCastsBtn.disabled = !isOwner;
  if (els.resetVehiclesBtn) els.resetVehiclesBtn.disabled = !isOwner;
  if (els.dangerResetBtn) els.dangerResetBtn.disabled = !isOwner;

  const activePanel = document.querySelector('.page-panel.active');
  if (activePanel && !allowedTabs.has(activePanel.id)) {
    activateTab('homeTab');
  }

  renderCurrentWorkspaceInfo();
  refreshAdminHeaderVisibility();
  renderAdminForceModeBanner();
}

function getProfileRowById(profileId) {
  return allUserProfilesCache.find(row => String(row.id) === String(profileId) || String(row.user_id) === String(profileId)) || null;
}

function renderProfilesTable() {
  if (!els.profilesTableWrap) return;
  const rows = Array.isArray(allUserProfilesCache) ? allUserProfilesCache : [];
  if (!rows.length) {
    els.profilesTableWrap.innerHTML = '<p class="soft-text">ログイン済みユーザーがまだいません。</p>';
    return;
  }

  const body = rows.map(row => {
    const isOwnerRow = normalizeAppRole(row.role) === 'owner';
    const isSelf = String(row.id || row.user_id || '') === String(currentUser?.id || currentUserProfile?.id || '');
    const isProfileMissing = row.profile_missing === true || (!row.email && String(row.display_name || '') === '未連携ユーザー');
    const nameValue = escapeHtml(row.display_name || '');
    const emailValue = escapeHtml(isProfileMissing ? 'プロフィール未連携' : (row.email || '-'));
    const roleOptions = isOwnerRow
      ? `<option value="owner" selected>owner</option>`
      : ['admin', 'user'].map(roleValue => `<option value="${roleValue}" ${normalizeAppRole(row.role) === roleValue ? 'selected' : ''}>${roleValue}</option>`).join('');
    const activeOptions = [
      `<option value="true" ${row.is_active !== false ? 'selected' : ''}>有効</option>`,
      `<option value="false" ${row.is_active === false ? 'selected' : ''}>無効</option>`
    ].join('');
    return `
      <tr data-profile-id="${escapeHtml(row.id || row.user_id || '')}">
        <td><input class="profile-display-input" type="text" value="${nameValue}" ${row.is_active === false ? 'style="opacity:.75;"' : ''}></td>
        <td>${emailValue}</td>
        <td><span class="chip">${escapeHtml(getRoleLabel(row.role))}</span></td>
        <td>${escapeHtml(row.team_name || '-')}</td>
        <td>
          <select class="profile-role-select" ${isOwnerRow ? 'disabled' : ''}>
            ${roleOptions}
          </select>
        </td>
        <td>
          <select class="profile-active-select" ${(isOwnerRow || isSelf) ? 'disabled' : ''}>
            ${activeOptions}
          </select>
        </td>
        <td>${escapeHtml(formatDateTimeLabel(row.created_at))}</td>
        <td class="action-row">
          <button class="btn ghost profile-save-btn">保存</button>
          <button class="btn danger profile-delete-btn" ${(isOwnerRow || isSelf) ? 'disabled' : ''}>削除</button>
        </td>
      </tr>`;
  }).join('');

  els.profilesTableWrap.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th>表示名</th>
          <th>メール</th>
          <th>現在権限</th>
          <th>所属チーム</th>
          <th>変更後権限</th>
          <th>状態</th>
          <th>作成日</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>`;
}

async function loadProfilesForSettings() {
  if (!isOwnerUser() || typeof loadUserProfilesCloud !== 'function') {
    allUserProfilesCache = [];
    renderProfilesTable();
    return;
  }
  try {
    const settingsTeamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || '').trim() || null;
    allUserProfilesCache = await loadUserProfilesCloud(settingsTeamId);
    renderProfilesTable();
  } catch (error) {
    console.error(error);
    els.profilesTableWrap.innerHTML = `<p class="soft-text">ユーザー一覧の読込に失敗しました: ${escapeHtml(error?.message || String(error))}</p>`;
  }
}

async function saveProfileRowFromTable(profileId) {
  if (!isOwnerUser() || typeof saveUserProfileCloud !== 'function') {
    alert('この操作はオーナーだけが行えます。');
    return;
  }
  const row = getProfileRowById(profileId);
  const tr = Array.from(els.profilesTableWrap?.querySelectorAll('tr[data-profile-id]') || []).find(node => String(node.dataset.profileId || '') === String(profileId));
  if (!row || !tr) return;

  const displayName = tr.querySelector('.profile-display-input')?.value.trim() || getDisplayNameSeed(row.email || '');
  const role = normalizeAppRole(tr.querySelector('.profile-role-select')?.value || row.role);
  const isActive = String(tr.querySelector('.profile-active-select')?.value || String(row.is_active !== false)) !== 'false';

  try {
    const saved = await saveUserProfileCloud({ ...row, display_name: displayName, role: normalizeAppRole(row.role) === 'owner' ? 'owner' : role, is_active: normalizeAppRole(row.role) === 'owner' ? true : isActive });
    const idx = allUserProfilesCache.findIndex(item => String(item.id || item.user_id) === String(profileId));
    if (idx >= 0) allUserProfilesCache[idx] = { ...allUserProfilesCache[idx], ...saved };
    if (String(saved.id || saved.user_id) === String(currentUser?.id || currentUserProfile?.id || '')) {
      setCurrentUserProfileState({ ...currentUserProfile, ...saved });
      applyRoleUi();
    syncHomeSetupCardVisibility();
    }
    renderProfilesTable();
    alert('ユーザー情報を保存しました');
  } catch (error) {
    console.error(error);
    alert('ユーザー情報の保存に失敗しました: ' + (error?.message || error));
  }
}

async function deleteProfileRowFromTable(profileId) {
  if (!isOwnerUser() || typeof deleteUserProfileCloud !== 'function') {
    alert('この操作はオーナーだけが行えます。');
    return;
  }
  const row = getProfileRowById(profileId);
  if (!row) return;
  if (normalizeAppRole(row.role) === 'owner') {
    alert('オーナーは削除できません。');
    return;
  }
  if (String(row.id || row.user_id || '') === String(currentUser?.id || currentUserProfile?.id || '')) {
    alert('自分自身は削除できません。');
    return;
  }
  if (!window.confirm(`このユーザー管理行を削除しますか？

${row.email || '-'}`)) return;

  try {
    await deleteUserProfileCloud(row);
    allUserProfilesCache = allUserProfilesCache.filter(item => String(item.id || item.user_id) !== String(profileId));
    renderProfilesTable();
    alert('ユーザー管理行を削除しました');
  } catch (error) {
    console.error(error);
    alert('削除に失敗しました: ' + (error?.message || error));
  }
}

function bindProfileTableEvents() {
  if (!els.profilesTableWrap) return;
  els.profilesTableWrap.addEventListener('click', async event => {
    const saveBtn = event.target.closest('.profile-save-btn');
    if (saveBtn) {
      const profileId = saveBtn.closest('tr')?.dataset.profileId || '';
      if (profileId) await saveProfileRowFromTable(profileId);
      return;
    }
    const deleteBtn = event.target.closest('.profile-delete-btn');
    if (deleteBtn) {
      const profileId = deleteBtn.closest('tr')?.dataset.profileId || '';
      if (profileId) await deleteProfileRowFromTable(profileId);
    }
  });
}

function getInvitationRoleLabel(role) {
  return normalizeInvitationRoleForUi(role) === 'admin' ? '管理者' : '利用者';
}

function normalizeInvitationRoleForUi(role) {
  if (typeof normalizeInvitationRole === 'function') return normalizeInvitationRole(role);
  return String(role || '').trim().toLowerCase() === 'admin' ? 'admin' : 'user';
}

function normalizeInvitationStatusForUi(status, expiresAt = null) {
  if (typeof normalizeInvitationStatus === 'function') return normalizeInvitationStatus(status, expiresAt);
  const value = String(status || '').trim().toLowerCase();
  if (value === 'accepted' || value === 'revoked' || value === 'expired') return value;
  const expires = expiresAt ? new Date(expiresAt) : null;
  if (expires && !Number.isNaN(expires.getTime()) && expires.getTime() < Date.now()) return 'expired';
  return 'pending';
}

function getInvitationStatusLabel(status) {
  const normalized = normalizeInvitationStatusForUi(status);
  if (normalized === 'accepted') return '参加済み';
  if (normalized === 'revoked') return '取消済み';
  if (normalized === 'expired') return '期限切れ';
  return '招待中';
}

function getInvitationTeamName(teamId) {
  const found = (Array.isArray(invitationTeamOptionsCache) ? invitationTeamOptionsCache : []).find(row => String(row.id || '') === String(teamId || ''));
  return found?.name || '-';
}

function getVisibleInvitationRows() {
  const rows = Array.isArray(allInvitationRowsCache) ? [...allInvitationRowsCache] : [];
  if (isOwnerUser()) return rows;
  if (!isManagerUser()) return [];
  const uid = String(currentUser?.id || currentUserProfile?.id || '');
  return rows.filter(row => String(row.invited_by_user_id || '') === uid);
}

function canCurrentUserInviteRole(role) {
  const normalized = normalizeInvitationRoleForUi(role);
  if (isOwnerUser()) return normalized === 'admin' || normalized === 'user';
  if (isManagerUser()) return normalized === 'user';
  return false;
}

function canCurrentUserManageInvitation(row) {
  if (!row) return false;
  const role = normalizeInvitationRoleForUi(row.invited_role);
  const status = normalizeInvitationStatusForUi(row.status, row.expires_at);
  if (status !== 'pending') return false;
  if (isOwnerUser()) return role === 'admin' || role === 'user';
  const uid = String(currentUser?.id || currentUserProfile?.id || '');
  return isManagerUser() && role === 'user' && String(row.invited_by_user_id || '') === uid;
}

function renderInvitationRoleOptions() {
  if (!els.inviteRoleSelect) return;
  const options = isOwnerUser() ? ['admin', 'user'] : ['user'];
  const currentValue = normalizeInvitationRoleForUi(els.inviteRoleSelect.value || options[0]);
  els.inviteRoleSelect.innerHTML = options.map(value => `<option value="${value}">${value}</option>`).join('');
  els.inviteRoleSelect.value = options.includes(currentValue) ? currentValue : options[0];
}

function renderInvitationTeamOptions() {
  if (!els.inviteTeamSelect) return;
  const options = Array.isArray(invitationTeamOptionsCache) ? invitationTeamOptionsCache : [];
  if (!options.length) {
    els.inviteTeamSelect.innerHTML = '<option value="">チームがまだありません</option>';
    return;
  }
  const current = String(els.inviteTeamSelect.value || '');
  els.inviteTeamSelect.innerHTML = ['<option value="">選択してください</option>']
    .concat(options.map(team => `<option value="${escapeHtml(team.id || '')}">${escapeHtml(team.name || 'チーム')}</option>`))
    .join('');
  const fallback = String(options[0]?.id || '');
  els.inviteTeamSelect.value = options.some(team => String(team.id || '') === current) ? current : fallback;
}

function normalizeInvitationTeamOptionName(row) {
  const candidates = [
    row?.team_name,
    row?.name,
    row?.workspace_name,
    row?.team_label,
    row?.label,
    row?.title
  ];
  for (const candidate of candidates) {
    const normalized = String(candidate || '').trim();
    if (normalized) return normalized;
  }
  return '';
}

async function resolveWorkspaceTeamNameForInvitation(teamId) {
  const normalizedTeamId = String(teamId || '').trim();
  if (!normalizedTeamId) return '';

  const cachedMeta = typeof window.getCachedDropOffTeamMeta === 'function' ? window.getCachedDropOffTeamMeta(normalizedTeamId) : null;
  const cacheCandidates = [
    String(cachedMeta?.name || cachedMeta?.team_name || '').trim(),
    String(window.currentUserProfile?.team_name || '').trim(),
    normalizeInvitationTeamOptionName(window.currentWorkspaceInfo || {}),
    String(getAdminForcedTeamName?.() || '').trim()
  ].filter(Boolean).filter(v => !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(v)));
  if (cacheCandidates.length) return cacheCandidates[0];

  if (typeof loadDropOffTeamsCloud === 'function') {
    try {
      const teams = await loadDropOffTeamsCloud();
      const matched = (Array.isArray(teams) ? teams : []).find(row => String(row?.id || '').trim() === normalizedTeamId);
      const matchedName = normalizeInvitationTeamOptionName(matched || {});
      if (matchedName) return matchedName;
    } catch (error) {
      console.warn('resolveWorkspaceTeamNameForInvitation loadDropOffTeamsCloud failed:', error);
    }
  }

  try {
    const teamTable = getTableName('teams');
    const { data, error } = await supabaseClient
      .from(teamTable)
      .select('*')
      .eq('id', normalizedTeamId)
      .maybeSingle();
    if (!error && data) {
      const name = normalizeInvitationTeamOptionName(data || {});
      if (name) return name;
    }
  } catch (error) {
    console.warn('resolveWorkspaceTeamNameForInvitation team query failed:', error);
  }

  return '';
}

function resetInvitationForm() {
  if (els.inviteEmailInput) els.inviteEmailInput.value = '';
  if (els.inviteDisplayNameInput) els.inviteDisplayNameInput.value = '';
  renderInvitationRoleOptions();
  renderInvitationTeamOptions();
  if (els.inviteStatusText) els.inviteStatusText.textContent = 'pending / accepted / expired / revoked を管理します。';
}

function renderInvitationsTable() {
  if (!els.invitationsTableWrap) return;
  const rows = getVisibleInvitationRows();
  if (!rows.length) {
    els.invitationsTableWrap.innerHTML = '<p class="soft-text">招待データはまだありません。</p>';
    return;
  }

  const inviterMap = new Map((Array.isArray(allUserProfilesCache) ? allUserProfilesCache : []).map(row => [String(row.user_id || row.id || ''), row.display_name || row.email || '-']));
  const body = rows.map(row => {
    const status = normalizeInvitationStatusForUi(row.status, row.expires_at);
    const canManage = canCurrentUserManageInvitation(row);
    return `
      <tr data-invitation-id="${escapeHtml(row.id || '')}">
        <td>${escapeHtml(row.display_name || '-')}</td>
        <td>${escapeHtml(row.email || '-')}</td>
        <td><span class="chip">${escapeHtml(getInvitationRoleLabel(row.invited_role))}</span></td>
        <td>${escapeHtml(getInvitationTeamName(row.team_id))}</td>
        <td>${escapeHtml(inviterMap.get(String(row.invited_by_user_id || '')) || '-')}</td>
        <td><span class="chip">${escapeHtml(getInvitationStatusLabel(status))}</span></td>
        <td>${escapeHtml(formatDateTimeLabel(row.created_at))}</td>
        <td>${escapeHtml(formatDateTimeLabel(row.expires_at))}</td>
        <td class="action-row">
          <button class="btn danger invitation-revoke-btn" ${canManage ? '' : 'disabled'}>取消</button>
        </td>
      </tr>`;
  }).join('');

  els.invitationsTableWrap.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th>表示名</th>
          <th>メール</th>
          <th>権限</th>
          <th>所属チーム</th>
          <th>招待者</th>
          <th>状態</th>
          <th>作成日</th>
          <th>有効期限</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>`;
}

async function loadInvitationSettingsData() {
  if (!isManagerUser()) {
    allInvitationRowsCache = [];
    invitationTeamOptionsCache = [];
    renderInvitationRoleOptions();
    renderInvitationTeamOptions();
    renderInvitationsTable();
    return;
  }

  try {
    const settingsTeamId = String(
      window.currentWorkspaceTeamId
      || getCurrentWorkspaceTeamIdSync()
      || (typeof ensureDropOffWorkspaceId === 'function' ? await ensureDropOffWorkspaceId() : '')
      || ''
    ).trim() || null;

    let settingsTeamName = settingsTeamId ? await resolveWorkspaceTeamNameForInvitation(settingsTeamId) : '';
    if (!settingsTeamName) {
      settingsTeamName = String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || window.currentUserProfile?.team_name || getAdminForcedTeamName() || '').trim();
    }
    if (!settingsTeamName || /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(settingsTeamName))) {
      settingsTeamName = await resolveCurrentWorkspaceTeamLabelAsync();
    }
    if (!settingsTeamName && settingsTeamId && Array.isArray(allUserProfilesCache) && allUserProfilesCache.length) {
      const matchedProfile = allUserProfilesCache.find(row => String(row?.team_id || '').trim() === settingsTeamId);
      settingsTeamName = normalizeInvitationTeamOptionName(matchedProfile || {});
    }
    if (!settingsTeamName) settingsTeamName = '現在のワークスペース';

    invitationTeamOptionsCache = settingsTeamId ? [{ id: settingsTeamId, name: settingsTeamName }] : [];
    renderInvitationRoleOptions();
    renderInvitationTeamOptions();

    if (typeof loadInvitationsCloud === 'function') {
      allInvitationRowsCache = await loadInvitationsCloud(settingsTeamId);
    } else {
      allInvitationRowsCache = [];
    }

    if (typeof loadCurrentTeamPlanUsage === 'function') {
      await loadCurrentTeamPlanUsage();
      renderPlanInfo();
    }

    renderInvitationsTable();
    if (els.inviteStatusText) {
      const visibleCount = getVisibleInvitationRows().length;
      els.inviteStatusText.textContent = `招待一覧 ${visibleCount}件 / pending・accepted・expired・revoked を管理します。`;
    }
  } catch (error) {
    console.error(error);
    if (els.invitationsTableWrap) {
      els.invitationsTableWrap.innerHTML = `<p class="soft-text">招待一覧の読込に失敗しました: ${escapeHtml(error?.message || String(error))}</p>`;
    }
  }
}

async function submitInvitationFromSettings() {
  if (!isManagerUser() || typeof createInvitationCloud !== 'function') {
    alert('この操作はオーナーまたは管理者のみ行えます。');
    return;
  }

  const email = String(els.inviteEmailInput?.value || '').trim().toLowerCase();
  const displayName = String(els.inviteDisplayNameInput?.value || '').trim();
  const invitedRole = normalizeInvitationRoleForUi(els.inviteRoleSelect?.value || 'user');
  const teamId = String(els.inviteTeamSelect?.value || '').trim();

  if (!email) {
    alert('招待先メールアドレスを入力してください。');
    return;
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    alert('メールアドレスの形式を確認してください。');
    return;
  }
  if (!displayName) {
    alert('表示名を入力してください。');
    return;
  }
  if (!teamId) {
    alert('所属チームを選択してください。');
    return;
  }
  if (!canCurrentUserInviteRole(invitedRole)) {
    alert('この権限では、その役割を招待できません。');
    return;
  }

  const inviteLimitCheck = canAddMemberInvite();
  if (!inviteLimitCheck.allowed) {
    alert(inviteLimitCheck.reason || 'このプランではこれ以上招待できません。');
    return;
  }

  try {
    const created = await createInvitationCloud({
      email,
      display_name: displayName,
      invited_role: invitedRole,
      team_id: teamId,
      invited_by_user_id: currentUser?.id || currentUserProfile?.id || null
    });
    allInvitationRowsCache = [created, ...allInvitationRowsCache.filter(row => String(row.id || '') !== String(created.id || ''))];
    renderInvitationsTable();
    resetInvitationForm();
    if (typeof loadCurrentTeamPlanUsage === 'function') {
      await loadCurrentTeamPlanUsage();
      renderPlanInfo();
    }
    if (els.inviteStatusText) {
      els.inviteStatusText.textContent = `${created.email} を ${getInvitationRoleLabel(created.invited_role)} として pending 登録しました。`;
    }
    alert('招待を登録しました。');
  } catch (error) {
    console.error(error);
    alert(error?.message || '招待の登録に失敗しました。');
  }
}

async function revokeInvitationFromTable(invitationId) {
  const row = (Array.isArray(allInvitationRowsCache) ? allInvitationRowsCache : []).find(item => String(item.id || '') === String(invitationId || ''));
  if (!row) return;
  if (!canCurrentUserManageInvitation(row) || typeof revokeInvitationCloud !== 'function') {
    alert('この招待は操作できません。');
    return;
  }
  if (!window.confirm(`この招待を取り消しますか？

${row.email || '-'}`)) return;

  try {
    const updated = await revokeInvitationCloud(invitationId);
    allInvitationRowsCache = allInvitationRowsCache.map(item => String(item.id || '') === String(invitationId) ? { ...item, ...updated } : item);
    if (typeof loadCurrentTeamPlanUsage === 'function') {
      await loadCurrentTeamPlanUsage();
      renderPlanInfo();
    }
    renderInvitationsTable();
    if (els.inviteStatusText) {
      els.inviteStatusText.textContent = `${row.email} の招待を取り消しました。`;
    }
  } catch (error) {
    console.error(error);
    alert(error?.message || '招待の取消に失敗しました。');
  }
}

function bindInvitationTableEvents() {
  if (!els.invitationsTableWrap) return;
  els.invitationsTableWrap.addEventListener('click', async event => {
    const revokeBtn = event.target.closest('.invitation-revoke-btn');
    if (revokeBtn) {
      const invitationId = revokeBtn.closest('tr')?.dataset.invitationId || '';
      if (invitationId) await revokeInvitationFromTable(invitationId);
    }
  });
}

function activateTab(tabId) {
  if (window.currentWorkspaceSuspended && !window.isPlatformAdminUser) {
    renderSuspendedWorkspaceMode();
    return;
  }
  const allowedTabs = getAllowedTabsForRole(getCurrentAppRole());
  const safeTabId = allowedTabs.has(tabId) ? tabId : 'homeTab';
  document.querySelectorAll(".main-tab").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.tab === safeTabId);
  });

  document.querySelectorAll(".page-panel").forEach(panel => {
    panel.classList.toggle("active", panel.id === safeTabId);
  });

  if (safeTabId === "vehiclesTab") {
    const targetDate = els.dispatchDate?.value || todayStr();
    forceResetMileageReportInputs(targetDate);
    window.requestAnimationFrame(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()));
    window.setTimeout(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()), 0);
    window.setTimeout(() => forceResetMileageReportInputs(els.dispatchDate?.value || todayStr()), 120);
  }
  if (safeTabId === 'castsTab' || safeTabId === 'castSearchTab') {
    if (typeof loadCasts === 'function') {
      Promise.resolve(loadCasts())
        .then(() => {
          syncCastBlankMetricsUi();
          try { refreshCastGoogleApiQuotaUi(); } catch (_) {}
          return refreshLiveOriginMetricsForDisplay(safeTabId);
        })
        .catch(err => console.warn('cast tab origin refresh failed:', err));
    } else {
      syncCastBlankMetricsUi();
      try { refreshCastGoogleApiQuotaUi(); } catch (_) {}
      Promise.resolve(refreshLiveOriginMetricsForDisplay(safeTabId)).catch(err => console.warn('cast tab live recalc failed:', err));
    }
  }
  if (safeTabId === 'homeTab' || safeTabId === 'operationTab' || safeTabId === 'scheduleTab' || safeTabId === 'actualTab') {
    Promise.resolve(refreshLiveOriginMetricsForDisplay(safeTabId)).catch(err => console.warn('display live recalc failed:', err));
  }
  if (safeTabId === 'settingsTab') {
    if (typeof renderPlanInfo === 'function') renderPlanInfo();
    if (typeof loadCurrentTeamPlan === 'function') {
      Promise.resolve(loadCurrentTeamPlan()).catch(err => console.warn('team plan refresh failed:', err));
    }
    if (typeof loadProfilesForSettings === 'function') loadProfilesForSettings();
    if (typeof loadInvitationSettingsData === 'function') loadInvitationSettingsData();
    if (typeof initializeOriginManagement === 'function') initializeOriginManagement().then(() => refreshCurrentOriginSlotFromServer({ reloadOrigins: true }).catch(() => {}));
  }
  if (safeTabId === 'platformAdminTab' && window.isPlatformAdminUser) {
    loadPlatformAdminTeams();
  }
}

function setupTabs() {
  document.querySelectorAll(".main-tab").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  document.querySelectorAll(".go-tab-btn").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.goTab));
  });
}

function normalizeOriginSlotNo(value) {
  const num = Number(value);
  return Number.isInteger(num) && num >= 1 && num <= ORIGIN_SLOT_LIMIT ? num : null;
}

function getActiveOriginSlotStorageKey() {
  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
  return teamId ? `${ACTIVE_ORIGIN_SLOT_STORAGE_KEY}_${teamId}` : ACTIVE_ORIGIN_SLOT_STORAGE_KEY;
}

function loadStoredActiveOriginSlotNo() {
  const dbManagedSlot = getCurrentWorkspaceOriginSlotFromState();
  if (dbManagedSlot) return dbManagedSlot;

  const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
  const cachedMetaSlot = teamId && typeof window.getCachedDropOffTeamMeta === 'function'
    ? normalizeOriginSlotNo(window.getCachedDropOffTeamMeta(teamId)?.current_origin_slot)
    : null;
  if (cachedMetaSlot) return cachedMetaSlot;

  try {
    const scopedKey = getActiveOriginSlotStorageKey();
    const scopedValue = window.localStorage.getItem(scopedKey);
    const normalizedScoped = normalizeOriginSlotNo(scopedValue);
    if (normalizedScoped) return normalizedScoped;

    const teamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
    if (!teamId) {
      return normalizeOriginSlotNo(window.localStorage.getItem(ACTIVE_ORIGIN_SLOT_STORAGE_KEY));
    }
    return null;
  } catch (_) {
    return null;
  }
}

function saveStoredActiveOriginSlotNo(slotNo) {
  try {
    const normalized = normalizeOriginSlotNo(slotNo);
    const scopedKey = getActiveOriginSlotStorageKey();
    if (normalized) {
      window.localStorage.setItem(scopedKey, String(normalized));
    } else {
      window.localStorage.removeItem(scopedKey);
    }
  } catch (_) {}
}

async function syncTeamCurrentOriginSlot(slotNo, options = {}) {
  const normalized = normalizeOriginSlotNo(slotNo);
  const teamId = getCurrentWorkspaceTeamIdForUi();
  saveStoredActiveOriginSlotNo(normalized);
  setCurrentWorkspaceMetaState({ id: teamId || window.currentWorkspaceInfo?.id || null, current_origin_slot: normalized });
  try {
    if (teamId && typeof window.cacheDropOffTeamMeta === 'function') {
      const label = getCurrentWorkspaceTeamLabel();
      window.cacheDropOffTeamMeta({ id: teamId, name: label && label !== '-' ? label : '', current_origin_slot: normalized });
    }
  } catch (_) {}

  if (!teamId || typeof window.updateDropOffTeamCurrentOriginSlot !== 'function') {
    return { data: window.currentWorkspaceInfo || null, error: null };
  }

  try {
    const result = await window.updateDropOffTeamCurrentOriginSlot(teamId, normalized);
    if (result?.data) setCurrentWorkspaceMetaState(result.data);

    const confirmedMeta = await loadCurrentWorkspaceTeamMeta(true);
    const confirmedSlot = normalizeOriginSlotNo(confirmedMeta?.current_origin_slot ?? getCurrentWorkspaceOriginSlotFromState());
    if (confirmedSlot && confirmedSlot !== normalized) {
      const retry = await window.updateDropOffTeamCurrentOriginSlot(teamId, normalized);
      if (retry?.data) setCurrentWorkspaceMetaState(retry.data);
    }

    const finalMeta = await loadCurrentWorkspaceTeamMeta(true);
    if (finalMeta) setCurrentWorkspaceMetaState(finalMeta);

    if (result?.error && options?.silent !== true) {
      console.warn('syncTeamCurrentOriginSlot warning:', result.error);
    }
    return result || { data: null, error: null };
  } catch (error) {
    if (options?.silent !== true) console.warn('syncTeamCurrentOriginSlot failed:', error);
    return { data: null, error };
  }
}

function getOriginDisplayLabel(label) {
  const raw = String(label || "").trim();
  if (!raw) return "起点";
  const lowered = raw.toLowerCase();
  if (lowered === "themis" || lowered === "dropoff" || lowered === "drop off") return "起点";
  return raw;
}

function getCurrentOriginRuntime() {
  const strictOrigin = getSingleCurrentOriginRuntimeRecord();
  if (strictOrigin) {
    const slotNo = normalizeOriginSlotNo(strictOrigin.slot_no);
    if (slotNo && normalizeOriginSlotNo(activeOriginSlotNo) !== slotNo) {
      activeOriginSlotNo = slotNo;
    }
    return strictOrigin;
  }

  const fallbackSlot = normalizeOriginSlotNo(
    getCurrentWorkspaceOriginSlotFromState()
    || activeOriginSlotNo
  );

  return {
    slot_no: fallbackSlot,
    name: getOriginDisplayLabel(ORIGIN_LABEL),
    address: "",
    lat: Number(ORIGIN_LAT),
    lng: Number(ORIGIN_LNG)
  };
}

function normalizeOriginRecord(row = {}) {
  const safeRow = row && typeof row === 'object' ? row : {};
  const lat = Number(safeRow.lat ?? safeRow.latitude ?? safeRow.origin_lat ?? NaN);
  const lng = Number(safeRow.lng ?? safeRow.longitude ?? safeRow.origin_lng ?? NaN);
  return {
    ...safeRow,
    slot_no: normalizeOriginSlotNo(safeRow.slot_no ?? safeRow.slotNo ?? safeRow.order_no ?? safeRow.position ?? safeRow.sort_order) || null,
    name: String(safeRow.name || safeRow.label || safeRow.origin_name || safeRow.title || "起点").trim() || "起点",
    address: String(safeRow.address || safeRow.memo || safeRow.note || "").trim(),
    lat: Number.isFinite(lat) ? lat : null,
    lng: Number.isFinite(lng) ? lng : null,
    is_active: safeRow.is_active !== false
  };
}

function filterOriginsForCurrentUser(rows = []) {
  const list = Array.isArray(rows) ? rows : [];
  return list;
}

function loadLocalOriginBackupRows() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(getOriginLocalBackupStorageKey()) || "[]");
    return Array.isArray(parsed) ? parsed.map(normalizeOriginRecord).filter(row => normalizeOriginSlotNo(row?.slot_no)) : [];
  } catch (_) {
    return [];
  }
}

function saveLocalOriginBackupRows(rows = []) {
  try {
    const sanitized = (Array.isArray(rows) ? rows : [])
      .map(normalizeOriginRecord)
      .filter(row => normalizeOriginSlotNo(row?.slot_no))
      .map(row => ({
        id: row?.id || row?.cloud_id || null,
        team_id: row?.team_id || getCurrentWorkspaceTeamIdSync() || null,
        is_default: row?.is_default === true,
        slot_no: normalizeOriginSlotNo(row?.slot_no),
        name: String(row?.name || "起点").trim() || "起点",
        address: String(row?.address || "").trim(),
        lat: Number(row?.lat),
        lng: Number(row?.lng),
        updated_at: row?.updated_at || new Date().toISOString(),
        storage: row?.storage || "local"
      }));
    window.localStorage.setItem(getOriginLocalBackupStorageKey(), JSON.stringify(sanitized));
  } catch (error) {
    console.warn("local origin backup save skipped:", error);
  }
}

function upsertLocalOriginBackupRow(row = {}) {
  const slotNo = normalizeOriginSlotNo(row?.slot_no);
  if (!slotNo) return;
  const list = loadLocalOriginBackupRows().filter(item => normalizeOriginSlotNo(item?.slot_no) !== slotNo);
  list.push({
    ...normalizeOriginRecord(row),
    slot_no: slotNo,
    updated_at: new Date().toISOString(),
    storage: "local"
  });
  saveLocalOriginBackupRows(list);
}

function removeLocalOriginBackupRow(slotNo) {
  const normalized = normalizeOriginSlotNo(slotNo);
  if (!normalized) return;
  const list = loadLocalOriginBackupRows().filter(item => normalizeOriginSlotNo(item?.slot_no) !== normalized);
  saveLocalOriginBackupRows(list);
}

function mergeOriginRows(dbRows = [], localRows = []) {
  const merged = new Map();
  const cloudById = new Map();

  for (const row of Array.isArray(dbRows) ? dbRows : []) {
    const normalized = normalizeOriginRecord(row);
    if (normalized?.id) cloudById.set(String(normalized.id), normalized);
    const slotNo = normalizeOriginSlotNo(normalized?.slot_no);
    if (slotNo) merged.set(slotNo, { ...normalized, storage: "cloud" });
  }

  for (const row of Array.isArray(localRows) ? localRows : []) {
    const normalized = normalizeOriginRecord(row);
    const slotNo = normalizeOriginSlotNo(normalized?.slot_no);
    if (!slotNo || merged.has(slotNo)) continue;
    const cloudMatch = normalized?.id ? cloudById.get(String(normalized.id)) : null;
    merged.set(slotNo, {
      ...cloudMatch,
      ...normalized,
      id: normalized?.id || cloudMatch?.id || null,
      team_id: normalized?.team_id || cloudMatch?.team_id || getCurrentWorkspaceTeamIdSync() || null,
      is_default: normalized?.is_default === true || cloudMatch?.is_default === true,
      storage: cloudMatch ? "cloud" : (normalized.storage || "local")
    });
  }

  return [...merged.values()].sort((a, b) => normalizeOriginSlotNo(a?.slot_no) - normalizeOriginSlotNo(b?.slot_no));
}

async function saveOriginRecordSmart(existing, payload) {
  const tableName = getTableName('origins');
  const exactPayload = {
    team_id: payload?.team_id || getCurrentWorkspaceTeamIdSync() || null,
    slot_no: normalizeOriginSlotNo(payload?.slot_no),
    name: String(payload?.name || '起点').trim() || '起点',
    address: String(payload?.address || '').trim() || null,
    lat: Number(payload?.lat),
    lng: Number(payload?.lng),
    updated_at: new Date().toISOString()
  };

  if (!Number.isFinite(exactPayload.lat) || !Number.isFinite(exactPayload.lng)) {
    return { ok: false, error: new Error('origin coordinates are invalid') };
  }
  if (!exactPayload.team_id) {
    return { ok: false, error: new Error('origin team_id is missing') };
  }
  if (!exactPayload.slot_no) {
    return { ok: false, error: new Error('origin slot_no is missing') };
  }

  let targetId = existing?.id && isUuidLike(existing.id) ? String(existing.id) : null;

  if (!targetId) {
    try {
      const lookup = await supabaseClient
        .from(tableName)
        .select('id, team_id, slot_no, updated_at')
        .eq('team_id', exactPayload.team_id)
        .eq('slot_no', exactPayload.slot_no)
        .order('updated_at', { ascending: false })
        .order('id', { ascending: true })
        .limit(1)
        .maybeSingle();
      if (!lookup?.error && lookup?.data?.id && isUuidLike(lookup.data.id)) {
        targetId = String(lookup.data.id);
      }
    } catch (_) {}
  }

  if (targetId) {
    const { data, error } = await supabaseClient
      .from(tableName)
      .update(exactPayload)
      .eq('id', targetId)
      .select('*')
      .single();

    if (!error) {
      return {
        ok: true,
        mode: 'update',
        storage: 'cloud',
        payload: exactPayload,
        data: data || { ...existing, ...exactPayload, id: targetId }
      };
    }
  }

  const insertResult = await insertSelectSingleWithColumnFallback(tableName, exactPayload);
  if (!insertResult?.error) {
    return {
      ok: true,
      mode: 'insert',
      storage: 'cloud',
      payload: insertResult.payload || exactPayload,
      data: insertResult.data || { ...exactPayload, id: insertResult?.data?.id || null }
    };
  }

  return { ok: false, error: insertResult?.error || new Error('origin save failed') };
}

function openOriginGoogleMapFromSettings() {
  const address = String(els.originAddressInput?.value || '').trim();
  const parsed = parseLatLngText(els.originLatLngInput?.value || '');
  let url = '';
  if (address) {
    url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;
  } else if (parsed && isValidLatLng(parsed.lat, parsed.lng)) {
    url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(`${parsed.lat},${parsed.lng}`)}`;
  } else {
    const current = getCurrentOriginRuntime();
    url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(`${current.lat},${current.lng}`)}`;
  }
  window.open(url, '_blank', 'noopener,noreferrer');
}

function getOriginRowBySlot(slotNo) {
  const normalized = normalizeOriginSlotNo(slotNo);
  return (Array.isArray(allOriginsCache) ? allOriginsCache : []).find(row => normalizeOriginSlotNo(row?.slot_no) === normalized) || null;
}

function getFirstAvailableOriginSlotNo() {
  const used = new Set((Array.isArray(allOriginsCache) ? allOriginsCache : []).map(row => normalizeOriginSlotNo(row?.slot_no)).filter(Boolean));
  const visibleLimit = getVisibleOriginSlotLimit();
  applyOriginSlotUi();

  for (let slotNo = 1; slotNo <= visibleLimit; slotNo++) {
    if (!used.has(slotNo)) return slotNo;
  }
  return 1;
}

function refreshOriginStatusUi() {
  applyOriginSlotUi();
  const origin = getCurrentOriginRuntime();
  if (els.originLabelText) {
    els.originLabelText.value = origin.name || "起点";
  }
  if (els.originStatusText) {
    const coords = Number.isFinite(origin.lat) && Number.isFinite(origin.lng)
      ? ` (${origin.lat.toFixed(6)}, ${origin.lng.toFixed(6)})`
      : "";
    const slotLabel = activeOriginSlotNo ? `起点${activeOriginSlotNo}` : "未保存";
    els.originStatusText.textContent = `現在の起点: ${slotLabel} / ${origin.name || "起点"}${coords}`;
  }
  updateCastDistanceHint();
}

function fillOriginForm(row = {}, options = {}) {
  const visibleLimit = getVisibleOriginSlotLimit();
  let slotNo = normalizeOriginSlotNo(options.slotNo ?? row?.slot_no) || getFirstAvailableOriginSlotNo();
  if (slotNo > visibleLimit) slotNo = 1;
  editingOriginSlotNo = slotNo;
  if (els.originSlotSelect) els.originSlotSelect.value = String(slotNo);
  if (els.originNameInput) els.originNameInput.value = String(row?.name || "").trim();
  if (els.originAddressInput) els.originAddressInput.value = String(row?.address || "").trim();
  if (els.originLatLngInput) {
    els.originLatLngInput.value = Number.isFinite(Number(row?.lat)) && Number.isFinite(Number(row?.lng))
      ? `${Number(row.lat)}, ${Number(row.lng)}`
      : "";
  }
  if (els.cancelOriginEditBtn) {
    els.cancelOriginEditBtn.classList.toggle("hidden", !row || (!row.name && !row.address && row.lat == null && row.lng == null));
  }
}

function resetOriginForm() {
  const preferredSlot = normalizeOriginSlotNo(activeOriginSlotNo) || getFirstAvailableOriginSlotNo();
  const existing = getOriginRowBySlot(preferredSlot);
  if (existing) {
    fillOriginForm(existing, { slotNo: preferredSlot });
    return;
  }
  editingOriginSlotNo = preferredSlot;
  if (els.originSlotSelect) els.originSlotSelect.value = String(preferredSlot);
  if (els.originNameInput) els.originNameInput.value = "";
  if (els.originAddressInput) els.originAddressInput.value = "";
  if (els.originLatLngInput) els.originLatLngInput.value = "";
  if (els.cancelOriginEditBtn) els.cancelOriginEditBtn.classList.add("hidden");
}

function renderOriginSlots() {
  if (!els.originSlotsWrap) return;
  const list = Array.isArray(allOriginsCache) ? allOriginsCache : [];
  const isReadonlyUser = isReadonlyUserRole();

  let html = `
    <table class="data-table">
      <thead>
        <tr>
          <th>保存枠</th>
          <th>状態</th>
          <th>起点名</th>
          <th>住所</th>
          <th>座標</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
  `;

  const visibleLimit = getVisibleOriginSlotLimit();
  applyOriginSlotUi();

  for (let slotNo = 1; slotNo <= visibleLimit; slotNo++) {
    const row = list.find(item => normalizeOriginSlotNo(item?.slot_no) === slotNo) || null;
    const isActive = normalizeOriginSlotNo(activeOriginSlotNo) === slotNo;
    const slotLabel = `起点${slotNo}`;
    const coordText = row && Number.isFinite(Number(row.lat)) && Number.isFinite(Number(row.lng))
      ? `${Number(row.lat).toFixed(6)}, ${Number(row.lng).toFixed(6)}`
      : "-";
    html += `
      <tr>
        <td>${slotLabel}</td>
        <td>${isActive ? '<span class="chip">使用中</span>' : '<span class="muted">待機</span>'}</td>
        <td>${escapeHtml(row?.name || "未保存")}</td>
        <td>${escapeHtml(row?.address || "-")}</td>
        <td>${escapeHtml(coordText)}</td>
        <td class="actions-cell">
          <button class="btn ghost origin-use-btn" data-slot="${slotNo}" ${row ? "" : "disabled"}>使う</button>
          <button class="btn ghost origin-edit-btn" data-slot="${slotNo}">${row ? "編集" : "入力"}</button>
          <button class="btn danger origin-delete-btn" data-slot="${slotNo}" ${row ? "" : "disabled"}>削除</button>
        </td>
      </tr>
    `;
  }

  html += `</tbody></table>`;
  els.originSlotsWrap.innerHTML = html;

  if (!els.originSlotsWrap.__boundOriginSlotActions) {
    els.originSlotsWrap.addEventListener('click', async (event) => {
      const btn = event.target.closest('button[data-slot]');
      if (!btn || !els.originSlotsWrap.contains(btn)) return;
      const slotNo = Number(btn.dataset.slot);
      if (!slotNo) return;
      if (btn.classList.contains('origin-use-btn')) {
        await useSavedOriginSlot(slotNo);
        return;
      }
      if (btn.classList.contains('origin-edit-btn')) {
        const row = getOriginRowBySlot(slotNo);
        fillOriginForm(row || {}, { slotNo });
        if (els.originName) els.originName.focus();
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }
      if (btn.classList.contains('origin-delete-btn')) {
        await deleteSavedOriginSlot(slotNo);
      }
    });
    els.originSlotsWrap.__boundOriginSlotActions = true;
  }
}

function applyRuntimeOrigin(row = {}, options = {}) {
  const nextLabel = getOriginDisplayLabel(row?.name || row?.label || row?.origin_name || DEFAULT_ORIGIN_LABEL || "起点");
  const nextLat = Number(row?.lat ?? row?.latitude ?? DEFAULT_ORIGIN_LAT);
  const nextLng = Number(row?.lng ?? row?.longitude ?? DEFAULT_ORIGIN_LNG);
  if (!Number.isFinite(nextLat) || !Number.isFinite(nextLng)) {
    alert('起点の座標が不正です。');
    return false;
  }

  ORIGIN_LABEL = nextLabel;
  ORIGIN_LAT = nextLat;
  ORIGIN_LNG = nextLng;
  activeOriginSlotNo = normalizeOriginSlotNo(options.slotNo ?? row?.slot_no ?? activeOriginSlotNo);

  if (window.APP_CONFIG) {
    window.APP_CONFIG.ORIGIN_LABEL = nextLabel;
    window.APP_CONFIG.ORIGIN_LAT = nextLat;
    window.APP_CONFIG.ORIGIN_LNG = nextLng;
  }

  if (options.persist !== false) {
    saveStoredActiveOriginSlotNo(activeOriginSlotNo);
  }

  refreshOriginStatusUi();
  renderOriginSlots();
  return true;
}

function getOriginDraftFromForm() {
  const slotNo = normalizeOriginSlotNo(els.originSlotSelect?.value || editingOriginSlotNo || activeOriginSlotNo || 1);
  const name = String(els.originNameInput?.value || "").trim() || "起点";
  const address = String(els.originAddressInput?.value || "").trim();
  const parsed = parseLatLngText(els.originLatLngInput?.value || "");
  if (!parsed) {
    alert('起点座標を「緯度, 経度」の形式で入力してください。');
    return null;
  }
  return {
    slot_no: slotNo,
    name,
    address,
    lat: Number(parsed.lat),
    lng: Number(parsed.lng)
  };
}

async function fetchOriginLatLngFromAddress() {
  const address = String(els.originAddressInput?.value || "").trim();
  if (!address) {
    alert('住所を入力してください。');
    return;
  }
  const access = await ensureGoogleApiCoordinateLookupAccess({ actionLabel: 'API座標取得', consume: true });
  if (!access.allowed) {
    alert(access.reason || '本日のAPI座標取得上限に達しました。');
    return;
  }
  const button = els.fetchOriginLatLngBtn;
  if (button) button.disabled = true;
  try {
    const geocoded = await geocodeAddressGoogle(address);
    if (!geocoded || !isValidLatLng(geocoded.lat, geocoded.lng)) {
      alert('住所から座標を取得できませんでした。');
      return;
    }
    if (els.originLatLngInput) {
      els.originLatLngInput.value = `${Number(geocoded.lat)}, ${Number(geocoded.lng)}`;
    }
  } catch (error) {
    console.error('fetchOriginLatLngFromAddress error:', error);
    alert('住所から座標を取得できませんでした。');
  } finally {
    if (button) button.disabled = false;
    try { refreshCastGoogleApiQuotaUi(); } catch (_) {}
  }
}

async function loadOriginsForSettings() {
  const workspaceTeamId = typeof ensureDropOffWorkspaceId === 'function'
    ? await ensureDropOffWorkspaceId()
    : getCurrentWorkspaceTeamIdForUi();
  const safeTeamId = String(workspaceTeamId || getCurrentWorkspaceTeamIdForUi() || '').trim() || null;

  if (!safeTeamId) {
    allOriginsCache = [];
    renderOriginSlots();
    refreshOriginStatusUi();
    if (typeof renderPlanInfo === 'function') renderPlanInfo();
    return [];
  }

  if (safeTeamId !== String(window.currentWorkspaceTeamId || '').trim()) {
    try { window.currentWorkspaceTeamId = safeTeamId; } catch (_) {}
  }

  const { data, error } = await selectRowsClientSideSafe(
    getTableName('origins'),
    'origins',
    [
      { column: 'slot_no', ascending: true },
      { column: 'updated_at', ascending: false },
      { column: 'id', ascending: true }
    ],
    { teamId: safeTeamId }
  );

  if (error) {
    if (isMissingTableError(error)) {
      warnMissingTableOnce('origins', error);
      allOriginsCache = [];
      renderOriginSlots();
      refreshOriginStatusUi();
      if (typeof renderPlanInfo === 'function') renderPlanInfo();
      return [];
    }
    console.error('loadOriginsForSettings error:', error);
    return Array.isArray(allOriginsCache) ? allOriginsCache : [];
  }

  const normalized = filterOriginsForCurrentUser((data || []).map(normalizeOriginRecord))
    .filter(row => String(row?.team_id || '').trim() === safeTeamId)
    .filter(row => Number.isFinite(Number(row.lat)) && Number.isFinite(Number(row.lng)));

  const localRows = normalized.length > 0 ? [] : loadLocalOriginBackupRows();
  const mergedOrigins = mergeOriginRows(normalized, localRows)
    .filter(row => String(row?.team_id || safeTeamId).trim() === safeTeamId)
    .filter(row => Number.isFinite(Number(row.lat)) && Number.isFinite(Number(row.lng)));

  const slotSeen = new Set();
  allOriginsCache = mergedOrigins.filter(row => {
    const slotNo = normalizeOriginSlotNo(row?.slot_no);
    if (!slotNo || slotSeen.has(slotNo)) return false;
    slotSeen.add(slotNo);
    return true;
  });

  renderOriginSlots();
  refreshOriginStatusUi();
  if (typeof renderPlanInfo === 'function') renderPlanInfo();
  return allOriginsCache;
}

let originSyncRefreshPromise = null;
let originSyncWatcherTimer = null;

function isSettingsTabActive() {
  const panel = document.getElementById('settingsTab');
  return !!(panel && panel.classList.contains('active'));
}

async function refreshCurrentOriginSlotFromServer(options = {}) {
  if (originSyncRefreshPromise && options.force !== true) return originSyncRefreshPromise;
  originSyncRefreshPromise = (async () => {
    const meta = await loadCurrentWorkspaceTeamMeta(true);
    const serverSlot = normalizeOriginSlotNo(meta?.current_origin_slot ?? getCurrentWorkspaceOriginSlotFromState());
    const shouldReloadOrigins = options.reloadOrigins === true || (serverSlot && !getOriginRowBySlot(serverSlot));
    if (shouldReloadOrigins) {
      await loadOriginsForSettings();
    }

    if (serverSlot) {
      const serverRow = getOriginRowBySlot(serverSlot) || (Array.isArray(allOriginsCache) ? allOriginsCache.find(row => normalizeOriginSlotNo(row?.slot_no) === serverSlot) : null);
      if (serverRow && normalizeOriginSlotNo(activeOriginSlotNo) !== serverSlot) {
        applyRuntimeOrigin(serverRow, { slotNo: serverSlot, persist: true });
      }
      return { slotNo: serverSlot, row: serverRow || null, meta };
    }

    return { slotNo: null, row: null, meta };
  })().finally(() => {
    originSyncRefreshPromise = null;
  });
  return originSyncRefreshPromise;
}

function ensureOriginSyncWatcher() {
  if (originSyncWatcherTimer) return;
  originSyncWatcherTimer = window.setInterval(() => {
    if (!isSettingsTabActive()) return;
    refreshCurrentOriginSlotFromServer({ reloadOrigins: true }).catch(err => console.warn('origin sync watcher warning:', err));
  }, 4000);
}

function bindOriginSyncEvents() {
  if (window.__DROP_OFF_ORIGIN_SYNC_BOUND__) return;
  window.__DROP_OFF_ORIGIN_SYNC_BOUND__ = true;

  window.addEventListener('focus', () => {
    if (!isSettingsTabActive()) return;
    refreshCurrentOriginSlotFromServer({ reloadOrigins: true }).catch(err => console.warn('origin sync focus warning:', err));
  });

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'visible' || !isSettingsTabActive()) return;
    refreshCurrentOriginSlotFromServer({ reloadOrigins: true }).catch(err => console.warn('origin sync visible warning:', err));
  });
}

async function initializeOriginManagement() {
  bindOriginSyncEvents();
  ensureOriginSyncWatcher();

  const workspaceMeta = await loadCurrentWorkspaceTeamMeta(true);
  const dbSlot = normalizeOriginSlotNo(workspaceMeta?.current_origin_slot ?? getCurrentWorkspaceOriginSlotFromState());
  activeOriginSlotNo = dbSlot || null;

  await loadOriginsForSettings();
  await refreshCurrentOriginSlotFromServer({ reloadOrigins: false, force: true });

  const preferred = getOriginRowBySlot(activeOriginSlotNo) || (Array.isArray(allOriginsCache) ? allOriginsCache[0] : null);
  if (preferred) {
    applyRuntimeOrigin(preferred, { slotNo: preferred.slot_no, persist: true });
  } else {
    applyRuntimeOrigin({
      name: DEFAULT_ORIGIN_LABEL,
      lat: DEFAULT_ORIGIN_LAT,
      lng: DEFAULT_ORIGIN_LNG
    }, { slotNo: null, persist: false });
  }

  resetOriginForm();
}

async function refreshAfterOriginChange() {
  refreshOriginStatusUi();
  renderOriginSlots();
  syncScheduleRendererDeps();
  await loadHomeAndAll();
  if (typeof refreshCastMetricsForCurrentOrigin === 'function') {
    await refreshCastMetricsForCurrentOrigin({ render: false });
  }
  if (typeof refreshCastFormMetricsForCurrentOrigin === 'function') {
    refreshCastFormMetricsForCurrentOrigin();
  }
  if (typeof renderCastsTable === 'function') renderCastsTable();
  if (typeof renderCastSearchResults === 'function') renderCastSearchResults();
  if (typeof renderCastSelects === 'function') renderCastSelects();
  renderManualLastVehicleInfo();
}

async function saveOriginFromSettings() {
  const draft = getOriginDraftFromForm();
  if (!draft) return;
  const slotNo = normalizeOriginSlotNo(draft.slot_no);
  if (!slotNo) {
    alert('保存枠を選択してください。');
    return;
  }

  const workspaceTeamId = typeof ensureDropOffWorkspaceId === 'function'
    ? await ensureDropOffWorkspaceId()
    : null;

  const payload = {
    slot_no: slotNo,
    name: draft.name,
    address: draft.address || null,
    lat: draft.lat,
    lng: draft.lng,
    team_id: workspaceTeamId || getCurrentWorkspaceTeamIdSync() || null,
    updated_at: new Date().toISOString()
  };

  const existing = getOriginRowBySlot(slotNo);
  const originLimitCheck = canAddOrigin({ isEditingExisting: Boolean(existing) });
  if (!originLimitCheck.allowed) {
    alert(originLimitCheck.reason);
    return;
  }

  const result = await saveOriginRecordSmart(existing, payload);

  upsertLocalOriginBackupRow({
    ...draft,
    slot_no: slotNo,
    team_id: payload.team_id,
    updated_at: new Date().toISOString(),
    storage: result?.ok ? 'cloud' : 'local'
  });

  if (!result?.ok) {
    console.warn('origin cloud save failed, local backup saved instead:', result?.error);
    await loadOriginsForSettings();
    const appliedFallback = getOriginRowBySlot(slotNo) || draft;
    applyRuntimeOrigin(appliedFallback, { slotNo, persist: true });
    resetOriginForm();
    await refreshAfterOriginChange();
    alert('クラウド保存は通りませんでしたが、この端末には保存しました。引き続き使えます。');
    return;
  }

  removeLocalOriginBackupRow(slotNo);
  await loadOriginsForSettings();
  const applied = getOriginRowBySlot(slotNo) || draft;
  applyRuntimeOrigin(applied, { slotNo, persist: true });
  await syncTeamCurrentOriginSlot(slotNo, { silent: true });
  await refreshCurrentOriginSlotFromServer({ reloadOrigins: true, force: true });
  resetOriginForm();
  await refreshAfterOriginChange();
}

async function useSavedOriginSlot(slotNo) {
  const row = getOriginRowBySlot(slotNo);
  if (!row) {
    alert(`起点${slotNo} はまだ保存されていません。`);
    return;
  }
  applyRuntimeOrigin(row, { slotNo, persist: true });
  await syncTeamCurrentOriginSlot(slotNo, { silent: true });
  await refreshCurrentOriginSlotFromServer({ reloadOrigins: true, force: true });
  await refreshAfterOriginChange();
}

async function useOriginDraftFromSettings() {
  const draft = getOriginDraftFromForm();
  if (!draft) return;
  applyRuntimeOrigin(draft, { slotNo: null, persist: false });
  await refreshAfterOriginChange();
}

async function deleteSavedOriginSlot(slotNo) {
  const row = getOriginRowBySlot(slotNo);
  if (!row) return;
  if (!window.confirm(`起点${slotNo} を削除しますか？`)) return;

  if (!row.id) {
    removeLocalOriginBackupRow(slotNo);
    await loadOriginsForSettings();
    if (normalizeOriginSlotNo(activeOriginSlotNo) === normalizeOriginSlotNo(slotNo)) {
      const fallback = (Array.isArray(allOriginsCache) ? allOriginsCache[0] : null);
      if (fallback) {
        applyRuntimeOrigin(fallback, { slotNo: fallback.slot_no, persist: true });
        await syncTeamCurrentOriginSlot(fallback.slot_no, { silent: true });
        await refreshCurrentOriginSlotFromServer({ reloadOrigins: true, force: true });
      } else {
        await syncTeamCurrentOriginSlot(null, { silent: true });
        await refreshCurrentOriginSlotFromServer({ reloadOrigins: true, force: true });
      }
    }
    resetOriginForm();
    refreshOriginStatusUi();
    renderOriginSlots();
    return;
  }

  let { error } = await supabaseClient
    .from(getTableName('origins'))
    .update({ is_active: false })
    .eq('id', row.id);

  if (error && isMissingColumnError(error) && /is_active/i.test(String(error?.message || ''))) {
    ({ error } = await supabaseClient
      .from(getTableName('origins'))
      .delete()
      .eq('id', row.id));
  }

  if (error) {
    alert(error.message || '起点の削除に失敗しました。');
    return;
  }

  removeLocalOriginBackupRow(slotNo);
  await loadOriginsForSettings();

  if (normalizeOriginSlotNo(activeOriginSlotNo) === normalizeOriginSlotNo(slotNo)) {
    const fallback = (Array.isArray(allOriginsCache) ? allOriginsCache[0] : null);
    if (fallback) {
      applyRuntimeOrigin(fallback, { slotNo: fallback.slot_no, persist: true });
    } else {
      applyRuntimeOrigin({
        name: DEFAULT_ORIGIN_LABEL,
        lat: DEFAULT_ORIGIN_LAT,
        lng: DEFAULT_ORIGIN_LNG
      }, { slotNo: null, persist: false });
      saveStoredActiveOriginSlotNo(null);
    }
    await refreshAfterOriginChange();
  } else {
    renderOriginSlots();
    refreshOriginStatusUi();
  }

  resetOriginForm();
}


function getSortedVehiclesForDisplay() {
  const list = Array.isArray(allVehiclesCache) ? [...allVehiclesCache] : [];
  return list.sort((a, b) => {
    const plateA = String(a?.plate_number || "").trim();
    const plateB = String(b?.plate_number || "").trim();

    const matchA = plateA.match(/^([A-Za-z])(\d+)?$/);
    const matchB = plateB.match(/^([A-Za-z])(\d+)?$/);

    if (matchA && matchB) {
      const alpha = matchA[1].localeCompare(matchB[1], "en", { sensitivity: "base" });
      if (alpha !== 0) return alpha;
      return Number(matchA[2] || 0) - Number(matchB[2] || 0);
    }

    const plain = plateA.localeCompare(plateB, "ja", { numeric: true, sensitivity: "base" });
    if (plain !== 0) return plain;

    return Number(a?.id || 0) - Number(b?.id || 0);
  });
}


function isPlatformAdminUiMode() {
  try {
    if (window.isPlatformAdminUser) return true;
    const forcedTeamId = String(getAdminForcedTeamId() || '').trim();
    if (forcedTeamId) return true;
    const search = String(window.location?.search || '');
    if (/(?:[?&])platform_admin=1(?:&|$)/.test(search)) return true;
    if (els.platformAdminTabBtn && !els.platformAdminTabBtn.classList.contains('hidden')) return true;
  } catch (_) {}
  return false;
}

function syncHomeSetupCardVisibility() {
  try {
    const isAdminMode = isPlatformAdminUiMode();
    const isOwner = isOwnerUser();
    const originCount = Array.isArray(allOriginsCache) ? allOriginsCache.length : 0;
    const vehicleCount = Array.isArray(allVehiclesCache) ? allVehiclesCache.length : 0;
    const castCount = Array.isArray(allCastsCache) ? allCastsCache.length : 0;
    const isSetupComplete = originCount > 0 && vehicleCount > 0 && castCount > 0;

    let onboardingRequested = false;
    try {
      const params = new URLSearchParams(String(window.location?.search || ''));
      onboardingRequested = params.get('onboarding') === '1';
      if (!isOwner || isAdminMode || isSetupComplete) {
        if (onboardingRequested) {
          params.delete('onboarding');
          const qs = params.toString();
          const nextUrl = `${window.location.pathname}${qs ? `?${qs}` : ''}${window.location.hash || ''}`;
          window.history.replaceState({}, '', nextUrl);
          onboardingRequested = false;
        }
      }
    } catch (_) {}

    const shouldShow = isOwner && !isAdminMode && onboardingRequested && !isSetupComplete;
    const cards = Array.from(document.querySelectorAll('#homeSetupPriorityCard, .home-setup-card, .home-setup-inline'));
    cards.forEach((card) => {
      if (!(card instanceof HTMLElement)) return;
      if (shouldShow) {
        card.classList.remove('hidden');
        card.removeAttribute('aria-hidden');
      } else {
        card.classList.add('hidden');
        card.setAttribute('aria-hidden', 'true');
      }
    });

    const homeTab = document.getElementById('homeTab');
    if (homeTab instanceof HTMLElement && !shouldShow) {
      homeTab.classList.remove('setup-priority-active');
      homeTab.removeAttribute('data-setup-step');
    }

    const originCountEl = document.getElementById('homeOriginCount');
    const vehicleCountEl = document.getElementById('homeSetupVehicleCount');
    const castCountEl = document.getElementById('homeSetupCastCount');
    if (originCountEl) originCountEl.textContent = String(originCount);
    if (vehicleCountEl) vehicleCountEl.textContent = String(vehicleCount);
    if (castCountEl) castCountEl.textContent = String(castCount);
  } catch (_) {}
}

function renderHomeSummary() {
  const actualDone = currentActualsCache.filter(x => normalizeStatus(x.status) === "done").length;
  const actualCancel = currentActualsCache.filter(x => normalizeStatus(x.status) === "cancel").length;

  if (els.homeCastCount) els.homeCastCount.textContent = String(allCastsCache.length);
  if (els.homeVehicleCount) els.homeVehicleCount.textContent = String(allVehiclesCache.length);
  if (els.homePlanCount) els.homePlanCount.textContent = String(currentPlansCache.length);
  if (els.homeActualCount) els.homeActualCount.textContent = String(currentActualsCache.length);
  if (els.homeDoneCount) els.homeDoneCount.textContent = String(actualDone);
  if (els.homeCancelCount) els.homeCancelCount.textContent = String(actualCancel);
  syncHomeSetupCardVisibility();
}

function getMileageSelectedRange() {
  const baseDate = els.dispatchDate?.value || todayStr();
  return {
    startDate: els.mileageReportStartDate?.value || getMonthStartStr(baseDate),
    endDate: els.mileageReportEndDate?.value || baseDate
  };
}

function getDashboardMonthlyRange() {
  const baseDate = els.dispatchDate?.value || todayStr();
  return {
    startDate: getMonthStartStr(baseDate),
    endDate: baseDate,
    monthKey: getMonthKey(baseDate)
  };
}

function getVehicleStatsMapForDashboardMonth(reportRows = currentDailyReportsCache) {
  const { endDate } = getDashboardMonthlyRange();
  return getUnifiedMonthlyUiStatsMap(reportRows, endDate);
}

function renderHomeMonthlyVehicleList(reportRows = currentDailyReportsCache) {
  if (!els.homeMonthlyVehicleList) return;

  const statsMap = getVehicleStatsMapForDashboardMonth(reportRows);

  els.homeMonthlyVehicleList.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.homeMonthlyVehicleList.innerHTML = `<div class="chip">車両なし</div>`;
    return;
  }

  getSortedVehiclesForDisplay().forEach(vehicle => {
    const stats = statsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };

    const row = document.createElement("div");
    row.className = "home-monthly-item";
    row.innerHTML = `
      <span class="chip">${escapeHtml(vehicle.driver_name || vehicle.plate_number || "-")}</span>
      <span class="chip">${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}</span>
      <span class="chip">帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}</span>
      <span class="chip">月間:${stats.totalDistance.toFixed(1)}km</span>
      <span class="chip">出勤:${stats.workedDays}日</span>
      <span class="chip">平均:${stats.avgDistance.toFixed(1)}km</span>
    `;
    els.homeMonthlyVehicleList.appendChild(row);
  });
}

async function refreshHomeMonthlyVehicleList() {
  if (!els.homeMonthlyVehicleList) return;
  const { startDate, endDate } = getDashboardMonthlyRange();

  if (typeof fetchDriverMileageRows !== 'function') {
    renderHomeMonthlyVehicleList(currentDailyReportsCache);
    return;
  }

  try {
    const freshRows = await fetchDriverMileageRows(startDate, endDate);
    renderHomeMonthlyVehicleList(Array.isArray(freshRows) ? freshRows : currentDailyReportsCache);
  } catch (error) {
    console.error('refreshHomeMonthlyVehicleList error:', error);
    renderHomeMonthlyVehicleList(currentDailyReportsCache);
  }
}
function resetCastForm() {
  editingCastId = null;
  if (els.castName) els.castName.value = "";
  if (els.castDistanceKm) els.castDistanceKm.value = "";
  if (els.castAddress) els.castAddress.value = "";
  if (els.castTravelMinutes) els.castTravelMinutes.value = "";
  if (els.castArea) els.castArea.value = "";
  if (els.castMemo) els.castMemo.value = "";
  if (els.castLatLngText) els.castLatLngText.value = "";
  if (els.castPhone) els.castPhone.value = "";
  if (els.castLat) els.castLat.value = "";
  if (els.castLng) els.castLng.value = "";
  lastCastGeocodeKey = "";
  setCastGeoStatus("idle", "未取得 | 住所入力後に「APIで座標取得」または Enter。未取得時は座標貼り付けで手動反映できます");
  updateCastDistanceHint();
  refreshCastGoogleApiQuotaUi();
  if (els.cancelEditBtn) els.cancelEditBtn.classList.add("hidden");
}

function fillCastForm(cast) {
  editingCastId = cast.id;
  const metrics = getCastDisplayMetrics(cast);
  const directionLabel = getCastDirectionDisplayLabel(cast, cast?.address || "") || normalizeAreaLabel(cast?.area || "");
  if (els.castName) els.castName.value = cast.name || "";
  if (els.castDistanceKm) els.castDistanceKm.value = formatCastDistanceDisplay(metrics.distance_km);
  if (els.castAddress) els.castAddress.value = cast.address || "";
  if (els.castTravelMinutes) els.castTravelMinutes.value = formatCastTravelMinutesDisplay(metrics.travel_minutes);
  if (els.castArea) els.castArea.value = directionLabel;
  if (els.castMemo) els.castMemo.value = cast.memo || "";
  if (els.castPhone) els.castPhone.value = cast.phone || "";
  if (els.castLat) els.castLat.value = cast.latitude ?? "";
  if (els.castLng) els.castLng.value = cast.longitude ?? "";
  if (els.castLatLngText) {
    els.castLatLngText.value =
      cast.latitude != null && cast.longitude != null
        ? `${cast.latitude},${cast.longitude}`
        : "";
  }
  lastCastGeocodeKey = normalizeGeocodeAddressKey(cast.address || "");
  if (cast.latitude != null && cast.longitude != null) {
    setCastGeoStatus("success", "取得済 | 保存済み座標があります");
  } else {
    setCastGeoStatus("idle", "未取得 | 住所入力後に「APIで座標取得」または Enter。未取得時は座標貼り付けで手動反映できます");
  }
  updateCastDistanceHint({ distanceKm: metrics.distance_km, invalidDistance: metrics.distance_km == null && cast.latitude != null && cast.longitude != null });
  refreshCastGoogleApiQuotaUi();
  if (els.cancelEditBtn) els.cancelEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function normalizeCastMetricDistanceValue(value) {
  const n = toNullableNumber(value);
  if (n == null || n <= 0) return null;

  let normalized = n;
  // ここは km 値を基本として扱う。
  // 明らかに meters と見なせる大きすぎる値だけ km に補正する。
  if (normalized > 3000) normalized = normalized / 1000;

  normalized = Number(Number(normalized).toFixed(1));
  if (!Number.isFinite(normalized) || normalized <= 0) return null;
  if (normalized > 3000) return null;
  return normalized;
}

function normalizeCastMetricTravelMinutesValue(value, fallbackDistanceKm = null, areaInput = "") {
  const parsedJaMinutes = parseJaDurationMinutes(value);
  const raw = parsedJaMinutes != null ? parsedJaMinutes : Number(value);
  if (!Number.isFinite(raw) || raw <= 0) return null;
  if (raw > 1440) {
    const secondsMinutes = Math.round(raw / 60);
    if (secondsMinutes > 0 && secondsMinutes <= 1440) return secondsMinutes;
  }
  const rounded = Math.round(raw);
  if (rounded > 0 && rounded <= 1440) return rounded;
  if (fallbackDistanceKm != null) {
    if (typeof estimateTravelMinutesByAreaSpeed === "function") {
      return Math.max(1, Math.round(estimateTravelMinutesByAreaSpeed(fallbackDistanceKm, areaInput)));
    }
    if (typeof estimateFallbackTravelMinutes === "function") {
      return Math.max(1, Math.round(estimateFallbackTravelMinutes(fallbackDistanceKm, areaInput)));
    }
  }
  return null;
}

function getCastDisplayMetrics(cast, addressOverride = "", runtimeOriginOverride = null) {
  const address = String(addressOverride || cast?.address || "").trim();
  const lat = toNullableNumber(cast?.latitude ?? cast?.lat);
  const lng = toNullableNumber(cast?.longitude ?? cast?.lng);
  const area = getCastDirectionDisplayLabel(cast, address) || normalizeAreaLabel(cast?.area || guessArea?.(lat, lng, address) || "無し");

  if (!(typeof isValidLatLng === "function" && isValidLatLng(lat, lng))) {
    return {
      distance_km: null,
      travel_minutes: null,
      origin: null
    };
  }

  const runtimeOrigin = runtimeOriginOverride || getStrictCurrentOriginRuntimeForLiveDisplay();
  const originLat = toNullableNumber(runtimeOrigin?.lat);
  const originLng = toNullableNumber(runtimeOrigin?.lng);
  if (!(typeof isValidLatLng === "function" && isValidLatLng(originLat, originLng))) {
    return {
      distance_km: null,
      travel_minutes: null,
      origin: runtimeOrigin || null
    };
  }

  let distanceKm = null;
  if (typeof estimateRoadKmBetweenPoints === "function") {
    distanceKm = normalizeCastMetricDistanceValue(estimateRoadKmBetweenPoints(originLat, originLng, lat, lng));
  }

  if (distanceKm == null && typeof getCastOriginMetrics === "function") {
    const live = getCastOriginMetrics(
      { latitude: lat, longitude: lng, area, address },
      address,
      {
        slot_no: runtimeOrigin?.slot_no ?? null,
        lat: originLat,
        lng: originLng,
        name: runtimeOrigin?.name || ORIGIN_LABEL || "起点"
      }
    );
    distanceKm = normalizeCastMetricDistanceValue(live?.distance_km);
  }

  let travelMinutes = null;
  if (distanceKm != null) {
    if (typeof estimateTravelMinutesByAreaSpeed === "function") {
      travelMinutes = Math.max(1, Math.round(estimateTravelMinutesByAreaSpeed(distanceKm, area || address)));
    } else if (typeof estimateFallbackTravelMinutes === "function") {
      travelMinutes = Math.max(1, Math.round(estimateFallbackTravelMinutes(distanceKm, area || address)));
    }
    travelMinutes = normalizeCastMetricTravelMinutesValue(travelMinutes, distanceKm, area || address);
  }

  return {
    distance_km: distanceKm,
    travel_minutes: travelMinutes,
    origin: runtimeOrigin
  };
}

function buildCastLiveMetricRows(casts = [], options = {}) {
  const list = Array.isArray(casts) ? casts.filter(Boolean) : [];
  const runtimeOrigin = options.runtimeOrigin || getStrictCurrentOriginRuntimeForLiveDisplay();
  return list.map(cast => ({
    cast,
    metrics: getCastDisplayMetrics(cast, options.addressOverride || "", runtimeOrigin),
    origin: runtimeOrigin || null,
    directionLabel: getCastDirectionDisplayLabel(cast, cast?.address || "") || normalizeAreaLabel(cast?.area || "")
  }));
}

async function refreshCastMetricsForCurrentOrigin(options = {}) {
  const list = Array.isArray(allCastsCache) ? allCastsCache : [];
  if (!list.length) return list;

  const runtimeOrigin = options.runtimeOrigin || getStrictCurrentOriginRuntimeForLiveDisplay();
  let changed = false;
  for (const cast of list) {
    const metrics = getCastDisplayMetrics(cast, "", runtimeOrigin);
    const nextDistance = metrics.distance_km ?? null;
    const nextMinutes = metrics.travel_minutes ?? null;
    if (toNullableNumber(cast.distance_km) !== toNullableNumber(nextDistance)) {
      cast.distance_km = nextDistance;
      changed = true;
    }
    if (toNullableNumber(cast.travel_minutes) !== toNullableNumber(nextMinutes)) {
      cast.travel_minutes = nextMinutes;
      changed = true;
    }
  }

  if ((changed || options.forceRender === true) && options.render !== false) {
    if (typeof renderCastsTable === "function") renderCastsTable();
    if (typeof renderCastSearchResults === "function") renderCastSearchResults();
    if (typeof renderCastSelects === "function") renderCastSelects();
  }
  return list;
}

function isDuplicateCast(name, address) {
  const normalizedName = String(name || "").trim();
  const normalizedAddress = String(address || "").trim();
  const editingId = typeof normalizeDispatchEntityId === "function"
    ? normalizeDispatchEntityId(editingCastId)
    : String(editingCastId || "").trim() || null;

  return allCastsCache.find(c => {
    const castId = typeof normalizeDispatchEntityId === "function"
      ? normalizeDispatchEntityId(c.id || c.cast_id)
      : String(c.id || c.cast_id || "").trim() || null;
    return (
      String(c.name || "").trim() === normalizedName &&
      String(c.address || "").trim() === normalizedAddress &&
      castId !== editingId
    );
  });
}

function renderCastsTable() {
  if (!els.castsTableBody) return;

  const isReadonlyUser = isReadonlyUserRole();
  els.castsTableBody.innerHTML = "";

  const castsTable = els.castsTableBody.closest("table");
  const castsHeaderRow = castsTable?.querySelector("thead tr");
  if (castsHeaderRow) {
    castsHeaderRow.innerHTML = isReadonlyUser
      ? `
      <th>氏名</th>
      <th>住所</th>
      <th>方面</th>
      <th>想定距離(km)</th>
      <th>片道予想時間</th>
      <th>メモ</th>
      <th>操作</th>
    `
      : `
      <th>氏名</th>
      <th>住所</th>
      <th>方面</th>
      <th>想定距離(km)</th>
      <th>片道予想時間</th>
      <th>メモ</th>
      <th>操作</th>
    `;
  }

  if (!allCastsCache.length) {
    els.castsTableBody.innerHTML = `<tr><td colspan="7" class="muted">送り先がありません</td></tr>`;
    return;
  }

  const castRows = buildCastLiveMetricRows(allCastsCache);

  castRows.forEach(({ cast, metrics, directionLabel }) => {
    const tr = document.createElement("tr");
    const actionHtml = isReadonlyUser
      ? `
        <button class="btn ghost cast-route-btn" data-id="${escapeHtml(String(cast.id || cast.cast_id || ""))}" data-address="${escapeHtml(cast.address || "")}">ルート</button>
      `
      : `
        <button class="btn ghost cast-edit-btn" data-id="${cast.id}">編集</button>
        <button class="btn ghost cast-route-btn" data-id="${escapeHtml(String(cast.id || cast.cast_id || ""))}" data-address="${escapeHtml(cast.address || "")}">ルート</button>
        <button class="btn danger cast-delete-btn" data-id="${cast.id}">削除</button>
      `;
    tr.innerHTML = `
      <td>
        ${
          buildCastMapUrl(cast)
            ? `<a href="${buildCastMapUrl(cast)}" target="_blank" rel="noopener noreferrer" class="cast-name-link">${escapeHtml(cast.name || "")} 📍</a>`
            : `${escapeHtml(cast.name || "")}`
        }
      </td>
      <td>${escapeHtml(cast.address || "")}</td>
      <td>${escapeHtml(directionLabel)}</td>
      <td>${formatCastDistanceDisplay(metrics.distance_km)}</td>
      <td>${formatCastTravelMinutesDisplay(metrics.travel_minutes)}</td>
      <td>${escapeHtml(cast.memo || "")}</td>
      <td class="actions-cell">${actionHtml}</td>
    `;
    els.castsTableBody.appendChild(tr);
  });

  if (!isReadonlyUser) {
    els.castsTableBody.querySelectorAll(".cast-edit-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const cast = allCastsCache.find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
        if (cast) fillCastForm(cast);
      });
    });

    els.castsTableBody.querySelectorAll(".cast-delete-btn").forEach(btn => {
      btn.addEventListener("click", async () => deleteCast(btn.dataset.id));
    });
  }

  els.castsTableBody.querySelectorAll(".cast-route-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
      if (cast) {
        openGoogleMap(cast.address || "", cast.latitude, cast.longitude);
        return;
      }
      openGoogleMap(btn.dataset.address || "");
    });
  });
}

function exportCastsCsv() {
  if (!ensureCsvFeatureAccess("送り先CSV")) return;
  const headers = [
    "cast_id",
    "name",
    "phone",
    "address",
    "area",
    "distance_km",
    "travel_minutes",
    "latitude",
    "longitude",
    "memo"
  ];

  const liveRows = buildCastLiveMetricRows(Array.isArray(allCastsCache) ? allCastsCache : []);
  const rows = liveRows.map(({ cast, metrics }) => {
    const liveDistanceKm = Number(metrics?.distance_km);
    const roundedDistanceKm = Number.isFinite(liveDistanceKm) && liveDistanceKm > 0
      ? Math.round(liveDistanceKm * 10) / 10
      : "";
    const liveTravelMinutes = normalizeCastMetricTravelMinutesValue(metrics?.travel_minutes) ?? "";
    const liveAreaLabel = getCastDirectionDisplayLabel(cast, cast?.address || "") || normalizeAreaLabel(cast?.area || "");
    return [
      Number(cast?.id || 0) || "",
      cast?.name || "",
      cast?.phone || "",
      cast?.address || "",
      liveAreaLabel,
      roundedDistanceKm,
      liveTravelMinutes,
      cast?.latitude ?? "",
      cast?.longitude ?? "",
      cast?.memo || ""
    ];
  });

  const csv = [headers.join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`casts_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}

function refreshCastFormMetricsForCurrentOrigin() {
  const address = String(els.castAddress?.value || '').trim();
  const latLngText = String(els.castLatLngText?.value || '').trim();
  if (!editingCastId && !address && !latLngText) {
    if (els.castLat) els.castLat.value = '';
    if (els.castLng) els.castLng.value = '';
    if (els.castDistanceKm) els.castDistanceKm.value = '';
    if (els.castTravelMinutes) els.castTravelMinutes.value = '';
    updateCastDistanceHint({ hidden: true });
    return;
  }
  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  const area = String(els.castArea?.value || '').trim();
  const hasTargetCoords = typeof isValidLatLng === 'function' && isValidLatLng(lat, lng);
  const strictOrigin = getStrictCurrentOriginRuntimeForLiveDisplay();
  const hasOriginCoords = typeof isValidLatLng === 'function' && isValidLatLng(strictOrigin?.lat, strictOrigin?.lng);

  if (!hasTargetCoords || !hasOriginCoords) {
    if (els.castDistanceKm) els.castDistanceKm.value = '';
    if (els.castTravelMinutes) els.castTravelMinutes.value = '';
    updateCastDistanceHint({ hidden: !address && !latLngText, distanceKm: null, invalidDistance: false });
    return;
  }

  const metrics = getCastDisplayMetrics({ latitude: lat, longitude: lng, area, address }, address);
  if (els.castDistanceKm) {
    els.castDistanceKm.value = formatCastDistanceDisplay(metrics.distance_km);
  }
  if (els.castTravelMinutes) {
    els.castTravelMinutes.value = formatCastTravelMinutesDisplay(metrics.travel_minutes);
  }
  updateCastDistanceHint({ distanceKm: metrics.distance_km, invalidDistance: metrics.distance_km == null });
}

async function refreshLiveOriginMetricsForDisplay(scope = '') {
  try {
    if (typeof refreshCurrentOriginSlotFromServer === 'function') {
      await refreshCurrentOriginSlotFromServer({ reloadOrigins: true, force: true });
    }
  } catch (error) {
    console.warn('refreshLiveOriginMetricsForDisplay origin sync warning:', error);
  }

  try {
    if (typeof refreshCastMetricsForCurrentOrigin === 'function' && Array.isArray(allCastsCache) && allCastsCache.length) {
      await refreshCastMetricsForCurrentOrigin({ render: false });
    }
  } catch (error) {
    console.warn('refreshLiveOriginMetricsForDisplay cast recalc warning:', error);
  }

  try {
    if (Array.isArray(currentPlansCache) && currentPlansCache.length && typeof applyCurrentOriginMetricsToDispatchRows === 'function') {
      currentPlansCache = await applyCurrentOriginMetricsToDispatchRows(currentPlansCache.map(enrichUnifiedDispatchRowWithCast));
    }
    if (Array.isArray(currentActualsCache) && currentActualsCache.length && typeof applyCurrentOriginMetricsToDispatchRows === 'function') {
      currentActualsCache = await applyCurrentOriginMetricsToDispatchRows(currentActualsCache.map(enrichUnifiedDispatchRowWithCast));
    }
  } catch (error) {
    console.warn('refreshLiveOriginMetricsForDisplay dispatch recalc warning:', error);
  }

  syncScheduleRendererDeps();
  refreshCastFormMetricsForCurrentOrigin();

  const targetScope = String(scope || '').trim();
  if (!targetScope || targetScope === 'castsTab' || targetScope === 'castSearchTab') {
    if (typeof renderCastsTable === 'function') renderCastsTable();
    if (typeof renderCastSearchResults === 'function') renderCastSearchResults();
    if (typeof renderCastSelects === 'function') renderCastSelects();
  }
  if (!targetScope || targetScope === 'scheduleTab' || targetScope === 'homeTab' || targetScope === 'operationTab') {
    if (typeof renderPlanGroupedTable === 'function') renderPlanGroupedTable();
    if (typeof renderPlansTimeAreaMatrix === 'function') renderPlansTimeAreaMatrix();
    if (typeof renderPlanSelect === 'function') renderPlanSelect();
    if (typeof renderPlanCastSelect === 'function') renderPlanCastSelect();
  }
  if (!targetScope || targetScope === 'actualTab' || targetScope === 'homeTab' || targetScope === 'operationTab') {
    if (typeof renderActualTable === 'function') renderActualTable();
    if (typeof renderActualTimeAreaMatrix === 'function') renderActualTimeAreaMatrix();
  }
  if (!targetScope || targetScope === 'homeTab' || targetScope === 'operationTab' || targetScope === 'scheduleTab' || targetScope === 'actualTab') {
    if (typeof renderHomeSummary === 'function') renderHomeSummary();
    if (typeof renderOperationAndSimulationUI === 'function') renderOperationAndSimulationUI();
    if (typeof renderManualLastVehicleInfo === 'function') renderManualLastVehicleInfo();
  }
}

function applyCastLatLng() {
  const parsed = parseLatLngText(els.castLatLngText?.value || "");
  if (!parsed) {
    alert("座標形式が正しくありません");
    return;
  }

  if (els.castLat) els.castLat.value = parsed.lat;
  if (els.castLng) els.castLng.value = parsed.lng;

  if (els.castArea) {
    els.castArea.value = getCastManagementAreaLabel(
      parsed.lat,
      parsed.lng,
      els.castAddress?.value || ""
    );
  }

  const metrics = getCastDisplayMetrics({
    latitude: parsed.lat,
    longitude: parsed.lng,
    area: els.castArea?.value || "",
    travel_minutes: els.castTravelMinutes?.value || null
  }, els.castAddress?.value || "");

  if (els.castDistanceKm) {
    els.castDistanceKm.value = formatCastDistanceDisplay(metrics.distance_km);
  }
  if (els.castTravelMinutes) {
    els.castTravelMinutes.value = formatCastTravelMinutesDisplay(metrics.travel_minutes);
  }
  updateCastDistanceHint({ distanceKm: metrics.distance_km, invalidDistance: metrics.distance_km == null });
  lastCastGeocodeKey = normalizeGeocodeAddressKey(els.castAddress?.value || "");
  setCastGeoStatus("manual", "手動反映済 | 貼り付けた座標を保存対象に反映しました");
}

function getCastManagementAreaLabel(lat, lng, address = "") {
  const addressLabel = buildCastDirectionDisplayLabelFromAddress(address, "");
  if (addressLabel) return normalizeAreaLabel(addressLabel);
  return normalizeAreaLabel(guessArea(lat, lng, address));
}

function guessCastArea() {
  const lat = toNullableNumber(els.castLat?.value);
  const lng = toNullableNumber(els.castLng?.value);
  if (els.castArea) {
    els.castArea.value = getCastManagementAreaLabel(lat, lng, els.castAddress?.value || "");
  }
}


function getFilteredCastsForSearch() {
  const nameQ = String(els.castSearchName?.value || "").trim().toLowerCase();
  const areaQ = String(els.castSearchArea?.value || "").trim().toLowerCase();
  const addressQ = String(els.castSearchAddress?.value || "").trim().toLowerCase();
  const phoneQ = String(els.castSearchPhone?.value || "").trim().toLowerCase();

  return allCastsCache.filter(cast => {
    const name = String(cast.name || "").toLowerCase();
    const areaDisplay = String(getCastDirectionDisplayLabel(cast, cast.address || "") || "").toLowerCase();
    const areaStored = String(normalizeAreaLabel(cast.area || "")).toLowerCase();
    const address = String(cast.address || "").toLowerCase();
    const phone = String(cast.phone || "").toLowerCase();

    if (nameQ && !name.includes(nameQ)) return false;
    if (areaQ && !areaDisplay.includes(areaQ) && !areaStored.includes(areaQ)) return false;
    if (addressQ && !address.includes(addressQ)) return false;
    if (phoneQ && !phone.includes(phoneQ)) return false;

    return true;
  });
}

function renderCastSearchResults() {
  if (!els.castSearchResultWrap) return;

  const rows = getFilteredCastsForSearch();
  if (els.castSearchCount) els.castSearchCount.textContent = String(rows.length);

  if (!rows.length) {
    els.castSearchResultWrap.innerHTML =
      `<div class="muted" style="padding:14px;">該当する送り先がありません</div>`;
    return;
  }

  const isReadonlyUser = isReadonlyUserRole();

  let html = `
    <table class="data-table">
      <thead>
        <tr>
          <th>氏名</th>
          <th>住所</th>
          <th>方面</th>
          <th>想定距離(km)</th>
          <th>片道予想時間</th>
          <th>電話</th>
          <th>メモ</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
  `;

  const searchRows = buildCastLiveMetricRows(rows);

  searchRows.forEach(({ cast, metrics, directionLabel }) => {
    html += `
      <tr>
        <td>
          ${
            buildCastMapUrl(cast)
              ? `<a href="${buildCastMapUrl(cast)}" target="_blank" rel="noopener noreferrer" class="cast-name-link">${escapeHtml(cast.name || "")} 📍</a>`
              : `${escapeHtml(cast.name || "")}`
          }
        </td>
        <td>${escapeHtml(cast.address || "")}</td>
        <td>${escapeHtml(directionLabel)}</td>
        <td>${formatCastDistanceDisplay(metrics.distance_km)}</td>
        <td>${formatCastTravelMinutesDisplay(metrics.travel_minutes)}</td>
        <td>${escapeHtml(cast.phone || "")}</td>
        <td>${escapeHtml(cast.memo || "")}</td>
        <td class="actions-cell">
          <button class="btn ghost cast-search-map-btn" data-id="${cast.id}">地図</button>
          <button class="btn ghost cast-search-route-btn" data-id="${escapeHtml(String(cast.id || cast.cast_id || ""))}" data-address="${escapeHtml(cast.address || "")}">ルート</button>
          ${isReadonlyUser ? '' : `<button class="btn ghost cast-search-edit-btn" data-id="${cast.id}">編集へ</button>`}
        </td>
      </tr>
    `;
  });

  html += `</tbody></table>`;
  els.castSearchResultWrap.innerHTML = html;

  els.castSearchResultWrap.querySelectorAll(".cast-search-map-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
      const url = buildCastMapUrl(cast);
      if (url) window.open(url, "_blank");
    });
  });

  els.castSearchResultWrap.querySelectorAll(".cast-search-route-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const cast = allCastsCache.find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
      if (cast) {
        openGoogleMap(cast.address || "", cast.latitude, cast.longitude);
        return;
      }
      openGoogleMap(btn.dataset.address || "");
    });
  });

  if (!isReadonlyUser) {
    els.castSearchResultWrap.querySelectorAll(".cast-search-edit-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const cast = allCastsCache.find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
        if (!cast) return;
        activateTab("castsTab");
        fillCastForm(cast);
      });
    });
  }
}

function resetCastSearchFilters() {
  if (els.castSearchName) els.castSearchName.value = "";
  if (els.castSearchArea) els.castSearchArea.value = "";
  if (els.castSearchAddress) els.castSearchAddress.value = "";
  if (els.castSearchPhone) els.castSearchPhone.value = "";
  renderCastSearchResults();
}


function setVehicleGeoStatus(kind, message) {
  if (!els.vehicleGeoStatus) return;
  els.vehicleGeoStatus.className = `geo-status ${kind}`;
  els.vehicleGeoStatus.textContent = message;
}

function syncVehicleLatLngTextFromHidden() {
  if (!els.vehicleHomeLatLngText) return;
  const lat = String(els.vehicleHomeLat?.value ?? "").trim();
  const lng = String(els.vehicleHomeLng?.value ?? "").trim();
  els.vehicleHomeLatLngText.value = lat && lng ? `${lat}, ${lng}` : "";
}

function applyVehicleLatLng() {
  const parsed = parseLatLngText(els.vehicleHomeLatLngText?.value || "");
  if (!parsed) {
    setVehicleGeoStatus("error", "座標形式が正しくありません。例：35.77265133165276, 139.92975099369397");
    alert("座標形式が正しくありません");
    return;
  }
  if (els.vehicleHomeLat) els.vehicleHomeLat.value = parsed.lat;
  if (els.vehicleHomeLng) els.vehicleHomeLng.value = parsed.lng;
  if (els.vehicleHomeLatLngText) els.vehicleHomeLatLngText.value = `${parsed.lat}, ${parsed.lng}`;
  setVehicleGeoStatus("manual", "✔ 帰宅座標を反映しました");
}

function resetVehicleForm() {
  editingVehicleId = null;
  if (els.vehiclePlateNumber) els.vehiclePlateNumber.value = "";
  if (els.vehicleArea) els.vehicleArea.value = "";
  if (els.vehicleHomeArea) els.vehicleHomeArea.value = "";
  if (els.vehicleHomeLat) els.vehicleHomeLat.value = "";
  if (els.vehicleHomeLng) els.vehicleHomeLng.value = "";
  if (els.vehicleHomeLatLngText) els.vehicleHomeLatLngText.value = "";
  setVehicleGeoStatus("idle", "座標を「緯度, 経度」の形式で貼り付けると自動反映します");
  if (els.vehicleSeatCapacity) els.vehicleSeatCapacity.value = "";
  if (els.vehicleDriverName) els.vehicleDriverName.value = "";
  if (els.vehicleLineId) els.vehicleLineId.value = "";
  if (els.vehicleStatus) els.vehicleStatus.value = "waiting";
  if (els.vehicleMemo) els.vehicleMemo.value = "";
  if (els.cancelVehicleEditBtn) els.cancelVehicleEditBtn.classList.add("hidden");
}

function fillVehicleForm(vehicle) {
  editingVehicleId = vehicle.id;
  if (els.vehiclePlateNumber) els.vehiclePlateNumber.value = vehicle.plate_number || "";
  if (els.vehicleArea) els.vehicleArea.value = normalizeAreaLabel(vehicle.vehicle_area || "");
  if (els.vehicleHomeArea) els.vehicleHomeArea.value = normalizeAreaLabel(vehicle.home_area || "");
  if (els.vehicleHomeLat) els.vehicleHomeLat.value = vehicle.home_lat ?? "";
  if (els.vehicleHomeLng) els.vehicleHomeLng.value = vehicle.home_lng ?? "";
  syncVehicleLatLngTextFromHidden();
  setVehicleGeoStatus((vehicle.home_lat != null && vehicle.home_lng != null) ? "success" : "idle", (vehicle.home_lat != null && vehicle.home_lng != null) ? "✔ 帰宅座標設定済" : "座標を「緯度, 経度」の形式で貼り付けると自動反映します");
  if (els.vehicleSeatCapacity) els.vehicleSeatCapacity.value = vehicle.seat_capacity ?? "";
  if (els.vehicleDriverName) els.vehicleDriverName.value = vehicle.driver_name || "";
  if (els.vehicleLineId) els.vehicleLineId.value = vehicle.line_id || "";
  if (els.vehicleStatus) els.vehicleStatus.value = vehicle.status || "waiting";
  if (els.vehicleMemo) els.vehicleMemo.value = vehicle.memo || "";
  if (els.cancelVehicleEditBtn) els.cancelVehicleEditBtn.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function isDuplicateVehicle(plateNumber) {
  const normalizedPlate = String(plateNumber || "").trim();
  return allVehiclesCache.find(
    v =>
      String(v.plate_number || "").trim() === normalizedPlate &&
      Number(v.id) !== Number(editingVehicleId)
  );
}

function renderVehiclesTable() {
  if (!els.vehiclesTableBody) return;

  const isReadonlyUser = isReadonlyUserRole();
  const statsMap = getVehicleStatsMapForDashboardMonth(currentDailyReportsCache);

  els.vehiclesTableBody.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.vehiclesTableBody.innerHTML = `<tr><td colspan="9" class="muted">車両がありません</td></tr>`;
    return;
  }

  getSortedVehiclesForDisplay().forEach(vehicle => {
    const stats = statsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };

    const tr = document.createElement("tr");
    const actionsHtml = isReadonlyUser
      ? '<span class="muted">閲覧専用</span>'
      : `
        <button class="btn ghost vehicle-edit-btn" data-id="${vehicle.id}">編集</button>
        <button class="btn danger vehicle-delete-btn" data-id="${vehicle.id}">削除</button>
      `;
    tr.innerHTML = `
      <tr>
      <td>${escapeHtml(vehicle.driver_name || "-")}</td>
      <td>${escapeHtml(vehicle.plate_number || "-")}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}</td>
      <td>${vehicle.seat_capacity ?? "-"}</td>
      <td>${stats.totalDistance.toFixed(1)}</td>
      <td>${stats.workedDays}</td>
      <td>${stats.avgDistance.toFixed(1)}</td>
      <td class="actions-cell">${actionsHtml}</td>
    `;
    els.vehiclesTableBody.appendChild(tr);
  });

  if (!isReadonlyUser) {
    els.vehiclesTableBody.querySelectorAll(".vehicle-edit-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const vehicle = allVehiclesCache.find(x => Number(x.id) === Number(btn.dataset.id));
        if (vehicle) fillVehicleForm(vehicle);
      });
    });

    els.vehiclesTableBody.querySelectorAll(".vehicle-delete-btn").forEach(btn => {
      btn.addEventListener("click", async () => deleteVehicle(Number(btn.dataset.id)));
    });
  }
}

function normalizeMileageExportRows(rows) {
  return (Array.isArray(rows) ? rows : []).map(row => {
    const reportDate = row?.report_date || row?.run_date || "";
    const distanceKm = Number(row?.distance_km ?? row?.reference_distance_km ?? 0);
    const normalizedVehicleId = typeof resolveVehicleLocalNumericId === "function"
      ? Number(resolveVehicleLocalNumericId(row?.vehicle_id || row?.vehicles?.id || 0) || 0)
      : Number(row?.vehicle_id || row?.vehicles?.id || 0);

    return {
      report_date: reportDate,
      vehicle_id: normalizedVehicleId,
      driver_name: row?.driver_name || row?.vehicles?.driver_name || "-",
      plate_number: row?.plate_number || row?.vehicles?.plate_number || "-",
      distance_km: distanceKm,
      worked_flag: distanceKm > 0 ? 1 : 0,
      note: row?.note || ""
    };
  });
}


function buildMileageCalendarRows(rows, startDate, endDate) {
  const start = new Date(`${startDate}T00:00:00`);
  const monthEnd = new Date(start.getFullYear(), start.getMonth() + 1, 0);

  const firstHalfDays = [];
  const secondHalfDays = [];

  for (let day = 1; day <= monthEnd.getDate(); day++) {
    const d = new Date(start.getFullYear(), start.getMonth(), day);
    const item = {
      day,
      key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`,
      label: `${day}日`
    };
    if (day <= 16) firstHalfDays.push(item);
    else secondHalfDays.push(item);
  }

  const grouped = new Map();

  (rows || []).forEach(row => {
    const vehicleId = Number(row.vehicle_id || 0);
    const driver = String(row.driver_name || "-").trim() || "-";
    const plateNumber = String(row.plate_number || "").trim();
    const key = vehicleId > 0 ? `vehicle:${vehicleId}` : `name:${driver}__${plateNumber}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        vehicle_id: vehicleId,
        plate_number: plateNumber,
        driver_name: driver,
        byDate: new Map(),
        total_distance_km: 0,
        worked_days: 0,
        avg_distance_km: 0
      });
    }

    const item = grouped.get(key);
    const dateKey = String(row.report_date || "");
    const distance = Number(row.distance_km || 0);
    item.byDate.set(dateKey, Number(((item.byDate.get(dateKey) || 0) + distance).toFixed(1)));
    item.total_distance_km += distance;
  });

  const vehicleOrderMap = new Map(
    (typeof getSortedVehiclesForDisplay === "function" ? getSortedVehiclesForDisplay() : (allVehiclesCache || []))
      .map((vehicle, index) => [Number(vehicle.id || 0), index])
  );

  const drivers = [...grouped.values()].sort((a, b) => {
    const orderA = vehicleOrderMap.has(Number(a.vehicle_id || 0)) ? vehicleOrderMap.get(Number(a.vehicle_id || 0)) : Number.MAX_SAFE_INTEGER;
    const orderB = vehicleOrderMap.has(Number(b.vehicle_id || 0)) ? vehicleOrderMap.get(Number(b.vehicle_id || 0)) : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    return String(a.driver_name || "").localeCompare(String(b.driver_name || ""), "ja", { numeric: true, sensitivity: "base" });
  });

  drivers.forEach(driver => {
    driver.worked_days = [...driver.byDate.values()].filter(v => Number(v || 0) > 0).length;
    driver.total_distance_km = Number(driver.total_distance_km.toFixed(1));
    driver.avg_distance_km = driver.worked_days ? Number((driver.total_distance_km / driver.worked_days).toFixed(1)) : 0;
  });

  return {
    firstHalfDays,
    secondHalfDays,
    drivers,
    monthLabel: `${start.getFullYear()}年${start.getMonth() + 1}月`
  };
}

function renderMileageMatrixSection(title, days, drivers, options = {}) {
  const showOverallColumns = Boolean(options.showOverallColumns);
  const headerDays = days.map(day => `<th>${escapeHtml(day.label)}</th>`).join("");
  const extraHeader = showOverallColumns ? `<th>全合計距離</th><th>全出勤日数</th><th>全平均距離</th>` : "";

  const bodyRows = drivers.map(driver => {
    const sectionValues = days.map(day => Number(driver.byDate.get(day.key) || 0));
    const sectionDistance = Number(sectionValues.reduce((sum, value) => sum + value, 0).toFixed(1));
    const sectionWorkedDays = sectionValues.filter(value => value > 0).length;
    const sectionAvgDistance = sectionWorkedDays ? Number((sectionDistance / sectionWorkedDays).toFixed(1)) : 0;
    const cells = sectionValues.map(value => `<td class="mileage-day-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");
    const extraOverallCells = showOverallColumns ? `
        <td>${Number(driver.total_distance_km || 0).toFixed(1)}km</td>
        <td>${Number(driver.worked_days || 0)}日</td>
        <td>${Number(driver.avg_distance_km || 0).toFixed(1)}km</td>` : "";

    return `<tr>
      <td class="mileage-driver-cell">${escapeHtml(driver.driver_name || "-")}</td>
      ${cells}
      <td>${sectionDistance.toFixed(1)}km</td>
      <td>${sectionWorkedDays}日</td>
      <td>${sectionAvgDistance.toFixed(1)}km</td>
      ${extraOverallCells}
    </tr>`;
  }).join("");

  const dailyTotals = days.map(day => Number(drivers.reduce((sum, driver) => sum + Number(driver.byDate.get(day.key) || 0), 0).toFixed(1)));
  const dailyWorkedCounts = days.map(day => drivers.reduce((count, driver) => count + (Number(driver.byDate.get(day.key) || 0) > 0 ? 1 : 0), 0));
  const dailyAverages = dailyTotals.map((total, index) => dailyWorkedCounts[index] > 0 ? Number((total / dailyWorkedCounts[index]).toFixed(1)) : 0);

  const sectionTotalDistance = Number(dailyTotals.reduce((sum, value) => sum + Number(value || 0), 0).toFixed(1));
  const sectionWorkedDays = dailyWorkedCounts.reduce((sum, value) => sum + Number(value || 0), 0);
  const sectionAverageDistance = sectionWorkedDays ? Number((sectionTotalDistance / sectionWorkedDays).toFixed(1)) : 0;

  const overallTotalDistance = Number(drivers.reduce((sum, driver) => sum + Number(driver.total_distance_km || 0), 0).toFixed(1));
  const overallWorkedDays = drivers.reduce((sum, driver) => sum + Number(driver.worked_days || 0), 0);
  const overallAvgDistance = overallWorkedDays ? Number((overallTotalDistance / overallWorkedDays).toFixed(1)) : 0;

  const footerTotalCells = dailyTotals.map(value => `<td class="mileage-day-cell mileage-summary-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");
  const footerAverageCells = dailyAverages.map(value => `<td class="mileage-day-cell mileage-summary-cell">${value > 0 ? `${value.toFixed(1)}km` : "-"}</td>`).join("");

  const footerOverallCells = showOverallColumns ? `
      <td class="mileage-summary-cell">${overallTotalDistance.toFixed(1)}km</td>
      <td class="mileage-summary-cell">${overallWorkedDays}日</td>
      <td class="mileage-summary-cell">${overallAvgDistance.toFixed(1)}km</td>` : "";

  const footerAverageOverallCells = showOverallColumns ? `
      <td class="mileage-summary-cell">${drivers.length ? (overallTotalDistance / drivers.length).toFixed(1) : "0.0"}km</td>
      <td class="mileage-summary-cell">${drivers.length ? (overallWorkedDays / drivers.length).toFixed(1) : "0.0"}日</td>
      <td class="mileage-summary-cell">${overallAvgDistance.toFixed(1)}km</td>` : "";

  const colspan = days.length + 4 + (showOverallColumns ? 3 : 0);

  return `<div class="mileage-matrix-section">
    <div class="grouped-hour-title">${escapeHtml(title)}</div>
    <div class="table-wrap mileage-matrix-wrap">
      <table class="matrix-table mileage-matrix-table">
        <thead>
          <tr>
            <th>ドライバー</th>
            ${headerDays}
            <th>合計距離</th>
            <th>合計出勤数</th>
            <th>1日平均距離</th>
            ${extraHeader}
          </tr>
        </thead>
        <tbody>
          ${bodyRows || `<tr><td colspan="${colspan}" class="muted">データがありません</td></tr>`}
          <tr class="mileage-summary-row">
            <td class="mileage-driver-cell">日別合計</td>
            ${footerTotalCells}
            <td class="mileage-summary-cell">${sectionTotalDistance.toFixed(1)}km</td>
            <td class="mileage-summary-cell">${sectionWorkedDays}日</td>
            <td class="mileage-summary-cell">${sectionAverageDistance.toFixed(1)}km</td>
            ${footerOverallCells}
          </tr>
          <tr class="mileage-summary-row">
            <td class="mileage-driver-cell">日別平均</td>
            ${footerAverageCells}
            <td class="mileage-summary-cell">${drivers.length ? (sectionTotalDistance / drivers.length).toFixed(1) : "0.0"}km</td>
            <td class="mileage-summary-cell">-</td>
            <td class="mileage-summary-cell">${sectionAverageDistance.toFixed(1)}km</td>
            ${footerAverageOverallCells}
          </tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderMileageReportTable(rows) {
  if (!els.mileageReportTableWrap) return;

  if (!rows.length) {
    els.mileageReportTableWrap.innerHTML = `<div class="muted" style="padding:14px;">対象期間の走行実績はありません</div>`;
    return;
  }

  const startDate = els.mileageReportStartDate?.value || todayStr();
  const endDate = els.mileageReportEndDate?.value || todayStr();
  const calendar = buildMileageCalendarRows(rows, startDate, endDate);

  let html = `<div class="grouped-plan-list mileage-report-grid">`;
  html += `<div class="grouped-hour-title">${escapeHtml(calendar.monthLabel)} / ${escapeHtml(startDate.replaceAll("-", "/"))} ～ ${escapeHtml(endDate.replaceAll("-", "/"))}</div>`;
  html += renderMileageMatrixSection("1日～16日", calendar.firstHalfDays, calendar.drivers);
  html += renderMileageMatrixSection("17日～月末", calendar.secondHalfDays, calendar.drivers, { showOverallColumns: true });
  html += `</div>`;

  els.mileageReportTableWrap.innerHTML = html;
}


function syncMileageReportRange(dateStr = todayStr(), force = false) {
  const targetDate = dateStr || todayStr();
  const start = getMonthStartStr(targetDate);
  const end = targetDate;

  if (force) {
    forceResetMileageReportInputs(targetDate);
    return;
  }

  if (els.mileageReportStartDate && !els.mileageReportStartDate.value) {
    applyMileageDateValue(els.mileageReportStartDate, start);
  }
  if (els.mileageReportEndDate && !els.mileageReportEndDate.value) {
    applyMileageDateValue(els.mileageReportEndDate, end);
  }
}

function initializeMileageReportDefaultDates() {
  const today = todayStr();
  forceResetMileageReportInputs(today);
}

function applyMileageDateValue(input, value) {
  if (!input) return;
  input.value = value;
  input.defaultValue = value;
  try { input.setAttribute("value", value); } catch (e) {}
}

function replaceMileageDateInput(inputId) {
  const current = document.getElementById(inputId);
  if (!current || !current.parentNode) return current;
  const clone = current.cloneNode(true);
  clone.value = "";
  clone.defaultValue = "";
  try { clone.removeAttribute("value"); } catch (e) {}
  current.parentNode.replaceChild(clone, current);
  return clone;
}

function forceResetMileageReportInputs(dateStr = todayStr()) {
  const targetDate = dateStr || todayStr();
  const start = getMonthStartStr(targetDate);
  const end = targetDate;

  const startInput = replaceMileageDateInput("mileageReportStartDate");
  const endInput = replaceMileageDateInput("mileageReportEndDate");

  els.mileageReportStartDate = startInput || document.getElementById("mileageReportStartDate");
  els.mileageReportEndDate = endInput || document.getElementById("mileageReportEndDate");

  applyMileageDateValue(els.mileageReportStartDate, start);
  applyMileageDateValue(els.mileageReportEndDate, end);
}

async function previewDriverMileageReport() {
  const startDate = els.mileageReportStartDate?.value;
  const endDate = els.mileageReportEndDate?.value;

  if (!startDate || !endDate) {
    alert("開始日と終了日を選択してください");
    return;
  }

  if (startDate > endDate) {
    alert("開始日は終了日以前にしてください");
    return;
  }

  const rawRows = await fetchDriverMileageRows(startDate, endDate);
  currentMileageExportRows = normalizeMileageExportRows(rawRows);
  renderMileageReportTable(currentMileageExportRows);
}
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

  const targetDate = todayStr();
  const startDate = els.mileageReportStartDate?.value || getMonthStartStr(targetDate);
  const endDate = els.mileageReportEndDate?.value || targetDate;

  const aoa = buildMileageMatrixCsvRows(currentMileageExportRows, startDate, endDate);
  const csv = aoa.map(row => row.map(value => csvEscape(value ?? "")).join(",")).join("\n");
  downloadTextFile(`driver_mileage_matrix_${startDate}_${endDate}.csv`, csv, "text/csv;charset=utf-8");
}

async function exportDriverMileageReportXlsx() {
  if (!ensureCsvFeatureAccess("CSVエクスポート")) return;
  return exportDriverMileageReportCsv();
}


function exportVehiclesCsv() {
  if (!ensureCsvFeatureAccess("車両CSV")) return;
  const headers = [
    "plate_number",
    "vehicle_area",
    "home_area",
    "home_lat",
    "home_lng",
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
    vehicle.home_lat ?? "",
    vehicle.home_lng ?? "",
    vehicle.seat_capacity ?? "",
    vehicle.driver_name || "",
    vehicle.line_id || "",
    vehicle.status || "waiting",
    vehicle.memo || ""
  ]);

  const csv = [headers.join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  downloadTextFile(`vehicles_${todayStr()}.csv`, csv, "text/csv;charset=utf-8");
}


function normalizePlanImportMode(input) {
  const value = String(input || "").trim();
  if (["1", "add", "append", "追加"].includes(value)) return "append";
  if (["2", "replace", "置換", "上書き"].includes(value)) return "replace";
  if (["3", "skip", "重複スキップ", "skip-duplicates"].includes(value)) return "skip";
  return "";
}

function normalizeCastMatchText(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}

function findCastForPlanImport(row, casts = []) {
  const castIdValue = normalizeDispatchEntityId(row.cast_id || row.person_id || "");
  const castName = normalizeCastMatchText(row.cast_name || row.person_name);
  const castAddress = normalizeCastMatchText(row.cast_address || row.destination_address);
  const castPhone = normalizeCastMatchText(row.cast_phone || row.phone);

  if (castIdValue) {
    const byId = casts.find(x => sameDispatchEntityId(x.id || x.cast_id, castIdValue));
    if (byId) return byId;
  }

  if (castName && castAddress) {
    const byNameAddress = casts.find(x => (
      normalizeCastMatchText(x.name) === castName &&
      normalizeCastMatchText(x.address) === castAddress
    ));
    if (byNameAddress) return byNameAddress;
  }

  if (castName && castPhone) {
    const byNamePhone = casts.find(x => (
      normalizeCastMatchText(x.name) === castName &&
      normalizeCastMatchText(x.phone) === castPhone
    ));
    if (byNamePhone) return byNamePhone;
  }

  if (castAddress) {
    const addressMatches = casts.filter(x => normalizeCastMatchText(x.address) === castAddress);
    if (addressMatches.length === 1) return addressMatches[0];
  }

  if (castName) {
    const nameMatches = casts.filter(x => normalizeCastMatchText(x.name) === castName);
    if (nameMatches.length === 1) return nameMatches[0];
  }

  return null;
}

function getPlanDuplicateKey(row) {
  return [
    String(row.plan_date || row.dispatch_date || "").trim(),
    Number(row.plan_hour ?? row.dispatch_hour ?? 0),
    normalizeDispatchEntityId(row.cast_id || row.person_id || ""),
    String(row.destination_address || "").trim(),
    normalizeAreaLabel(row.planned_area || row.destination_area || ""),
    String(row.note || "").trim()
  ].join("|");
}

function exportPlansCsv() {
  if (!ensureCsvFeatureAccess("予定CSVエクスポート")) return;
  const planDate = els.planDate?.value || todayStr();
  const rows = [...currentPlansCache]
    .sort((a, b) => Number(a.plan_hour || a.dispatch_hour || 0) - Number(b.plan_hour || b.dispatch_hour || 0) || String(a.id || "").localeCompare(String(b.id || ""), "ja"))
    .map(plan => ({
      plan_date: plan.plan_date || plan.dispatch_date || planDate,
      plan_hour: Number(plan.plan_hour || plan.dispatch_hour || 0),
      cast_id: normalizeDispatchEntityId(plan.cast_id || plan.person_id || ""),
      cast_name: plan.casts?.name || plan.person_name || "",
      cast_address: plan.casts?.address || "",
      cast_phone: plan.casts?.phone || "",
      destination_address: plan.destination_address || plan.casts?.address || "",
      planned_area: normalizeAreaLabel(plan.planned_area || plan.destination_area || plan.casts?.area || ""),
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
    "cast_address",
    "cast_phone",
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

async function triggerImportPlansCsv() {
  if (!ensureCsvFeatureAccess("予定CSVインポート")) return;
  els.plansCsvFileInput?.click();
}

async function importPlansCsvFile() {
  if (!ensureCsvFeatureAccess("予定CSVインポート")) {
    if (els.plansCsvFileInput) els.plansCsvFileInput.value = "";
    return;
  }
  const file = els.plansCsvFileInput?.files?.[0];
  if (!file) return;

  const selectedDate = els.planDate?.value || todayStr();
  const modeInput = window.prompt(
    "予定CSVの取込方法を選んでください\n1: 追加\n2: 同日データを置換\n3: 重複をスキップ",
    "3"
  );
  const mode = normalizePlanImportMode(modeInput);
  if (!mode) {
    alert("取込を中止しました");
    els.plansCsvFileInput.value = "";
    return;
  }

  try {
    const text = await readCsvFileAsText(file);
    let rows = parseCsv(text);
    rows = normalizeCsvRows(rows);

    if (!rows.length) {
      alert("CSVにデータがありません");
      els.plansCsvFileInput.value = "";
      return;
    }

    const plansTable = getDispatchUnifiedTableName();
    let existingQuery = supabaseClient
      .from(plansTable)
      .select("id, plan_date, dispatch_date, plan_hour, dispatch_hour, cast_id, person_id, destination_address, planned_area, destination_area, note");

    existingQuery = existingQuery
      .eq("dispatch_kind", "plan")
      .eq("dispatch_date", selectedDate)
      .order("dispatch_hour", { ascending: true });

    const { data: existingRows, error: existingError } = await existingQuery;

    if (existingError) {
      alert(existingError.message);
      els.plansCsvFileInput.value = "";
      return;
    }

    const existingList = Array.isArray(existingRows)
      ? existingRows.map(mapUnifiedDispatchRowToPlan)
      : [];
    const existingKeys = new Set(existingList.map(getPlanDuplicateKey));

    if (mode === "replace" && existingList.length) {
      const deleteQuery = supabaseClient
        .from(plansTable)
        .delete()
        .eq("dispatch_kind", "plan")
        .eq("dispatch_date", selectedDate);
      const { error: deleteError } = await deleteQuery;
      if (deleteError) {
        alert(deleteError.message);
        els.plansCsvFileInput.value = "";
        return;
      }
      existingKeys.clear();
    }

    const inserts = [];
    const skipped = [];
    const missingCasts = [];

    for (const row of rows) {
      const cast = findCastForPlanImport(row, allCastsCache);
      if (!cast) {
        missingCasts.push(String(row.cast_name || row.cast_address || row.cast_id || "不明"));
        continue;
      }

      const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
      const hour = Number(row.plan_hour || row.dispatch_hour || 0);
      const address = String(row.destination_address || cast.address || "").trim();
      const area = normalizeAreaLabel(String(row.planned_area || row.destination_area || cast.area || "無し"));
      const note = String(row.note || "").trim();
      const status = String(row.status || "planned").trim() || "planned";
      const vehicleGroup = String(row.vehicle_group || "").trim();

      let payload;
      if (isSingleDispatchTableMode()) {
        const snap = await buildDispatchMetricSnapshot(cast, address, area);
        const csvDistanceKm = toNullableNumber(row.distance_km);
        payload = {
          dispatch_kind: "plan",
          dispatch_date: selectedDate,
          plan_date: selectedDate,
          dispatch_hour: hour,
          plan_hour: hour,
          cast_id: castId,
          person_id: castId,
          person_name: cast.name || "",
          destination_address: snap.destination_address,
          planned_area: area,
          destination_area: area,
          straight_km: snap.straight_km,
          distance_km: csvDistanceKm ?? snap.distance_km,
          travel_minutes: snap.travel_minutes,
          origin_slot_no: snap.origin_slot_no,
          origin_label: snap.origin_label,
          origin_address: snap.origin_address,
          origin_lat: snap.origin_lat,
          origin_lng: snap.origin_lng,
          destination_lat: snap.destination_lat,
          destination_lng: snap.destination_lng,
          note,
          status
        };
      } else {
        payload = {
          plan_date: selectedDate,
          plan_hour: hour,
          cast_id: castId,
          destination_address: address,
          planned_area: area,
          distance_km: toNullableNumber(row.distance_km),
          note,
          status,
          created_by: currentUser?.id || null
        };
      }

      const key = getPlanDuplicateKey(payload);
      if (mode === "skip" && existingKeys.has(key)) {
        skipped.push(`${getHourLabel(hour)} / ${cast.name}`);
        continue;
      }
      existingKeys.add(key);
      inserts.push(payload);
    }

    if (!inserts.length) {
      let msg = "取り込める予定がありませんでした。";
      if (missingCasts.length) msg += `\n未登録送り先: ${[...new Set(missingCasts)].join(", ")}`;
      if (skipped.length) msg += `\n重複スキップ: ${skipped.length}件`;
      alert(msg);
      els.plansCsvFileInput.value = "";
      await loadPlansByDate(selectedDate);
      return;
    }

    for (const payload of inserts) {
      const safePayload = isSingleDispatchTableMode()
        ? {
            dispatch_kind: "plan",
            dispatch_date: payload.dispatch_date,
            plan_date: payload.plan_date,
            dispatch_hour: Number(payload.dispatch_hour || 0),
            plan_hour: Number(payload.plan_hour || 0),
            cast_id: payload.cast_id,
            person_id: payload.person_id,
            person_name: payload.person_name || "",
            destination_address: payload.destination_address || "",
            planned_area: normalizeAreaLabel(payload.planned_area || payload.destination_area || "無し"),
            destination_area: normalizeAreaLabel(payload.destination_area || payload.planned_area || "無し"),
            straight_km: toNullableNumber(payload.straight_km),
            distance_km: toNullableNumber(payload.distance_km),
            travel_minutes: Math.max(0, Number(payload.travel_minutes || 0) || 0),
            origin_slot_no: payload.origin_slot_no ?? null,
            origin_label: payload.origin_label || null,
            origin_address: payload.origin_address || null,
            origin_lat: toNullableNumber(payload.origin_lat),
            origin_lng: toNullableNumber(payload.origin_lng),
            destination_lat: toNullableNumber(payload.destination_lat),
            destination_lng: toNullableNumber(payload.destination_lng),
            note: payload.note || "",
            status: String(payload.status || "planned").trim() || "planned"
          }
        : payload;

      const { error: insertError } = await supabaseClient
        .from(plansTable)
        .insert(safePayload);

      if (insertError) {
        console.error("plans csv insert failed:", insertError, safePayload);
        alert(insertError.message);
        els.plansCsvFileInput.value = "";
        return;
      }
    }

    let summary = `${inserts.length}件の予定をCSV取込しました`;
    if (mode === "replace") summary += "（同日置換）";
    if (skipped.length) summary += ` / 重複スキップ ${skipped.length}件`;
    if (missingCasts.length) summary += ` / 未登録送り先 ${[...new Set(missingCasts)].length}件`;

    await addHistory(null, null, "import_plans_csv", summary);
    alert(summary + `\n取込日付: ${selectedDate}`);
    els.plansCsvFileInput.value = "";
    await loadPlansByDate(selectedDate);
  } catch (error) {
    console.error("importPlansCsvFile error:", error);
    alert("予定CSV取込に失敗しました");
    els.plansCsvFileInput.value = "";
  }
}

function clearPlanCastDerivedFields() {
  if (els.planAddress) els.planAddress.value = "";
  if (els.planArea) els.planArea.value = "";
  if (els.planDistanceKm) els.planDistanceKm.value = "";
}

function clearActualCastDerivedFields() {
  if (els.actualAddress) els.actualAddress.value = "";
  if (els.actualArea) els.actualArea.value = "";
  if (els.actualDistanceKm) els.actualDistanceKm.value = "";
}

async function resolveFormDisplayDistanceKm(cast, address = "", area = "") {
  if (!cast) return null;

  try {
    if (typeof buildDispatchMetricSnapshot === "function") {
      const snap = await buildDispatchMetricSnapshot(cast, address, area);
      const km = toNullableNumber(snap?.distance_km);
      if (km !== null) return km;
    }
  } catch (error) {
    console.warn("resolveFormDisplayDistanceKm snapshot fallback:", error);
  }

  const fallback = await resolveDistanceKmForCastRecord(cast, address);
  const km = toNullableNumber(fallback);
  if (km !== null) return km;

  const stored = toNullableNumber(cast?.distance_km);
  if (stored !== null) {
    if (stored >= 1000) return Number((stored / 1000).toFixed(1));
    return stored;
  }

  return null;
}

async function syncPlanFieldsFromCastInput(forceFill = false) {
  const cast = resolveLinkedCastFromInput(els.planCastSelect, getPlanSelectableCasts);
  if (!cast) {
    clearPlanCastDerivedFields();
    return null;
  }

  const address = cast.address || "";
  const area = normalizeAreaLabel(
    cast.area ||
      guessArea(
        toNullableNumber(cast.latitude),
        toNullableNumber(cast.longitude),
        address
      )
  );
  const distance = await resolveFormDisplayDistanceKm(cast, address, area);

  if (els.planAddress) els.planAddress.value = address;
  if (els.planArea) els.planArea.value = area;
  if (els.planDistanceKm) els.planDistanceKm.value = distance ?? "";

  return cast;
}

async function syncActualFieldsFromCastInput(forceFill = false) {
  const cast = resolveLinkedCastFromInput(els.castSelect, getActualSelectableCasts);
  if (!cast) {
    clearActualCastDerivedFields();
    return null;
  }

  const address = cast.address || "";
  const area = normalizeAreaLabel(
    cast.area ||
      guessArea(
        toNullableNumber(cast.latitude),
        toNullableNumber(cast.longitude),
        address
      )
  );
  const distance = await resolveFormDisplayDistanceKm(cast, address, area);

  if (els.actualAddress) els.actualAddress.value = address;
  if (els.actualArea) els.actualArea.value = area;
  if (els.actualDistanceKm) els.actualDistanceKm.value = distance ?? "";

  return cast;
}

function resetPlanForm() {
  editingPlanId = null;
  if (els.planCastSelect) {
    els.planCastSelect.value = "";
    clearLinkedCastSelection(els.planCastSelect);
  }
  if (els.planHour) els.planHour.value = "0";
  clearPlanCastDerivedFields();
  if (els.planNote) els.planNote.value = "";
}

function fillPlanForm(plan) {
  editingPlanId = plan.id;
  if (els.planCastSelect) {
    els.planCastSelect.value = plan.casts?.name || plan.person_name || "";
    if (plan.cast_id) {
      els.planCastSelect.dataset.selectedCastId = normalizeDispatchEntityId(plan.cast_id);
    } else {
      clearLinkedCastSelection(els.planCastSelect);
    }
  }
  if (els.planHour) els.planHour.value = String(plan.plan_hour ?? 0);
  if (els.planDistanceKm) els.planDistanceKm.value = plan.distance_km ?? "";
  if (els.planAddress) els.planAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.planArea) els.planArea.value = normalizeAreaLabel(plan.planned_area || "");
  if (els.planNote) els.planNote.value = plan.note || "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function fillPlanFormFromSelectedCast() {
  const cast = resolveLinkedCastFromInput(els.planCastSelect, getPlanSelectableCasts);
  if (!cast) return;

  const address = cast.address || "";
  const area = normalizeAreaLabel(
    cast.area ||
      guessArea(
        toNullableNumber(cast.latitude),
        toNullableNumber(cast.longitude),
        address
      )
  );

  if (els.planAddress && !els.planAddress.value.trim()) {
    els.planAddress.value = address;
  }

  if (els.planArea && !els.planArea.value.trim()) {
    els.planArea.value = area;
  }

  if (els.planDistanceKm && !els.planDistanceKm.value.trim()) {
    const distance = await resolveFormDisplayDistanceKm(cast, address, area);
    els.planDistanceKm.value = distance ?? "";
  }
}


function syncActualRendererDeps() {
  if (typeof window.setActualRendererDeps !== "function") return;

  window.setActualRendererDeps({
    getEls: () => els,
    getActuals: () => currentActualsCache,
    getActions: () => ({
      fillActualForm,
      openGoogleMap,
      deleteActual,
      updateActualStatus
    }),
    getHelpers: () => ({
      normalizeStatus,
      getHourLabel,
      getGroupedAreasByDisplay,
      getGroupedAreaHeaderHtml,
      buildMapLinkHtml,
      escapeHtml,
      normalizeAreaLabel,
      getStatusText,
      getAreaDisplayGroup,
      getMatrixLegendHtml: typeof window.getMatrixLegendHtml === "function" ? window.getMatrixLegendHtml : (() => ""),
      buildMatrixNameLine: typeof window.buildMatrixNameLine === "function" ? window.buildMatrixNameLine : (() => ""),
      getCurrentOriginRuntime,
      computeBearingDeg,
      directionLabelFromDeg,
      extractActiveHours,
      buildDirectionClusters,
      buildTimeDirectionMatrix,
      normalizeDirectionUiSourceItem
    })
  });
}

function syncScheduleRendererDeps() {
  if (typeof window.setScheduleRendererDeps === "function") {
    window.setScheduleRendererDeps({
      els,
      plans: currentPlansCache,
      actuals: currentActualsCache,
      actions: {
        fillPlanForm,
        openGoogleMap,
        deletePlan
      },
      helpers: {
        normalizeStatus,
        buildMapLinkHtml,
        escapeHtml,
        normalizeAreaLabel,
        getStatusText,
        getHourLabel,
        getAreaDisplayGroup,
        getGroupedAreasByDisplay,
        getGroupedAreaHeaderHtml,
        todayStr,
        getCurrentOriginRuntime,
        computeBearingDeg,
        directionLabelFromDeg,
        extractActiveHours,
        buildDirectionClusters,
        buildTimeDirectionMatrix,
        normalizeDirectionUiSourceItem
      }
    });
  }

  syncActualRendererDeps();
}


function isSkippableDropoffMissingTableError(error) {
  if (typeof isMissingTableError === "function" && isMissingTableError(error)) return true;
  const code = String(error?.code || "");
  const message = String(error?.message || "");
  return code === "PGRST205" || /Could not find the table/i.test(message) || /schema cache/i.test(message);
}

function isSingleDispatchTableMode() {
  return true;
}

function getDispatchUnifiedTableName() {
  return getTableName("dispatches");
}

function getOriginSnapshotForDispatch() {
  const runtime = typeof getCurrentOriginRuntime === "function"
    ? getCurrentOriginRuntime()
    : { slot_no: null, name: ORIGIN_LABEL || "起点", address: "", lat: null, lng: null };

  return {
    origin_slot_no: runtime?.slot_no ?? null,
    origin_label: runtime?.name || ORIGIN_LABEL || "起点",
    origin_address: runtime?.address || "",
    origin_lat: toNullableNumber(runtime?.lat),
    origin_lng: toNullableNumber(runtime?.lng)
  };
}

async function buildDispatchMetricSnapshot(cast, addressOverride = "", areaOverride = "") {
  const address = String(addressOverride || cast?.address || "").trim();
  const destLat = toNullableNumber(cast?.latitude ?? cast?.lat);
  const destLng = toNullableNumber(cast?.longitude ?? cast?.lng);
  const origin = getOriginSnapshotForDispatch();
  const originRuntime = {
    name: String(origin.origin_label || origin.name || ORIGIN_LABEL || "起点").trim() || "起点",
    lat: toNullableNumber(origin.origin_lat),
    lng: toNullableNumber(origin.origin_lng)
  };
  const area = normalizeAreaLabel(areaOverride || cast?.area || guessArea(destLat, destLng, address) || "無し");

  const metrics = (typeof getCastOriginMetrics === "function")
    ? getCastOriginMetrics({ latitude: destLat, longitude: destLng, area }, address, originRuntime)
    : null;

  const straightKm = toNullableNumber(metrics?.straight_km);
  const distanceKm = toNullableNumber(metrics?.distance_km);
  const travelMinutes = Number.isFinite(Number(metrics?.travel_minutes))
    ? Number(metrics.travel_minutes)
    : (distanceKm != null ? estimateTravelMinutesByAreaSpeed(Number(distanceKm), area) : 0);

  return {
    ...origin,
    destination_address: address,
    destination_lat: destLat,
    destination_lng: destLng,
    straight_km: straightKm,
    distance_km: distanceKm,
    travel_minutes: travelMinutes,
    normalized_area: area
  };
}

function buildDispatchMetricCastLike(row = {}) {
  const linkedCast = row?.casts || null;
  const latitude = toNullableNumber(row?.destination_lat ?? linkedCast?.latitude ?? linkedCast?.lat);
  const longitude = toNullableNumber(row?.destination_lng ?? linkedCast?.longitude ?? linkedCast?.lng);
  const address = String(row?.destination_address || linkedCast?.address || "").trim();
  const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || linkedCast?.area || guessArea(latitude, longitude, address) || "無し");

  return {
    castLike: {
      latitude,
      longitude,
      area,
      address
    },
    address,
    area,
    linkedCast,
    latitude,
    longitude
  };
}

async function applyCurrentOriginMetricsToDispatchRow(row = {}) {
  const { castLike, address, area, linkedCast, latitude, longitude } = buildDispatchMetricCastLike(row);
  const hasTargetCoords = typeof isValidLatLng === 'function' && isValidLatLng(latitude, longitude);
  const snap = hasTargetCoords
    ? await buildDispatchMetricSnapshot(castLike, address, area)
    : getOriginSnapshotForDispatch();

  const nextArea = normalizeAreaLabel(area || row?.destination_area || row?.planned_area || linkedCast?.area || "無し");
  const nextRow = {
    ...row,
    pickup_label: snap?.origin_label || row?.pickup_label || row?.origin_label || ORIGIN_LABEL,
    origin_slot_no: snap?.origin_slot_no ?? row?.origin_slot_no ?? null,
    origin_label: snap?.origin_label || row?.origin_label || ORIGIN_LABEL,
    origin_address: snap?.origin_address || row?.origin_address || "",
    origin_lat: snap?.origin_lat ?? row?.origin_lat ?? null,
    origin_lng: snap?.origin_lng ?? row?.origin_lng ?? null,
    destination_address: address || row?.destination_address || linkedCast?.address || "",
    destination_area: nextArea,
    planned_area: row?.dispatch_kind === 'plan'
      ? nextArea
      : (row?.planned_area || row?.destination_area || nextArea),
    destination_lat: latitude ?? toNullableNumber(row?.destination_lat ?? linkedCast?.latitude ?? linkedCast?.lat),
    destination_lng: longitude ?? toNullableNumber(row?.destination_lng ?? linkedCast?.longitude ?? linkedCast?.lng)
  };

  if (snap && hasTargetCoords) {
    nextRow.straight_km = toNullableNumber(snap.straight_km);
    nextRow.distance_km = toNullableNumber(snap.distance_km);
    nextRow.travel_minutes = Number.isFinite(Number(snap.travel_minutes)) ? Number(snap.travel_minutes) : null;
  }

  if (linkedCast) {
    nextRow.casts = {
      ...linkedCast,
      address: address || linkedCast?.address || "",
      area: nextArea,
      latitude: latitude ?? toNullableNumber(linkedCast?.latitude ?? linkedCast?.lat),
      longitude: longitude ?? toNullableNumber(linkedCast?.longitude ?? linkedCast?.lng),
      lat: latitude ?? toNullableNumber(linkedCast?.lat ?? linkedCast?.latitude),
      lng: longitude ?? toNullableNumber(linkedCast?.lng ?? linkedCast?.longitude),
      distance_km: hasTargetCoords && snap ? toNullableNumber(snap.distance_km) : linkedCast?.distance_km,
      travel_minutes: hasTargetCoords && snap && Number.isFinite(Number(snap.travel_minutes)) ? Number(snap.travel_minutes) : linkedCast?.travel_minutes
    };
  }

  return nextRow;
}

async function applyCurrentOriginMetricsToDispatchRows(rows = []) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return [];
  return await Promise.all(list.map(row => applyCurrentOriginMetricsToDispatchRow(row)));
}

function mapUnifiedDispatchRowToPlan(row) {
  return {
    ...row,
    cast_id: row.cast_id ?? row.person_id ?? null,
    plan_date: row.plan_date || row.dispatch_date || null,
    plan_hour: row.plan_hour ?? row.dispatch_hour ?? 0,
    planned_area: row.planned_area || row.destination_area || "",
    casts: row.casts || null
  };
}

function mapUnifiedDispatchRowToActual(row) {
  return {
    ...row,
    cast_id: row.cast_id ?? row.person_id ?? null,
    plan_date: row.plan_date || row.dispatch_date || null,
    actual_hour: row.actual_hour ?? row.dispatch_hour ?? 0,
    pickup_label: row.pickup_label || row.origin_label || ORIGIN_LABEL,
    destination_area: row.destination_area || row.planned_area || "",
    casts: row.casts || null
  };
}

function enrichUnifiedDispatchRowWithCast(row = {}) {
  const castId = normalizeDispatchEntityId(row.cast_id ?? row.person_id);
  const linkedCast = castId
    ? (Array.isArray(allCastsCache) ? allCastsCache.find(c => sameDispatchEntityId(c.id || c.cast_id, castId)) : null)
    : null;
  const casts = linkedCast || row.casts || null;
  return {
    ...row,
    cast_id: castId || null,
    person_id: castId || row.person_id || null,
    person_name: row.person_name || casts?.name || '',
    destination_address: row.destination_address || casts?.address || '',
    destination_lat: row.destination_lat ?? casts?.latitude ?? casts?.lat ?? null,
    destination_lng: row.destination_lng ?? casts?.longitude ?? casts?.lng ?? null,
    planned_area: row.planned_area || row.destination_area || casts?.area || '',
    destination_area: row.destination_area || row.planned_area || casts?.area || '',
    casts
  };
}

async function loadPlansByDateSingle(dateStr) {
  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();
  const { data, error } = await supabaseClient
    .from(getDispatchUnifiedTableName())
    .select(remapRelationSelect(`
      *,
      casts (
        id,
        team_id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude,
        lat,
        lng
      )
    `))
    .eq("dispatch_kind", "plan")
    .eq("dispatch_date", dateStr)
    .order("dispatch_hour", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  const scopedRows = filterDispatchRowsByWorkspace(data, workspaceTeamId);
  const hydratedRows = scopedRows.map(mapUnifiedDispatchRowToPlan).map(enrichUnifiedDispatchRowWithCast);
  currentPlansCache = await applyCurrentOriginMetricsToDispatchRows(hydratedRows);
  syncScheduleRendererDeps();
  renderPlanGroupedTable();
  renderPlansTimeAreaMatrix();
  renderPlanSelect();
  renderPlanCastSelect();
  renderHomeSummary();
  renderOperationAndSimulationUI();
}

async function loadActualsByDateSingle(dateStr) {
  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();
  const { data, error } = await supabaseClient
    .from(getDispatchUnifiedTableName())
    .select(remapRelationSelect(`
      *,
      casts (
        id,
        team_id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude,
        lat,
        lng
      )
    `))
    .eq("dispatch_kind", "actual")
    .eq("dispatch_date", dateStr)
    .order("actual_hour", { ascending: true })
    .order("stop_order", { ascending: true })
    .order("id", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  currentDispatchId = `single:${dateStr}`;
  const scopedRows = filterDispatchRowsByWorkspace(data, workspaceTeamId);
  const hydratedRows = scopedRows.map(mapUnifiedDispatchRowToActual).map(enrichUnifiedDispatchRowWithCast);
  currentActualsCache = await applyCurrentOriginMetricsToDispatchRows(hydratedRows);
  syncScheduleRendererDeps();
  renderActualTable();
  renderActualTimeAreaMatrix();
  renderHomeSummary();
  renderCastSelects();
  renderManualLastVehicleInfo();
}

async function savePlanSingle() {
  const selectedCast = resolveLinkedCastFromInput(els.planCastSelect, getPlanSelectableCasts);
  const castId = normalizeDispatchEntityId(selectedCast?.id || selectedCast?.cast_id) || null;
  if (!castId) {
    alert("送り先を選択または入力してください");
    return;
  }

  const cast = (Array.isArray(allCastsCache) ? allCastsCache.find(c => sameDispatchEntityId(c.id || c.cast_id, castId)) : null) || selectedPlanCast;
  if (!cast?.name) {
    alert("選択した送り先情報を取得できませんでした。もう一度選択してください");
    return;
  }

  const planDate = els.planDate?.value || todayStr();
  const hour = Number(els.planHour?.value || 0);
  const address = els.planAddress?.value.trim() || cast.address || "";
  const area = normalizeAreaLabel((els.planArea?.value || cast.area || "無し").trim() || "無し");
  const note = els.planNote?.value.trim() || "";
  const snap = await buildDispatchMetricSnapshot(cast, address, area);
  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();

  const payload = {
    team_id: workspaceTeamId,
    dispatch_kind: "plan",
    dispatch_date: planDate,
    plan_date: planDate,
    dispatch_hour: hour,
    plan_hour: hour,
    cast_id: castId,
    person_id: castId,
    person_name: cast.name || "",
    destination_address: snap.destination_address,
    planned_area: area,
    destination_area: area,
    straight_km: snap.straight_km,
    distance_km: snap.distance_km,
    travel_minutes: snap.travel_minutes,
    origin_slot_no: snap.origin_slot_no,
    origin_label: snap.origin_label,
    origin_address: snap.origin_address,
    origin_lat: snap.origin_lat,
    origin_lng: snap.origin_lng,
    destination_lat: snap.destination_lat,
    destination_lng: snap.destination_lng,
    note,
    status: "planned"
  };

  let error;
  if (editingPlanId) {
    ({ error } = await supabaseClient.from(getDispatchUnifiedTableName()).update(payload).eq("id", editingPlanId).eq("dispatch_kind", "plan"));
  } else {
    ({ error } = await supabaseClient.from(getDispatchUnifiedTableName()).insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, editingPlanId ? "update_plan" : "create_plan", editingPlanId ? "予定を更新" : "予定を作成");
  resetPlanForm();
  await loadPlansByDate(planDate);
}

async function deletePlanSingle(planId) {
  const safePlanId = resolveDispatchActionId(planId, [".plan-delete-btn", ".plan-edit-btn"]);
  if (!safePlanId) {
    alert("削除対象の予定IDを取得できませんでした");
    return;
  }
  if (!window.confirm("この予定を削除しますか？")) return;
  const { error } = await supabaseClient.from(getDispatchUnifiedTableName()).delete().eq("id", safePlanId).eq("dispatch_kind", "plan");
  if (error) {
    alert(error.message);
    return;
  }
  await addHistory(null, null, "delete_plan", `予定ID ${safePlanId} を削除`);
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function clearAllPlansSingle() {
  if (!window.confirm("この日の予定を全消去しますか？")) return;
  const planDate = els.planDate?.value || todayStr();
  const targetIds = currentPlansCache
    .map(row => normalizeDispatchEntityId(row?.id))
    .filter(Boolean);

  if (!targetIds.length) {
    await loadPlansByDate(planDate);
    return;
  }

  const { error } = await supabaseClient
    .from(getDispatchUnifiedTableName())
    .delete()
    .in("id", targetIds);
  if (error) {
    alert(error.message);
    return;
  }
  await addHistory(null, null, "clear_plans", `${planDate} の予定を全削除`);
  await loadPlansByDate(planDate);
}

async function saveActualSingle() {
  const inputValue = els.castSelect?.value || "";
  const typedCast = findCastByInputValue(inputValue);
  const selectedCast = resolveLinkedCastFromInput(els.castSelect, getActualSelectableCasts);
  const castId = normalizeDispatchEntityId(selectedCast?.id || selectedCast?.cast_id) || null;
  if (!castId) {
    if (typedCast) {
      alert("この送り先はすでに実際の送りに入っているため、再選択できません");
    } else {
      alert("送り先を選択または入力してください");
    }
    return;
  }

  const cast = (Array.isArray(allCastsCache) ? allCastsCache.find(c => sameDispatchEntityId(c.id || c.cast_id, castId)) : null) || selectedCast;
  if (!cast?.name) {
    alert("選択した送り先情報を取得できませんでした。もう一度選択してください");
    return;
  }

  const dateStr = els.actualDate?.value || todayStr();
  const hour = Number(els.actualHour?.value || 0);
  const address = els.actualAddress?.value.trim() || cast.address || "";
  const area = normalizeAreaLabel(els.actualArea?.value.trim() || cast.area || "無し");
  const status = els.actualStatus?.value || "pending";
  const note = els.actualNote?.value.trim() || "";
  const existingActual = editingActualId ? currentActualsCache.find(x => sameDispatchEntityId(x.id, editingActualId)) : null;
  const stopOrder = existingActual
    ? Number(existingActual.stop_order || 1)
    : currentActualsCache.filter(x => Number(x.actual_hour) === hour && !sameDispatchEntityId(x.id, editingActualId)).length + 1;
  const snap = await buildDispatchMetricSnapshot(cast, address, area);
  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();

  const payload = {
    team_id: workspaceTeamId,
    dispatch_kind: "actual",
    dispatch_date: dateStr,
    plan_date: dateStr,
    dispatch_hour: hour,
    actual_hour: hour,
    cast_id: castId,
    person_id: castId,
    person_name: cast.name || "",
    stop_order: stopOrder,
    pickup_label: snap.origin_label,
    origin_slot_no: snap.origin_slot_no,
    origin_label: snap.origin_label,
    origin_address: snap.origin_address,
    origin_lat: snap.origin_lat,
    origin_lng: snap.origin_lng,
    destination_address: snap.destination_address,
    destination_area: area,
    destination_lat: snap.destination_lat,
    destination_lng: snap.destination_lng,
    straight_km: snap.straight_km,
    distance_km: snap.distance_km,
    travel_minutes: snap.travel_minutes,
    status,
    note
  };

  let error;
  if (editingActualId) {
    ({ error } = await supabaseClient.from(getDispatchUnifiedTableName()).update(payload).eq("id", editingActualId).eq("dispatch_kind", "actual"));
  } else {
    ({ error } = await supabaseClient.from(getDispatchUnifiedTableName()).insert(payload));
  }

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(currentDispatchId, editingActualId || null, editingActualId ? "update_actual" : "create_actual", editingActualId ? "実際の送りを更新" : "実際の送りを追加");
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

async function deleteActualSingle(itemId) {
  const safeItemId = resolveDispatchActionId(itemId, [".actual-delete-btn", ".action-delete-btn", ".btn.danger"]);
  if (!safeItemId) {
    alert("削除対象のActual IDを取得できませんでした");
    return;
  }
  if (!window.confirm("このActualを削除しますか？")) return;
  const { error } = await supabaseClient.from(getDispatchUnifiedTableName()).delete().eq("id", safeItemId).eq("dispatch_kind", "actual");
  if (error) {
    alert(error.message);
    return;
  }
  await addHistory(currentDispatchId, safeItemId, "delete_actual", `Actual ID ${safeItemId} を削除`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

async function updateActualStatusSingle(itemId, status) {
  const safeItemId = resolveDispatchActionId(itemId, [".actual-done-btn", ".actual-pending-btn", ".actual-cancel-btn", ".status-done-btn", ".status-pending-btn", ".status-cancel-btn", ".btn"]);
  if (!safeItemId) {
    alert("更新対象のActual IDを取得できませんでした");
    return;
  }
  const item = currentActualsCache.find(x => sameDispatchEntityId(x.id, safeItemId));
  if (!item) {
    alert("対象のActualが見つかりません");
    return;
  }
  const { error } = await supabaseClient.from(getDispatchUnifiedTableName()).update({ status }).eq("id", safeItemId).eq("dispatch_kind", "actual");
  if (error) {
    alert(error.message);
    return;
  }
  const targetPlan = currentPlansCache.find(plan => sameDispatchEntityId(plan.cast_id, item.cast_id) && plan.plan_date === (els.actualDate?.value || todayStr()) && Number(plan.plan_hour) === Number(item.actual_hour ?? -1));
  if (targetPlan) {
    let nextPlanStatus = targetPlan.status;
    if (status === "done") nextPlanStatus = "done";
    else if (status === "cancel") nextPlanStatus = "cancel";
    else if (status === "pending") nextPlanStatus = "assigned";
    const { error: planError } = await supabaseClient.from(getDispatchUnifiedTableName()).update({ status: nextPlanStatus }).eq("id", targetPlan.id).eq("dispatch_kind", "plan");
    if (planError) console.error(planError);
  }
  await addHistory(currentDispatchId, safeItemId, "update_actual_status", `Actual状態を ${status} に変更`);
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
}

function getAddablePlansForActual(plans) {
  const targetDate = String(els.actualDate?.value || todayStr()).trim();
  const doneCastIds = getDoneCastIdsInActuals();
  return (Array.isArray(plans) ? plans : [])
    .filter(plan => String(plan?.plan_date || plan?.dispatch_date || "").trim() === targetDate)
    .filter(plan => {
      const rawStatus = String(plan?.status || "").trim().toLowerCase();
      const normalized = typeof normalizeStatus === "function" ? normalizeStatus(plan?.status) : rawStatus || "pending";
      return rawStatus === "planned" || normalized === "pending";
    })
    .filter(plan => !doneCastIds.has(normalizeDispatchEntityId(plan?.cast_id)))
    .filter(plan => !isPlanAlreadyAddedToActual(plan));
}

function closeAddFromPlansDialog() {
  addFromPlansDialogCache = [];
  if (els.addFromPlansDialog) els.addFromPlansDialog.classList.add("hidden");
}

function renderAddFromPlansDialogRows(plans, dateStr) {
  if (!els.addFromPlansList || !els.addFromPlansEmpty || !els.addFromPlansDateLabel) return;

  const rows = Array.isArray(plans) ? plans : [];
  addFromPlansDialogCache = rows;
  els.addFromPlansDateLabel.textContent = `${dateStr} の予定`;
  els.addFromPlansList.innerHTML = "";

  if (!rows.length) {
    els.addFromPlansEmpty.classList.remove("hidden");
    return;
  }

  els.addFromPlansEmpty.classList.add("hidden");

  rows.forEach(plan => {
    const row = document.createElement("label");
    row.className = "plan-picker-row";
    row.innerHTML = `
      <input type="checkbox" data-plan-id="${escapeHtml(String(plan.id || ""))}">
      <div class="plan-picker-meta">
        <div class="plan-picker-main">${escapeHtml(getHourLabel(plan.plan_hour))} / ${escapeHtml(plan.person_name || plan.casts?.name || "-")}</div>
        <div class="plan-picker-sub">${escapeHtml(normalizeAreaLabel(plan.planned_area || plan.destination_area || "無し"))} / ${escapeHtml(plan.destination_address || plan.casts?.address || "住所未設定")}</div>
      </div>
    `;
    els.addFromPlansList.appendChild(row);
  });
}

async function fetchPlansForActualAddDialog(dateStr) {
  const targetDate = String(dateStr || els.actualDate?.value || todayStr()).trim();
  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();
  const { data, error } = await supabaseClient
    .from(getDispatchUnifiedTableName())
    .select(remapRelationSelect(`
      *,
      casts (
        id,
        team_id,
        name,
        phone,
        address,
        area,
        distance_km,
        travel_minutes,
        latitude,
        longitude,
        lat,
        lng
      )
    `))
    .eq("dispatch_kind", "plan")
    .eq("dispatch_date", targetDate)
    .order("dispatch_hour", { ascending: true })
    .order("id", { ascending: true });

  if (error) throw error;

  const scopedRows = filterDispatchRowsByWorkspace(data, workspaceTeamId);
  return scopedRows.map(mapUnifiedDispatchRowToPlan).map(enrichUnifiedDispatchRowWithCast);
}

async function openAddFromPlansDialog() {
  const targetDate = String(els.actualDate?.value || todayStr()).trim();
  try {
    const freshPlans = await fetchPlansForActualAddDialog(targetDate);
    const addablePlans = getAddablePlansForActual(freshPlans);
    renderAddFromPlansDialogRows(addablePlans, targetDate);
    if (els.addFromPlansDialog) els.addFromPlansDialog.classList.remove("hidden");
  } catch (error) {
    console.error(error);
    alert(error?.message || "予定の読込に失敗しました");
  }
}

async function addPlanToActualBySource(planSource, options = {}) {
  const { skipReload = false, silent = false } = options;
  const plan = typeof planSource === "string"
    ? currentPlansCache.find(x => String(x?.id || "") === String(planSource).trim())
    : planSource;

  if (!plan) {
    if (!silent) alert("予定が見つかりません");
    return { ok: false, reason: "missing" };
  }
  if (isPlanAlreadyAddedToActual(plan)) {
    if (!silent) alert("その予定はすでにActualへ追加されています");
    return { ok: false, reason: "duplicate" };
  }
  if (currentActualsCache.some(x => sameDispatchEntityId(x.cast_id, plan.cast_id) && normalizeStatus(x.status) !== "cancel")) {
    if (!silent) alert("その送り先はすでにActualにあります");
    return { ok: false, reason: "cast_exists" };
  }
  const doneCastIds = getDoneCastIdsInActuals();
  if (doneCastIds.has(normalizeDispatchEntityId(plan.cast_id))) {
    if (!silent) alert("この送り先はすでに送り完了です");
    return { ok: false, reason: "done" };
  }

  const workspaceTeamId = await resolveCurrentWorkspaceTeamId();

  const payload = {
    team_id: workspaceTeamId || plan.team_id || plan.casts?.team_id || null,
    dispatch_kind: "actual",
    dispatch_date: plan.plan_date || plan.dispatch_date || (els.actualDate?.value || todayStr()),
    plan_date: plan.plan_date || plan.dispatch_date || (els.actualDate?.value || todayStr()),
    dispatch_hour: Number(plan.plan_hour || 0),
    actual_hour: Number(plan.plan_hour || 0),
    cast_id: plan.cast_id,
    person_id: plan.cast_id,
    person_name: plan.person_name || plan.casts?.name || "",
    stop_order: currentActualsCache.filter(x => Number(x.actual_hour) === Number(plan.plan_hour || 0)).length + 1,
    pickup_label: plan.pickup_label || plan.origin_label || ORIGIN_LABEL,
    origin_slot_no: plan.origin_slot_no ?? null,
    origin_label: plan.origin_label || ORIGIN_LABEL,
    origin_address: plan.origin_address || "",
    origin_lat: plan.origin_lat ?? null,
    origin_lng: plan.origin_lng ?? null,
    destination_address: plan.destination_address || plan.casts?.address || "",
    destination_area: normalizeAreaLabel(plan.planned_area || plan.destination_area || "無し"),
    destination_lat: toNullableNumber(plan.destination_lat ?? plan.casts?.latitude ?? plan.casts?.lat),
    destination_lng: toNullableNumber(plan.destination_lng ?? plan.casts?.longitude ?? plan.casts?.lng),
    straight_km: toNullableNumber(plan.straight_km),
    distance_km: toNullableNumber(plan.distance_km),
    travel_minutes: Number(plan.travel_minutes || 0) || estimateTravelMinutesByAreaSpeed(Number(plan.distance_km || 0), plan.planned_area || plan.destination_area || ""),
    status: "pending",
    note: plan.note || ""
  };

  const { error } = await supabaseClient.from(getDispatchUnifiedTableName()).insert(payload);
  if (error) {
    if (!silent) alert(error.message);
    return { ok: false, reason: "insert_error", error };
  }

  await supabaseClient.from(getDispatchUnifiedTableName()).update({ status: "assigned" }).eq("id", plan.id).eq("dispatch_kind", "plan");
  await addHistory(currentDispatchId, null, "add_plan_to_actual", `予定ID ${plan.id} をActualへ追加`);

  if (!skipReload) {
    await loadActualsByDate(els.actualDate?.value || todayStr());
    await loadPlansByDate(els.planDate?.value || todayStr());
    if (els.planSelect) els.planSelect.value = "";
    renderPlanSelect();
  }

  return { ok: true };
}

async function addPlanToActualSingle() {
  const planId = String(els.planSelect?.value || "").trim();
  if (!planId) {
    alert("予定を選択してください");
    return;
  }
  await addPlanToActualBySource(planId);
}

async function clearAllActualsSingle() {
  if (!window.confirm("この日のActualを全消去しますか？")) return;
  const targetDate = els.actualDate?.value || todayStr();
  const targetIds = currentActualsCache
    .map(row => normalizeDispatchEntityId(row?.id))
    .filter(Boolean);

  if (targetIds.length) {
    const { error } = await supabaseClient
      .from(getDispatchUnifiedTableName())
      .delete()
      .in("id", targetIds);
    if (error) {
      alert(error.message);
      return;
    }
  }

  const assignedPlanIds = currentPlansCache
    .filter(row => normalizeStatus(row?.status) === "assigned")
    .map(row => normalizeDispatchEntityId(row?.id))
    .filter(Boolean);

  if (assignedPlanIds.length) {
    const { error: planResetError } = await supabaseClient
      .from(getDispatchUnifiedTableName())
      .update({ status: "planned" })
      .in("id", assignedPlanIds);
    if (planResetError) {
      alert(planResetError.message);
      return;
    }
  }

  await addHistory(currentDispatchId, null, "clear_actual", "Actualを全消去");
  await loadActualsByDate(targetDate);
  await loadPlansByDate(els.planDate?.value || targetDate);
}

async function loadPlansByDate(dateStr) {
  return await loadPlansByDateSingle(dateStr);
}


function getPlanSelectableCasts() {
  const plannedIds = getPlannedCastIds();
  const editingPlan = editingPlanId
    ? currentPlansCache.find(x => String(x.id) === String(editingPlanId))
    : null;
  const editingCastIdForPlan = normalizeDispatchEntityId(editingPlan?.cast_id);

  return allCastsCache.filter(cast => {
    const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
    return sameDispatchEntityId(castId, editingCastIdForPlan) || !plannedIds.has(castId);
  });
}

function getActualSelectableCasts() {
  const usedCastIds = new Set();
  const doneCastIds = getDoneCastIdsInActuals();

  currentActualsCache.forEach(item => {
    const castId = normalizeDispatchEntityId(item.cast_id);
    if (castId && normalizeStatus(item.status) !== "cancel") {
      usedCastIds.add(castId);
    }
  });

  const editingActual = editingActualId
    ? currentActualsCache.find(x => String(x.id) === String(editingActualId))
    : null;
  const editingCastIdForActual = normalizeDispatchEntityId(editingActual?.cast_id);

  return allCastsCache.filter(cast => {
    const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
    return sameDispatchEntityId(castId, editingCastIdForActual) || (!usedCastIds.has(castId) && !doneCastIds.has(castId));
  });
}

function findActualSelectableCastByInputValue(value) {
  const cast = findCastByInputValue(value);
  if (!cast) return null;
  const selectable = getActualSelectableCasts();
  const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
  return selectable.find(x => sameDispatchEntityId(x.id || x.cast_id, castId)) || null;
}

function getCastSearchText(cast) {
  return [
    String(cast.name || "").trim(),
    normalizeAreaLabel(cast.area || "-"),
    String(cast.address || "").trim()
  ].join(" / ");
}

function filterCastCandidates(casts, query) {
  const q = String(query || "").trim().toLowerCase();

  const sorted = [...casts].sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "ja")
  );

  if (!q) return sorted;

  return sorted.filter(cast => {
    const hay = [
      cast.name || "",
      cast.address || "",
      cast.area || "",
      cast.phone || "",
      cast.memo || ""
    ]
      .join(" ")
      .toLowerCase();

    return hay.includes(q);
  });
}

function renderCastSearchSuggest(container, casts, onPick) {
  if (!container) return;

  if (!casts.length) {
    container.innerHTML = "";
    container.classList.add("hidden");
    return;
  }

  container.innerHTML = casts
    .map(
      cast => {
        const castId = normalizeDispatchEntityId(cast.id || cast.cast_id);
        return `
        <button type="button" class="cast-search-item" data-id="${escapeHtml(castId)}">
          <span>${escapeHtml(cast.name || "-")}</span>
          <small>${escapeHtml(normalizeAreaLabel(cast.area || "-"))} / ${escapeHtml(cast.address || "")}</small>
        </button>
      `;
      }
    )
    .join("");

  container.classList.remove("hidden");

  container.querySelectorAll(".cast-search-item").forEach(btn => {
    btn.addEventListener("mousedown", event => {
      event.preventDefault();
      const cast = (Array.isArray(casts) ? casts : []).find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id))
        || (Array.isArray(allCastsCache) ? allCastsCache : []).find(x => sameDispatchEntityId(x.id || x.cast_id, btn.dataset.id));
      if (cast) onPick(cast);
      container.classList.add("hidden");
    });
  });
}

function setupSearchableCastInput(input, suggest, getCandidates, onPick) {
  if (!input || !suggest) return;
  if (input.dataset.searchBound === "1") return;
  input.dataset.searchBound = "1";

  const openSuggest = () => {
    const casts = filterCastCandidates(getCandidates(), input.value || "");
    renderCastSearchSuggest(suggest, casts, onPick);
  };

  input.addEventListener("focus", openSuggest);
  input.addEventListener("input", openSuggest);

  input.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      suggest.classList.add("hidden");
      return;
    }

    if (event.key === "Enter") {
      const exact = findCastByInputValue(input.value || "");
      if (exact) {
        onPick(exact);
        suggest.classList.add("hidden");
        return;
      }

      const candidates = filterCastCandidates(getCandidates(), input.value || "");
      if (candidates.length === 1) {
        onPick(candidates[0]);
        suggest.classList.add("hidden");
      }
    }
  });

  input.addEventListener("blur", () => {
    window.setTimeout(() => {
      suggest.classList.add("hidden");
    }, 150);
  });
}

function setupSearchableCastInputs() {
  setupSearchableCastInput(
    els.planCastSelect,
    els.planCastSuggest,
    getPlanSelectableCasts,
    cast => {
      if (els.planCastSelect) setLinkedCastSelection(els.planCastSelect, cast);
      fillPlanFormFromSelectedCast();
    }
  );

  setupSearchableCastInput(
    els.castSelect,
    els.castSuggest,
    getActualSelectableCasts,
    cast => {
      if (els.castSelect) setLinkedCastSelection(els.castSelect, cast);
      fillActualFormFromSelectedCast();
    }
  );
}

function renderPlanCastSelect() {
  const input = els.planCastSelect;
  const list = document.getElementById("planCastList");
  if (!input || !list) return;

  const editingPlan = editingPlanId
    ? currentPlansCache.find(x => String(x.id) === String(editingPlanId))
    : null;

  list.innerHTML = "";

  getPlanSelectableCasts().forEach(cast => {
    const option = document.createElement("option");
    option.value = String(cast.name || "").trim();
    option.label = getCastSearchText(cast);
    list.appendChild(option);
  });

  if (editingPlan) {
    input.value = editingPlan.casts?.name || editingPlan.person_name || "";
    if (editingPlan.cast_id) input.dataset.selectedCastId = normalizeDispatchEntityId(editingPlan.cast_id);
  }
}

function isPlanAlreadyAddedToActual(plan, excludeActualId = null) {
  if (!plan) return false;

  const targetDate = String(plan.plan_date || els.actualDate?.value || todayStr()).trim();
  const targetCastId = normalizeDispatchEntityId(plan.cast_id || plan.person_id || plan.casts?.id);
  const targetHour = Number(plan.plan_hour || 0);
  const targetAddress = String(plan.destination_address || plan.casts?.address || "").trim();
  const safeExcludeActualId = excludeActualId != null ? normalizeDispatchEntityId(excludeActualId) : null;

  return currentActualsCache.some(item => {
    if (safeExcludeActualId && sameDispatchEntityId(item?.id, safeExcludeActualId)) return false;

    const itemDate = String(item.plan_date || els.actualDate?.value || todayStr()).trim();
    const itemCastId = normalizeDispatchEntityId(item.cast_id || item.person_id || item.casts?.id);
    const itemHour = Number(item.actual_hour || 0);
    const itemAddress = String(item.destination_address || item.casts?.address || "").trim();

    if (itemDate !== targetDate) return false;

    const sameCastHour = sameDispatchEntityId(itemCastId, targetCastId) && itemHour === targetHour;
    const sameAddressHour = !!targetAddress && itemAddress === targetAddress && itemHour === targetHour;
    const sameCastAddress = sameDispatchEntityId(itemCastId, targetCastId) && !!targetAddress && itemAddress === targetAddress;

    return sameCastHour || sameAddressHour || sameCastAddress;
  });
}

function getLinkedPlanForActual(actualItem) {
  if (!actualItem) return null;

  const actualDate = String(actualItem.plan_date || els.actualDate?.value || todayStr()).trim();
  const actualCastId = normalizeDispatchEntityId(actualItem.cast_id);
  const actualHour = Number(actualItem.actual_hour || 0);
  const actualAddress = String(actualItem.destination_address || actualItem.casts?.address || "").trim();

  return currentPlansCache.find(plan => {
    const planDate = String(plan.plan_date || "").trim();
    const planCastId = normalizeDispatchEntityId(plan.cast_id);
    const planHour = Number(plan.plan_hour || 0);
    const planAddress = String(plan.destination_address || plan.casts?.address || "").trim();

    if (planDate !== actualDate) return false;

    const sameCastHour = sameDispatchEntityId(planCastId, actualCastId) && planHour === actualHour;
    const sameAddressHour = !!actualAddress && planAddress === actualAddress && planHour === actualHour;
    const sameCastAddress = sameDispatchEntityId(planCastId, actualCastId) && !!actualAddress && planAddress === actualAddress;

    return sameCastHour || sameAddressHour || sameCastAddress;
  }) || null;
}

function renderPlanSelect() {
  if (!els.planSelect) return;

  const targetDate = els.actualDate?.value || todayStr();
  const doneCastIds = getDoneCastIdsInActuals();
  const editingActual = editingActualId
    ? currentActualsCache.find(x => String(x.id) === String(editingActualId))
    : null;
  const editingPlan = getLinkedPlanForActual(editingActual);
  const selectedValueBefore = String(els.planSelect.value || "");
  const appendedPlanIds = new Set();

  els.planSelect.innerHTML = `<option value="">予定から選択</option>`;

  const appendOption = plan => {
    const planId = String(plan?.id || "");
    if (!plan || !planId || appendedPlanIds.has(planId)) return;
    appendedPlanIds.add(planId);

    const option = document.createElement("option");
    option.value = planId;
    option.textContent = `${getHourLabel(plan.plan_hour)} / ${plan.casts?.name || "-"} / ${normalizeAreaLabel(plan.planned_area || "-")}`;
    if (editingPlan && String(plan.id) === String(editingPlan.id) && editingActualId) {
      option.textContent += " [編集中]";
    }
    els.planSelect.appendChild(option);
  };

  currentPlansCache
    .filter(plan => plan.plan_date === targetDate)
    .filter(plan => plan.status === "planned" || (editingPlan && String(plan.id) === String(editingPlan.id)))
    .filter(plan => !doneCastIds.has(normalizeDispatchEntityId(plan.cast_id)) || (editingPlan && String(plan.id) === String(editingPlan.id)))
    .filter(plan => !isPlanAlreadyAddedToActual(plan, editingActualId || null) || (editingPlan && String(plan.id) === String(editingPlan.id)))
    .forEach(appendOption);

  if (editingPlan) {
    appendOption(editingPlan);
  }

  if (editingPlan) {
    els.planSelect.value = String(editingPlan.id);
  } else if (selectedValueBefore && appendedPlanIds.has(String(selectedValueBefore))) {
    els.planSelect.value = selectedValueBefore;
  } else {
    els.planSelect.value = "";
  }
}

async function savePlan() {
  return await savePlanSingle();
}

async function deletePlan(planId) {
  return await deletePlanSingle(planId);
}

function guessPlanArea() {
  if (els.planArea) {
    els.planArea.value = normalizeAreaLabel(
      classifyAreaByAddress(els.planAddress?.value || "") || "無し"
    );
  }
}

async function clearAllPlans() {
  return await clearAllPlansSingle();
}

function resetActualForm() {
  editingActualId = null;
  if (els.planSelect) els.planSelect.value = "";
  if (els.castSelect) {
    els.castSelect.value = "";
    clearLinkedCastSelection(els.castSelect);
  }
  if (els.actualHour) els.actualHour.value = "0";
  if (els.actualStatus) els.actualStatus.value = "pending";
  clearActualCastDerivedFields();
  if (els.actualNote) els.actualNote.value = "";
}

function fillActualForm(item) {
  editingActualId = item.id;
  renderPlanSelect();
  if (els.castSelect) {
    els.castSelect.value = item.casts?.name || item.person_name || "";
    if (item.cast_id) els.castSelect.dataset.selectedCastId = normalizeDispatchEntityId(item.cast_id);
  }
  if (els.actualHour) els.actualHour.value = String(item.actual_hour ?? 0);
  if (els.actualDistanceKm) els.actualDistanceKm.value = item.distance_km ?? "";
  if (els.actualStatus) els.actualStatus.value = item.status || "pending";
  if (els.actualAddress) els.actualAddress.value = item.destination_address || item.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(item.destination_area || item.casts?.area || "");
  if (els.actualNote) els.actualNote.value = item.note || "";
  const linkedPlan = getLinkedPlanForActual(item);
  if (els.planSelect) els.planSelect.value = linkedPlan ? String(linkedPlan.id) : "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function fillActualFormFromSelectedCast() {
  const cast = resolveLinkedCastFromInput(els.castSelect, getActualSelectableCasts);
  if (!cast) return;

  const address = cast.address || "";
  const area = normalizeAreaLabel(
    cast.area ||
      guessArea(
        toNullableNumber(cast.latitude),
        toNullableNumber(cast.longitude),
        address
      )
  );

  if (els.actualAddress && !els.actualAddress.value.trim()) {
    els.actualAddress.value = address;
  }

  if (els.actualArea && !els.actualArea.value.trim()) {
    els.actualArea.value = area;
  }

  if (els.actualDistanceKm && !els.actualDistanceKm.value.trim()) {
    const distance = await resolveFormDisplayDistanceKm(cast, address, area);
    els.actualDistanceKm.value = distance ?? "";
  }
}

function fillActualFormFromSelectedPlan() {
  const planId = String(els.planSelect?.value || "").trim();
  if (!planId) return;

  const plan = currentPlansCache.find(p => String(p?.id || "") === planId);
  if (!plan) return;

  if (els.castSelect) {
    const castName = plan.person_name || plan.casts?.name || "";
    els.castSelect.value = castName;
    const linkedCastId = normalizeDispatchEntityId(plan.cast_id || plan.person_id);
    if (linkedCastId) {
      els.castSelect.dataset.selectedCastId = linkedCastId;
    } else {
      clearLinkedCastSelection(els.castSelect);
    }
  }
  if (els.actualHour) els.actualHour.value = String(plan.plan_hour ?? plan.dispatch_hour ?? 0);
  if (els.actualAddress) els.actualAddress.value = plan.destination_address || plan.casts?.address || "";
  if (els.actualArea) els.actualArea.value = normalizeAreaLabel(plan.planned_area || plan.destination_area || plan.casts?.area || "無し");
  if (els.actualDistanceKm) els.actualDistanceKm.value = plan.distance_km ?? plan.casts?.distance_km ?? "";
  if (els.actualNote) els.actualNote.value = plan.note || "";
}

function renderCastSelects() {
  const editingActual = editingActualId
    ? currentActualsCache.find(x => sameDispatchEntityId(x.id, editingActualId))
    : null;

  const input = els.castSelect;
  const list = document.getElementById("castList");

  if (input && list) {
    list.innerHTML = "";

    getActualSelectableCasts().forEach(cast => {
      const option = document.createElement("option");
      option.value = String(cast.name || "").trim();
      option.label = getCastSearchText(cast);
      list.appendChild(option);
    });

    if (editingActual?.casts?.name) {
      input.value = editingActual.casts.name;
    }
  }

  renderPlanCastSelect();
  setupSearchableCastInputs();
}

async function loadActualsByDate(dateStr) {
  return await loadActualsByDateSingle(dateStr);
}

async function saveActual() {
  return await saveActualSingle();
}

async function deleteActual(itemId) {
  return await deleteActualSingle(itemId);
}

async function updateActualStatus(itemId, status) {
  return await updateActualStatusSingle(itemId, status);
}

async function addPlanToActual() {
  await openAddFromPlansDialog();
}
function guessActualArea() {
  if (els.actualArea) {
    els.actualArea.value = normalizeAreaLabel(
      classifyAreaByAddress(els.actualAddress?.value || "") || "無し"
    );
  }
}

function renderDailyVehicleChecklist() {
  if (!els.dailyVehicleChecklist) return;

  els.dailyVehicleChecklist.innerHTML = "";

  if (!allVehiclesCache.length) {
    els.dailyVehicleChecklist.innerHTML = `<div class="muted">車両がありません</div>`;
    return;
  }

  const monthKey = getMonthKey(els.dispatchDate?.value || todayStr());
  const monthlyStatsMap = getVehicleMonthlyStatsMap(currentDailyReportsCache, monthKey);

  const header = document.createElement("div");
  header.className = "vehicle-check-header";
  header.innerHTML = `
    <div class="vehicle-check-header-info"></div>
    <div class="vehicle-check-header-col">可能車両</div>
    <div class="vehicle-check-header-col">ラスト便</div>
  `;
  els.dailyVehicleChecklist.appendChild(header);

  getSortedVehiclesForDisplay().forEach(vehicle => {
    const stats = monthlyStatsMap.get(Number(vehicle.id)) || {
      totalDistance: 0,
      workedDays: 0,
      avgDistance: 0
    };
    const avgDistanceText = `${Number(stats.avgDistance || 0).toFixed(1)}km`;

    const row = document.createElement("div");
    row.className = "vehicle-check-item";
    row.innerHTML = `
      <div class="vehicle-check-info">
        <div class="vehicle-check-name">${escapeHtml(vehicle.driver_name || "-")}</div>
        <div class="vehicle-check-car">車両 ${escapeHtml(vehicle.plate_number || "-")}</div>
        <div class="vehicle-check-meta">担当 ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))} / 帰宅 ${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))} / 定員 ${vehicle.seat_capacity ?? "-"} / 1日平均距離 ${avgDistanceText}</div>
      </div>
      <label class="vehicle-check-toggle vehicle-check-toggle-work">
        <input class="vehicle-check-input" type="checkbox" data-id="${vehicle.id}" ${activeVehicleIdsForToday.has(Number(vehicle.id)) ? "checked" : ""} />
        <span>可能車両</span>
      </label>
      <label class="vehicle-check-toggle vehicle-check-toggle-last">
        <input class="driver-last-trip-input" type="checkbox" data-id="${vehicle.id}" ${isDriverLastTripChecked(vehicle.id) ? "checked" : ""} />
        <span>ラスト便</span>
      </label>
    `;
    els.dailyVehicleChecklist.appendChild(row);
  });

  renderOperationAndSimulationUI();
  els.dailyVehicleChecklist.querySelectorAll(".vehicle-check-input").forEach(input => {
    input.addEventListener("change", () => {
      const id = Number(input.dataset.id);

      if (input.checked) activeVehicleIdsForToday.add(id);
      else activeVehicleIdsForToday.delete(id);

      renderDailyMileageInputs();
      renderDailyDispatchResult();
      renderDailyVehicleChecklist();
    });
  });

  els.dailyVehicleChecklist.querySelectorAll(".driver-last-trip-input").forEach(input => {
    input.addEventListener("change", () => {
      const id = Number(input.dataset.id);
      setDriverLastTripChecked(id, input.checked);

      if (input.checked && !activeVehicleIdsForToday.has(id)) {
        activeVehicleIdsForToday.add(id);
      }

      renderDailyMileageInputs();
      renderDailyDispatchResult();
      renderDailyVehicleChecklist();
    });
  });
}

function getSelectedVehiclesForToday() {
  return getSortedVehiclesForDisplay().filter(v => activeVehicleIdsForToday.has(Number(v.id)));
}

function toggleAllVehicles(checked) {
  if (checked) {
    activeVehicleIdsForToday = new Set(allVehiclesCache.map(v => Number(v.id)));
  } else {
    activeVehicleIdsForToday = new Set();
  }
  renderDailyVehicleChecklist();
  renderDailyMileageInputs();
  renderDailyDispatchResult();
}


function getActiveDispatchItemsForAutoAssign() {
  return currentActualsCache.filter(
    item => !["done", "cancel"].includes(normalizeStatus(item.status))
  );
}

function getCurrentClockMinutes() {
  const now = new Date();
  return now.getHours() * 60 + now.getMinutes();
}


const AREA_AVERAGE_SPEED_KMH = {
  "松戸近郊": 33,
  "柏方面": 33,
  "柏の葉方面": 34,
  "流山方面": 33,
  "野田方面": 34,
  "我孫子方面": 36,
  "取手方面": 39,
  "藤代方面": 39,
  "守谷方面": 39,
  "牛久方面": 39,
  "葛飾方面": 27,
  "足立方面": 27,
  "江戸川方面": 27,
  "墨田方面": 27,
  "江東方面": 27,
  "荒川方面": 27,
  "台東方面": 27,
  "市川方面": 31,
  "船橋方面": 31,
  "鎌ヶ谷方面": 31,
  "三郷方面": 34,
  "八潮方面": 34,
  "草加方面": 34,
  "吉川方面": 34,
  "越谷方面": 34,
  "千葉方面": 31
};

function normalizeAreaSpeedLookupInput(input) {
  if (Array.isArray(input)) return getRepresentativeAreaFromRows(input);
  if (input && typeof input === "object") {
    return normalizeAreaLabel(
      input.destination_area ||
      input.planned_area ||
      input.cluster_area ||
      input.area ||
      input.home_area ||
      input.vehicle_area ||
      ""
    );
  }
  return normalizeAreaLabel(input || "");
}

function getRepresentativeAreaFromRows(rows) {
  const list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!list.length) return "";

  const counts = new Map();
  list.forEach(row => {
    const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.cluster_area || row?.casts?.area || "");
    if (!area || area === "無し") return;
    counts.set(area, Number(counts.get(area) || 0) + 1);
  });

  if (!counts.size) {
    return normalizeAreaLabel(
      list[list.length - 1]?.destination_area ||
      list[list.length - 1]?.planned_area ||
      list[list.length - 1]?.cluster_area ||
      list[list.length - 1]?.casts?.area ||
      ""
    );
  }

  const sorted = [...counts.entries()].sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0], "ja");
  });

  return normalizeAreaLabel(sorted[0]?.[0] || "");
}

function getAreaAverageSpeedKmh(areaInput, distanceKm = 0) {
  const rawArea = normalizeAreaSpeedLookupInput(areaInput);
  const canonical = getCanonicalArea(rawArea);

  if (AREA_AVERAGE_SPEED_KMH[canonical]) {
    return Number(AREA_AVERAGE_SPEED_KMH[canonical]);
  }

  if (AREA_AVERAGE_SPEED_KMH[rawArea]) {
    return Number(AREA_AVERAGE_SPEED_KMH[rawArea]);
  }

  if (canonical && /方面$/.test(canonical)) {
    if (/取手|藤代|守谷|牛久/.test(canonical)) return 39;
    if (/吉川|三郷|八潮|草加|越谷/.test(canonical)) return 34;
    if (/市川|船橋|鎌ヶ谷|鎌ケ谷|千葉/.test(canonical)) return 31;
    if (/葛飾|足立|江戸川|墨田|江東|荒川|台東/.test(canonical)) return 27;
    if (/我孫子/.test(canonical)) return 36;
    if (/柏|流山|野田|松戸/.test(canonical)) return 33;
  }

  const km = Number(distanceKm || 0);
  if (km <= 10) return 30;
  if (km <= 25) return 33;
  if (km <= 50) return 36;
  if (km <= 100) return 45;
  if (km <= 200) return 75;
  return 88;
}

function estimateTravelMinutesByAreaSpeed(distanceKm, areaInput) {
  const km = Math.max(0, Number(distanceKm || 0));
  const speed = getAreaAverageSpeedKmh(areaInput, km);
  if (!speed) return 0;
  return Math.round((km / speed) * 60);
}

function estimateFallbackTravelMinutes(distanceKm, areaInput = "") {
  return estimateTravelMinutesByAreaSpeed(distanceKm, areaInput);
}

function getRowOneWayTravelMinutes(row, fallbackDistanceKm = null, fallbackArea = "") {
  const rowStored = getStoredTravelMinutes(row?.travel_minutes ?? row?.travelMinutes);
  if (rowStored > 0) return rowStored;

  const stored = getCastTravelMinutesValue(row?.casts || row);
  if (stored > 0) return stored;

  const km = Number(fallbackDistanceKm ?? row?.distance_km ?? row?.casts?.distance_km ?? 0);
  const area = fallbackArea || normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "");
  return estimateFallbackTravelMinutes(km, area);
}

function getRowsOutboundMinutes(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) return 0;
  const representativeArea = getRepresentativeAreaFromRows(ordered);
  const routeDistanceKm = Number(calculateRouteDistanceGlobal(ordered) || 0);
  const lastRow = ordered[ordered.length - 1] || {};
  const storedLast = getCastTravelMinutesValue(lastRow?.casts || lastRow);
  if (storedLast > 0) return storedLast;
  return estimateFallbackTravelMinutes(routeDistanceKm, representativeArea);
}

function getRowsReturnMinutes(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) return 0;
  const lastRow = ordered[ordered.length - 1] || {};
  const area = normalizeAreaLabel(lastRow.destination_area || lastRow.casts?.area || getRepresentativeAreaFromRows(ordered) || "");
  return getRowOneWayTravelMinutes(lastRow, lastRow.distance_km || 0, area);
}

function getRowsTravelTimeSummary(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) {
    return { outboundMinutes: 0, returnMinutes: 0, totalMinutes: 0, sendOnlyMinutes: 0, stopCount: 0 };
  }
  const outboundMinutes = Math.round(getRowsOutboundMinutes(ordered));
  const returnMinutes = Math.round(getRowsReturnMinutes(ordered));
  const stopCount = ordered.length;
  const sendOnlyMinutes = Math.round(outboundMinutes + stopCount);
  const totalMinutes = Math.round(outboundMinutes + returnMinutes + stopCount);
  return { outboundMinutes, returnMinutes, totalMinutes, sendOnlyMinutes, stopCount };
}

function formatClockTimeFromMinutesGlobal(totalMinutes) {
  const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
  const h = Math.floor(safe / 60) % 24;
  const m = safe % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function getDistanceZoneInfoGlobal(distanceKm, areaInput = "") {
  const km = Number(distanceKm || 0);
  const speedKmh = getAreaAverageSpeedKmh(areaInput, km);
  const canonical = getCanonicalArea(normalizeAreaSpeedLookupInput(areaInput));
  return {
    key: canonical || (km <= 10 ? "short" : km <= 25 ? "middle" : "long"),
    label: canonical || "-",
    speedKmh
  };
}function calculateRouteDistanceGlobal(items) {
  if (!items || !items.length) return 0;
  let total = 0;
  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;

  items.forEach(item => {
    const point = getItemLatLng(item);
    if (point) {
      total += estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng);
      currentLat = point.lat;
      currentLng = point.lng;
    } else {
      total += Number(item.distance_km || 0);
    }
  });

  return Number(total.toFixed(1));
}

function calcVehicleRotationForecastGlobal(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) {
    return {
      routeDistanceKm: 0,
      returnDistanceKm: 0,
      zoneLabel: "-",
      predictedReturnTime: "-",
      predictedReadyTime: "-",
      predictedReturnMinutes: 0,
      extraSharedDelayMinutes: 0,
      stopCount: 0,
      returnAfterLabel: "-"
    };
  }

  const firstHour = rows.reduce((min, row) => {
    const val = Number(row.actual_hour ?? row.plan_hour ?? 0);
    return Number.isFinite(val) ? Math.min(min, val) : min;
  }, 99);

  const baseHour = firstHour === 99 ? 0 : firstHour;
  const routeDistanceKm = Number(calculateRouteDistanceGlobal(rows) || 0);
  const lastRow = rows[rows.length - 1] || {};
  const returnDistanceKm = Number(lastRow.distance_km || 0);
    const representativeArea = getRepresentativeAreaFromRows(rows);
  const returnArea = normalizeAreaLabel(lastRow.destination_area || lastRow.casts?.area || representativeArea || "");
  const primaryZone = getDistanceZoneInfoGlobal(Math.max(routeDistanceKm, returnDistanceKm), representativeArea);
  const timeSummary = getRowsTravelTimeSummary(rows);

  let departDelayMinutes = 20;
  if (baseHour === 3) departDelayMinutes = 18;
  else if (baseHour === 4) departDelayMinutes = 12;
  else if (baseHour >= 5) departDelayMinutes = 8;

  const outboundMinutes = timeSummary.outboundMinutes;
  const returnMinutes = timeSummary.returnMinutes;
  const dropoffMinutes = rows.length * 1;

  const baseStartMinutes = Number.isFinite(lastAutoDispatchRunAtMinutes) && lastAutoDispatchRunAtMinutes !== null
    ? lastAutoDispatchRunAtMinutes
    : (baseHour * 60 + departDelayMinutes);

  const predictedReturnMinutes = Math.round(outboundMinutes + dropoffMinutes + returnMinutes);
  const predictedReturnAbs = baseStartMinutes + predictedReturnMinutes;
  const predictedReadyAbs = predictedReturnAbs + 1;

  let extraSharedDelayMinutes = 0;
  if (rows.length >= 2) {
    const firstOnly = [rows[0]];
    const singleRouteDistanceKm = Number(calculateRouteDistanceGlobal(firstOnly) || rows[0].distance_km || 0);
    const singleReturnDistanceKm = Number(rows[0].distance_km || 0);
    const singleArea = normalizeAreaLabel(rows[0]?.destination_area || rows[0]?.casts?.area || representativeArea || "");
    const singleOutbound = getRowOneWayTravelMinutes(rows[0], singleRouteDistanceKm, singleArea);
    const singleReturn = getRowOneWayTravelMinutes(rows[0], singleReturnDistanceKm, singleArea);
    const singlePredictedReturnMinutes = Math.round(singleOutbound + 1 + singleReturn);
    extraSharedDelayMinutes = Math.max(0, predictedReturnMinutes - singlePredictedReturnMinutes);
  }

  return {
    routeDistanceKm,
    returnDistanceKm,
    zoneLabel: primaryZone.label,
    predictedReturnTime: formatClockTimeFromMinutesGlobal(predictedReturnAbs),
    predictedReadyTime: formatClockTimeFromMinutesGlobal(predictedReadyAbs),
    predictedReturnMinutes,
    extraSharedDelayMinutes: Math.round(extraSharedDelayMinutes),
    stopCount: rows.length,
    returnAfterLabel: `${predictedReturnMinutes}分後`
  };
}

async function applyAutoDispatchAssignments(assignments) {
  const groupedOrderMap = new Map();

  for (const assignment of Array.isArray(assignments) ? assignments : []) {
    const safeItemId = normalizeDispatchEntityId(assignment?.item_id);
    const safeVehicleId = normalizeDispatchEntityId(resolveVehicleCloudRowId(assignment?.vehicle_id));
    if (!safeItemId) {
      throw new Error("auto dispatch target id is missing");
    }
    if (!safeVehicleId) {
      throw new Error(`auto dispatch vehicle id is unresolved: ${assignment?.vehicle_id ?? ""}`);
    }

    const safeHour = Number(assignment?.actual_hour ?? 0);
    const key = `${safeVehicleId}_${safeHour}`;
    const nextOrder = (groupedOrderMap.get(key) || 0) + 1;
    groupedOrderMap.set(key, nextOrder);

    const { error } = await supabaseClient
      .from(getDispatchUnifiedTableName())
      .update({
        vehicle_id: safeVehicleId,
        driver_name: assignment?.driver_name || "",
        stop_order: nextOrder,
        status: "pending"
      })
      .eq("id", safeItemId);

    if (error) {
      console.error("applyAutoDispatchAssignments failed:", { assignment, safeItemId, safeVehicleId, error });
      throw error;
    }
  }
}


function __getDisplayGroupAreaLabel(item) {
  const rawGroup =
    item?.display_group ||
    item?.area_group ||
    item?.group_area ||
    item?.actual_group ||
    item?.group ||
    item?.destination_area ||
    item?.cluster_area ||
    item?.planned_area ||
    item?.casts?.area ||
    "無し";
  const area = normalizeAreaLabel(rawGroup);
  if (typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(area)) return area;
  return getAreaDisplayGroup(area) || area || "東京方面";
}

function __rowDistanceForCapacitySplit(row) {
  return Number(
    row?.distance_km ??
    row?.casts?.distance_km ??
    row?.distanceKm ??
    0
  ) || 0;
}

function __rowTravelMinutesForCapacitySplit(row) {
  return Number(
    row?.travel_minutes ??
    row?.casts?.travel_minutes ??
    row?.travelMinutes ??
    0
  ) || 0;
}

function __sortRowsForGroupCapacitySplit(rows) {
  const safeRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
  return [...safeRows].sort((a, b) => {
    const dist = __rowDistanceForCapacitySplit(b) - __rowDistanceForCapacitySplit(a);
    if (dist !== 0) return dist;
    const tm = __rowTravelMinutesForCapacitySplit(b) - __rowTravelMinutesForCapacitySplit(a);
    if (tm !== 0) return tm;
    return Number(a?.id || 0) - Number(b?.id || 0);
  });
}

function __hasEnoughVehiclesForDisplayGroups(items, vehicles) {
  if (!ENABLE_DISPLAY_GROUP_FORCE_BRANCH) return false;
  const safeItems = Array.isArray(items) ? items.filter(Boolean) : [];
  const safeVehicles = Array.isArray(vehicles) ? vehicles.filter(Boolean) : [];
  if (!safeItems.length || !safeVehicles.length) return false;

  const hourMap = new Map();
  safeItems.forEach(item => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    const group = __getDisplayGroupAreaLabel(item);
    if (!hourMap.has(hour)) hourMap.set(hour, new Set());
    hourMap.get(hour).add(group);
  });

  for (const groups of hourMap.values()) {
    if (groups.size > safeVehicles.length) return false;
  }
  return true;
}

function __getGroupFirstVehicleScore(vehicle, rows, monthlyMap) {
  const primary = rows[0] || {};
  const area = normalizeAreaLabel(
    primary?.destination_area ||
    primary?.cluster_area ||
    primary?.planned_area ||
    primary?.casts?.area ||
    "無し"
  );
  const month = monthlyMap?.get(Number(vehicle?.id)) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
  let score = 0;
  score += getVehicleAreaMatchScore(vehicle, area) * 2.6;
  score += getStrictHomeCompatibilityScore(area, vehicle?.home_area || "") * 1.9;
  score += Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || "")) * 0.8;
  score += getAreaAffinityScore(area, vehicle?.home_area || "") * 0.7;
  score -= Number(month.totalDistance || 0) * 0.02;
  score -= Number(month.avgDistance || 0) * 0.11;
  return score;
}

function __buildAssignmentsPreserveDisplayGroups(items, vehicles, monthlyMap) {
  const safeItems = Array.isArray(items) ? items.filter(Boolean) : [];
  const safeVehicles = Array.isArray(vehicles) ? vehicles.filter(Boolean) : [];
  if (!safeItems.length || !safeVehicles.length) return [];

  const assignments = [];
  const byHour = new Map();

  safeItems.forEach(item => {
    const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
    if (!byHour.has(hour)) byHour.set(hour, []);
    byHour.get(hour).push(item);
  });

  const hours = [...byHour.keys()].sort((a, b) => a - b);

  for (const hour of hours) {
    const hourItems = byHour.get(hour) || [];
    const byGroup = new Map();

    hourItems.forEach(item => {
      const group = __getDisplayGroupAreaLabel(item);
      if (!byGroup.has(group)) byGroup.set(group, []);
      byGroup.get(group).push(item);
    });

    const groupEntries = [...byGroup.entries()].sort((a, b) => {
      if (b[1].length !== a[1].length) return b[1].length - a[1].length;
      const aMax = Math.max(...a[1].map(row => Number(row?.distance_km || 0)), 0);
      const bMax = Math.max(...b[1].map(row => Number(row?.distance_km || 0)), 0);
      if (bMax !== aMax) return bMax - aMax;
      return String(a[0] || "").localeCompare(String(b[0] || ""), "ja");
    });

    const vehicleState = safeVehicles.map(vehicle => ({
      vehicle,
      id: Number(vehicle.id),
      capacity: Math.max(1, Number(vehicle.seat_capacity || 4)),
      count: 0,
      group: "",
      used: false
    }));

    for (const [group, rows] of groupEntries) {
      let remaining = __sortRowsForGroupCapacitySplit(rows);

      while (remaining.length) {
        const candidates = vehicleState
          .filter(state => state.count < state.capacity)
          .filter(state => !state.group || state.group === group)
          .sort((a, b) => {
            const aGroupBoost = a.group === group ? 5000 : (a.used ? 0 : 1800);
            const bGroupBoost = b.group === group ? 5000 : (b.used ? 0 : 1800);
            const aScore = __getGroupFirstVehicleScore(a.vehicle, remaining, monthlyMap) + aGroupBoost - a.count * 30;
            const bScore = __getGroupFirstVehicleScore(b.vehicle, remaining, monthlyMap) + bGroupBoost - b.count * 30;
            return bScore - aScore;
          });

        const picked = candidates[0];
        if (!picked) return [];

        picked.group = group;
        picked.used = true;

        const freeSeats = Math.max(0, picked.capacity - picked.count);
        const chunk = __sortRowsForGroupCapacitySplit(remaining.splice(0, freeSeats));
        let stopOrder = picked.count + 1;

        for (const row of chunk) {
          assignments.push({
            item_id: Number(row?.id || 0),
            actual_hour: hour,
            vehicle_id: picked.id,
            driver_name: picked.vehicle?.driver_name || "",
            distance_km: __rowDistanceForCapacitySplit(row),
            stop_order: stopOrder
          });
          stopOrder += 1;
        }

        picked.count += chunk.length;
      }
    }
  }

  const perVehicleHour = new Map();
  assignments.forEach(a => {
    const key = `${Number(a.vehicle_id)}__${Number(a.actual_hour)}`;
    if (!perVehicleHour.has(key)) perVehicleHour.set(key, []);
    perVehicleHour.get(key).push(a);
  });
  for (const rows of perVehicleHour.values()) {
    rows.sort((a, b) => Number(a.stop_order || 0) - Number(b.stop_order || 0) || Number(a.item_id || 0) - Number(b.item_id || 0));
    rows.forEach((a, idx) => {
      a.stop_order = idx + 1;
    });
  }

  return assignments;
}



async function assignUnassignedActualsForToday() {
  const selectedVehicles = getSelectedVehiclesForToday();
  if (!selectedVehicles.length) return [];

  const unassignedItems = currentActualsCache.filter(item => {
    const status = normalizeStatus(item.status);
    if (status === "cancel" || status === "done") return false;
    if (Number(item.vehicle_id || 0) > 0) return false;
    return true;
  });

  if (!unassignedItems.length) return [];

  const prepared = prepareActualRowsForDispatchCore(unassignedItems);
  if (!prepared.rows.length) return [];

  let assignments = [];
  try {
    const vehicleBridge = createDispatchVehicleBridge(selectedVehicles);
    assignments = callDispatchCoreSafe(prepared.rows, selectedVehicles, buildMonthlyDistanceMapForCurrentMonth(), {});
    assignments = remapAutoDispatchAssignments(assignments, prepared.sourceIdByTempId, vehicleBridge);
    console.groupCollapsed("[DispatchBridge][ASSIGN_UNASSIGNED]");
    console.log("bridge", vehicleBridge?.snapshot?.());
    console.table((Array.isArray(assignments) ? assignments : []).map(a => ({ coreVehicleId: a?.__core_vehicle_id, resolvedLocalVehicleId: a?.__resolved_local_vehicle_id, savedVehicleId: a?.vehicle_id, itemId: a?.item_id, driver: a?.driver_name })));
    console.groupEnd();
  } catch (error) {
    console.error("assignUnassignedActualsForToday dispatchCore error:", error);
    assignments = [];
  }

  if (!Array.isArray(assignments) || !assignments.length) return [];
  await applyAutoDispatchAssignments(assignments);
  return assignments;
}


/* 旧自動配車の緊急割当ロジックは削除しました。 */


/* 旧自動配車のoverflow再配分ロジックは削除しました。 */

async function runAutoDispatch() {
  const selectedVehicles = Array.isArray(getSelectedVehiclesForToday())
    ? getSelectedVehiclesForToday().filter(Boolean)
    : [];
  if (!selectedVehicles.length) {
    alert("可能車両を選択してください");
    return [];
  }

  const activeItems = Array.isArray(getActiveDispatchItemsForAutoAssign())
    ? getActiveDispatchItemsForAutoAssign().filter(Boolean)
    : [];
  if (!activeItems.length) {
    alert("自動配車対象のActualがありません");
    return [];
  }

  const prepared = prepareActualRowsForDispatchCore(activeItems);
  if (!prepared.rows.length) {
    alert("自動配車対象のActualがありません");
    return [];
  }

  const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
  let assignments = [];
  try {
    const vehicleBridge = createDispatchVehicleBridge(selectedVehicles);
    assignments = callDispatchCoreSafe(prepared.rows, selectedVehicles, monthlyMap, {});
    assignments = remapAutoDispatchAssignments(assignments, prepared.sourceIdByTempId, vehicleBridge);
    console.groupCollapsed("[DispatchBridge][RUN_AUTO]");
    console.log("bridge", vehicleBridge?.snapshot?.());
    console.table((Array.isArray(assignments) ? assignments : []).map(a => ({ coreVehicleId: a?.__core_vehicle_id, resolvedLocalVehicleId: a?.__resolved_local_vehicle_id, savedVehicleId: a?.vehicle_id, itemId: a?.item_id, driver: a?.driver_name })));
    console.groupEnd();
  } catch (error) {
    console.error("runAutoDispatch optimize error:", error);
    alert(`自動配車エラー: ${error.message}`);
    return [];
  }

  if (!Array.isArray(assignments) || !assignments.length) {
    alert("配車結果がありません");
    return [];
  }

  try {
    await applyAutoDispatchAssignments(assignments);
  } catch (error) {
    console.error(error);
    alert(`配車更新エラー: ${error.message}`);
    return [];
  }

  await addHistory(currentDispatchId, null, "auto_dispatch", "自動配車を実行");
  await loadActualsByDate(els.actualDate?.value || todayStr());
  await loadPlansByDate(els.planDate?.value || todayStr());
  renderDailyDispatchResult();
  scrollToDispatchResult();
  return assignments;
}

function scrollToDispatchResult() {
  try {
    const target = els.dailyDispatchResult || document.getElementById("dailyDispatchResult");
    if (!target) return;
    window.requestAnimationFrame(() => {
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  } catch (error) {
    console.warn("scrollToDispatchResult error:", error);
  }
}

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
          item => sameVehicleAssignmentId(item?.vehicle_id, vehicle?.id)
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

function parseJaDurationMinutes(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return null;
  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const n = Math.round(Number(raw));
    return Number.isFinite(n) && n > 0 ? n : null;
  }
  const compact = raw.replace(/\s+/g, '');
  const h = compact.match(/(\d+)時間/);
  const m = compact.match(/(\d+)分/);
  const hours = h ? Number(h[1]) : 0;
  const minutes = m ? Number(m[1]) : 0;
  const total = hours * 60 + minutes;
  return total > 0 ? total : null;
}

function formatCastTravelMinutesDisplay(value) {
  const normalized = normalizeCastMetricTravelMinutesValue(value);
  return normalized == null ? '' : formatMinutesAsJa(normalized);
}

function normalizeCastDirectionAddressSource(address = "") {
  return String(address || "")
    .replace(/^\s*日本[、,]?\s*/, "")
    .replace(/^\s*〒\d{3}-?\d{4}\s*/, "")
    .trim();
}

function normalizeCastDirectionLocalityToken(value = "") {
  return String(value || "")
    .replace(/[0-9０-９].*$/, "")
    .replace(/[\-－ー].*$/, "")
    .replace(/(?:丁目|番地|番|号).*$/, "")
    .replace(/[\s、,，]+.*$/, "")
    .trim();
}

function buildCastDirectionDisplayLabelFromAddress(address = "", fallbackLabel = "") {
  const normalizedAddress = normalizeCastDirectionAddressSource(address);
  if (!normalizedAddress) return normalizeAreaLabel(fallbackLabel || "");

  const body = normalizedAddress
    .replace(/^(東京都|北海道|京都府|大阪府|[^都道府県\s]+県)\s*/, "")
    .trim();

  const wardMatch = body.match(/^(?:[^0-9０-９\s、,，]+市)?([^0-9０-９\s、,，]+区)([^0-9０-９\s、,，\-－ー丁目番地号]+)/);
  if (wardMatch) {
    const ward = String(wardMatch[1] || "").trim();
    const locality = normalizeCastDirectionLocalityToken(wardMatch[2] || "");
    if (ward && locality) return `${ward}${locality}方面`;
    if (ward) return `${ward}方面`;
  }

  const countyMatch = body.match(/^([^0-9０-９\s、,，]+郡)([^0-9０-９\s、,，]+(?:町|村))/);
  if (countyMatch) {
    const county = String(countyMatch[1] || "").trim();
    const municipality = String(countyMatch[2] || "").trim();
    const locality = normalizeCastDirectionLocalityToken(municipality).replace(/[町村]$/, "");
    if (county && locality) return `${county}${locality}方面`;
    if (county && municipality) return `${county}${municipality}方面`;
    if (county) return `${county}方面`;
  }

  const cityMatch = body.match(/^([^0-9０-９\s、,，]+市)([^0-9０-９\s、,，\-－ー丁目番地号]+)/);
  if (cityMatch) {
    const city = String(cityMatch[1] || "").trim();
    const locality = normalizeCastDirectionLocalityToken(cityMatch[2] || "");
    if (city && locality) return `${city}${locality}方面`;
    if (city) return `${city}方面`;
  }

  return normalizeAreaLabel(fallbackLabel || "");
}

function getCastDirectionDisplayLabel(cast, addressOverride = "") {
  const address = String(addressOverride || cast?.address || "").trim();
  const fallbackLabel = normalizeAreaLabel(cast?.area || "");
  const addressLabel = buildCastDirectionDisplayLabelFromAddress(address, "");
  return normalizeAreaLabel(addressLabel || fallbackLabel);
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
          sameVehicleAssignmentId(item?.vehicle_id, numericVehicleId) &&
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
  const sendOnlyMinutes = Math.round(getRowsTravelTimeSummary(orderedRows).sendOnlyMinutes);
  const displayTotalKm = (!summary.hasFixedReport && isLastTripDriver)
    ? Number(summary.sendKm || 0)
    : Number(summary.totalKm || 0);
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

function getVehicleLiveDailySummary(vehicle, orderedRows) {
  const vehicleId = Number(vehicle?.id || 0);
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  const actualRows = Array.isArray(currentActualsCache)
    ? currentActualsCache.filter(
        item =>
          sameVehicleAssignmentId(item?.vehicle_id, vehicleId) &&
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
  const isLastTripDriver = isDriverLastTripChecked(vehicleId);
  const totalKm = Number((sendKm + returnKm).toFixed(1));
  const travelSummary = getRowsTravelTimeSummary(baseRows);
  const driveMinutes = isLastTripDriver
    ? Math.round(Number(travelSummary.sendOnlyMinutes || 0))
    : Math.round(Number(travelSummary.totalMinutes || 0));

  return {
    sendKm: Number(sendKm.toFixed(1)),
    returnKm: Number(returnKm.toFixed(1)),
    totalKm: isLastTripDriver ? Number(sendKm.toFixed(1)) : totalKm,
    driveMinutes,
    jobCount: baseRows.length,
    hasFixedReport: false
  };
}

function getVehicleProjectedMonthlyDistance(vehicleId, monthlyMap, orderedRows) {
  const currentMonth = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0 };
  const todaySummary = getVehicleDailySummary({ id: vehicleId }, orderedRows);
  if (todaySummary.hasFixedReport) {
    return Number(Number(currentMonth.totalDistance || 0).toFixed(1));
  }
  return Number(Number(currentMonth.totalDistance || 0) + Number(todaySummary.totalKm || 0));
}function buildVehicleLineLabel(vehicle) {
  return "";
}


let dispatchOverviewMap = null;
let dispatchOverviewLayer = null;
let dispatchOverviewHost = null;
let lastDispatchOverviewCards = [];
const dispatchOverviewFilterState = new Map();

function getDispatchOverviewVehicleColor(index) {
  const colors = ["#3b82f6", "#22c55e", "#ef4444", "#f59e0b", "#a855f7", "#06b6d4", "#84cc16", "#ec4899"];
  return colors[Math.abs(Number(index) || 0) % colors.length];
}

function getVehicleDisplayName(vehicle) {
  return String(vehicle?.driver_name || vehicle?.plate_number || `車両${vehicle?.id || ""}` || "車両").trim() || "車両";
}

function getDispatchRowLatLng(row) {
  const lat = Number(row?.casts?.latitude ?? row?.latitude ?? row?.lat);
  const lng = Number(row?.casts?.longitude ?? row?.longitude ?? row?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  return { lat, lng };
}

function buildVehicleRouteMapUrl(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) return "";
  const origin = `${ORIGIN_LAT},${ORIGIN_LNG}`;
  const points = rows
    .map(row => {
      const p = getDispatchRowLatLng(row);
      if (p) return `${p.lat},${p.lng}`;
      const address = String(row?.destination_address || row?.casts?.address || "").trim();
      return address || "";
    })
    .filter(Boolean);
  if (!points.length) return "";
  const destination = points[points.length - 1];
  const waypoints = points.slice(0, -1);
  const params = new URLSearchParams({
    api: "1",
    origin,
    destination,
    travelmode: "driving"
  });
  if (waypoints.length) params.set("waypoints", waypoints.join("|"));
  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

function renderDispatchOverviewLegend(cards) {
  const host = document.getElementById("dispatchOverviewLegend");
  if (!host) return;
  host.innerHTML = "";
  (cards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    const isOn = dispatchOverviewFilterState.get(vehicleId) !== false;
    const color = getDispatchOverviewVehicleColor(index);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn ghost";
    button.style.cssText = `display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;border:1px solid ${isOn ? color : 'rgba(148,163,184,.35)'};background:${isOn ? 'rgba(15,23,42,.85)' : 'rgba(15,23,42,.45)'};color:${isOn ? '#fff' : '#94a3b8'};cursor:pointer;`;
    button.innerHTML = `<span style="display:inline-block;width:12px;height:12px;border-radius:999px;background:${color};box-shadow:0 0 0 2px rgba(255,255,255,.18) inset;"></span><span>${escapeHtml(getVehicleDisplayName(vehicle))}</span><span style="font-size:11px;opacity:.8;">${isOn ? '表示中' : '非表示'}</span>`;
    button.addEventListener("click", () => {
      dispatchOverviewFilterState.set(vehicleId, !isOn);
      renderDispatchOverviewLegend(lastDispatchOverviewCards);
      renderDispatchOverviewMap(lastDispatchOverviewCards);
    });
    host.appendChild(button);
  });
}

function renderDispatchOverviewMap(cards) {
  const host = document.getElementById("dispatchOverviewMap");
  lastDispatchOverviewCards = Array.isArray(cards) ? cards : [];
  renderDispatchOverviewLegend(lastDispatchOverviewCards);
  if (!host || !window.L) return;

  if (dispatchOverviewMap && dispatchOverviewHost !== host) {
    try { dispatchOverviewMap.remove(); } catch (_) {}
    dispatchOverviewMap = null;
    dispatchOverviewLayer = null;
  }

  dispatchOverviewHost = host;

  if (!dispatchOverviewMap) {
    host.innerHTML = "";
    dispatchOverviewMap = window.L.map(host, { preferCanvas: true, zoomControl: true });
    window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(dispatchOverviewMap);
  }

  if (dispatchOverviewLayer) {
    try { dispatchOverviewLayer.remove(); } catch (_) {}
  }
  dispatchOverviewLayer = window.L.layerGroup().addTo(dispatchOverviewMap);

  const bounds = [];
  const originMarker = window.L.circleMarker([ORIGIN_LAT, ORIGIN_LNG], {
    radius: 9,
    color: "#111827",
    fillColor: "#111827",
    fillOpacity: 1,
    weight: 2
  }).bindPopup(`<b>${escapeHtml(ORIGIN_LABEL || '起点')}</b>`);
  originMarker.addTo(dispatchOverviewLayer);
  bounds.push([ORIGIN_LAT, ORIGIN_LNG]);

  (lastDispatchOverviewCards || []).forEach(({ vehicle, orderedRows }, index) => {
    if (!orderedRows?.length) return;
    const vehicleId = Number(vehicle?.id || 0);
    if (!dispatchOverviewFilterState.has(vehicleId)) dispatchOverviewFilterState.set(vehicleId, true);
    if (dispatchOverviewFilterState.get(vehicleId) === false) return;
    const color = getDispatchOverviewVehicleColor(index);
    const routePoints = [[ORIGIN_LAT, ORIGIN_LNG]];
    const vehicleName = getVehicleDisplayName(vehicle);

    orderedRows.forEach((row, rowIndex) => {
      const point = getDispatchRowLatLng(row);
      if (!point) return;
      routePoints.push([point.lat, point.lng]);
      const pinNo = rowIndex + 1;
      const marker = window.L.marker([point.lat, point.lng], {
        icon: window.L.divIcon({
          className: 'dispatch-overview-pin',
          html: `<div style="width:26px;height:26px;border-radius:999px;background:${color};border:2px solid rgba(255,255,255,.96);box-shadow:0 6px 16px rgba(15,23,42,.38);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:12px;line-height:1;">${pinNo}</div>`,
          iconSize: [26, 26],
          iconAnchor: [13, 13],
          popupAnchor: [0, -14]
        })
      });
      const castName = String(row?.casts?.name || "-");
      const area = normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || "-");
      const distanceKm = Number(row?.distance_km || 0).toFixed(1);
      marker.bindPopup(`<div style="min-width:180px;line-height:1.7;"><b>${escapeHtml(vehicleName)}-${pinNo}</b><br>送り先: ${escapeHtml(castName)}<br>方面: ${escapeHtml(area)}<br>距離: ${escapeHtml(distanceKm)}km</div>`);
      marker.bindTooltip(`${escapeHtml(vehicleName)}-${pinNo}`, { direction: "top", opacity: 0.95 });
      marker.addTo(dispatchOverviewLayer);
      bounds.push([point.lat, point.lng]);
    });

    if (routePoints.length >= 2) {
      const polyline = window.L.polyline(routePoints, {
        color,
        weight: 4,
        opacity: 0.78,
        lineCap: 'round',
        lineJoin: 'round'
      });
      polyline.bindTooltip(`${escapeHtml(vehicleName)} ルート`, { sticky: true, opacity: 0.92 });
      polyline.addTo(dispatchOverviewLayer);
    }
  });

  window.setTimeout(() => {
    try {
      dispatchOverviewMap.invalidateSize();
      if (bounds.length >= 2) {
        dispatchOverviewMap.fitBounds(bounds, { padding: [24, 24] });
      } else if (bounds.length === 1) {
        dispatchOverviewMap.setView(bounds[0], 13);
      } else {
        dispatchOverviewMap.setView([ORIGIN_LAT, ORIGIN_LNG], 11);
      }
    } catch (error) {
      console.error("dispatchOverviewMap fit error:", error);
    }
  }, 50);
}

function buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap) {
  return vehicles.map(vehicle => {
    const rows = activeItems
      .filter(item => sameVehicleAssignmentId(item?.vehicle_id, vehicle?.id))
      .sort((a, b) => {
        const ah = Number(a.actual_hour ?? 0);
        const bh = Number(b.actual_hour ?? 0);
        if (ah !== bh) return ah - bh;

        const ao = Number(a.stop_order || 0);
        const bo = Number(b.stop_order || 0);
        if (ao !== bo) return ao - bo;

        return Number(a.id || 0) - Number(b.id || 0);
      });

    const orderedRows = [...rows];
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

  const pinLinks = rows
    .map((row, index) => {
      const pinUrl = buildDispatchItemMapUrl(row);
      if (!pinUrl) return "";
      const castName = String(row?.casts?.name || `送り先${index + 1}`).trim() || `送り先${index + 1}`;
      return `${index + 1}️⃣ ${castName} 📍\n${pinUrl}`;
    })
    .filter(Boolean);

  if (pinLinks.length) {
    footer.push("📍送り先ピン");
    footer.push(...pinLinks);
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
    .join("\n\n\n")
    .trim();
}

function getOvernightLooseHourBucket(item) {
  const hour = Number(item?.actual_hour ?? item?.plan_hour ?? 0);
  return hour >= 0 && hour <= 5 ? "overnight" : String(hour);
}

function getAssignmentItemHourGroupKey(item) {
  const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
  const canonical = getCanonicalArea(area) || area;
  const group = getAreaDisplayGroup(area) || canonical || area;
  const hourBucket = getOvernightLooseHourBucket(item);
  return `${hourBucket}__${group}__${canonical}`;
}

function getSoftBridgeAreaScore(baseArea, compareArea) {
  const base = getCanonicalArea(normalizeAreaLabel(baseArea)) || normalizeAreaLabel(baseArea);
  const compare = getCanonicalArea(normalizeAreaLabel(compareArea)) || normalizeAreaLabel(compareArea);
  if (!base || !compare) return 0;
  if (base === compare) return 140;

  const northeast = new Set(["我孫子方面", "取手方面", "藤代方面", "守谷方面", "牛久方面"]);
  const eastUrban = new Set(["葛飾方面", "足立方面", "墨田方面", "荒川方面", "江戸川方面", "市川方面"]);
  const downtownBridge = new Set(["墨田方面", "荒川方面"]);

  if ((base === "我孫子方面" && eastUrban.has(compare)) || (compare === "我孫子方面" && eastUrban.has(base))) return 108;
  if ((northeast.has(base) && eastUrban.has(compare)) || (northeast.has(compare) && eastUrban.has(base))) return 92;
  if ((base === "我孫子方面" && northeast.has(compare)) || (compare === "我孫子方面" && northeast.has(base))) return 132;
  if ((downtownBridge.has(base) && eastUrban.has(compare)) || (downtownBridge.has(compare) && eastUrban.has(base))) return 86;

  const affinity = getAreaAffinityScore(base, compare);
  const direction = getDirectionAffinityScore(base, compare);
  if (affinity >= 72 && direction >= 0) return 74;
  if (affinity >= 58 && direction >= 18) return 54;
  return 0;
}

function getRoundTripMinutesForItem(item) {
  const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
  const travel = getStoredTravelMinutes(item?.casts?.travel_minutes || item?.travel_minutes);
  const oneWay = travel > 0 ? travel : estimateTravelMinutesByDistance(Number(item?.distance_km || 0), area);
  return Math.round(oneWay * 2);
}

function countAssignmentsByVehicle(assignments) {
  const map = new Map();
  (assignments || []).forEach(a => {
    map.set(Number(a.vehicle_id), Number(map.get(Number(a.vehicle_id)) || 0) + 1);
  });
  return map;
}

function rebundleLongDistanceDirectionalClusters(assignments, items, vehicles, monthlyMap, options = {}) {
  const working = Array.isArray(assignments) ? assignments.map(a => ({ ...a })) : [];
  if (!working.length || !Array.isArray(items) || !items.length || !Array.isArray(vehicles) || !vehicles.length) return working;

  const itemMap = new Map(items.map(item => [Number(item.id), item]));
  const vehicleMap = new Map(vehicles.map(v => [Number(v.id), v]));
  const threshold = Number(options.roundTripThreshold || 55);

  const rebuildVehicleLoadMap = () => countAssignmentsByVehicle(working);

  const byKey = new Map();
  working.forEach(a => {
    const item = itemMap.get(Number(a.item_id));
    if (!item) return;
    const key = getAssignmentItemHourGroupKey(item);
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(a);
  });

  for (const [, clusterAssignments] of byKey.entries()) {
    if (!clusterAssignments || clusterAssignments.length < 2) continue;

    const itemsForCluster = clusterAssignments.map(a => itemMap.get(Number(a.item_id))).filter(Boolean);
    if (itemsForCluster.length < 2) continue;

    const minRoundTrip = Math.min(...itemsForCluster.map(getRoundTripMinutesForItem));
    if (minRoundTrip < threshold) continue;

    const canonical = getCanonicalArea(normalizeAreaLabel(itemsForCluster[0]?.destination_area || itemsForCluster[0]?.cluster_area || itemsForCluster[0]?.casts?.area || "無し"));
    const group = getAreaDisplayGroup(itemsForCluster[0]?.destination_area || itemsForCluster[0]?.cluster_area || itemsForCluster[0]?.casts?.area || "無し");

    let bestVehicleId = null;
    let bestScore = -Infinity;
    const vehicleLoadMap = rebuildVehicleLoadMap();

    for (const vehicle of vehicles) {
      const vehicleId = Number(vehicle.id);
      const homeArea = normalizeAreaLabel(vehicle?.home_area || "");
      const currentVehicleAssignments = working.filter(a => Number(a.vehicle_id) === vehicleId);
      const currentAreas = currentVehicleAssignments.map(a => {
        const item = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.casts?.area || "無し");
      }).filter(Boolean);

      if (currentAreas.length && currentAreas.some(area => hasHardReverseMix(area, [canonical || group]))) continue;
      if (isHardReverseForHome(canonical || group, homeArea)) continue;

      const month = monthlyMap?.get(vehicleId) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
      const existingCount = Number(vehicleLoadMap.get(vehicleId) || 0);
      const strict = getStrictHomeCompatibilityScore(canonical || group, homeArea);
      const direction = Math.max(0, getDirectionAffinityScore(canonical || group, homeArea));
      const affinity = currentAreas.length
        ? Math.max(...currentAreas.map(area => getAreaAffinityScore(canonical || group, area)))
        : getAreaAffinityScore(canonical || group, homeArea);
      const sameGroupBonus = currentAreas.length
        ? Math.max(...currentAreas.map(area => {
            const areaGroup = getAreaDisplayGroup(area);
            const areaCanonical = getCanonicalArea(area);
            return (areaCanonical && canonical && areaCanonical === canonical) ? 160 : (areaGroup === group ? 110 : 0);
          }))
        : 40;
      const currentCountInCluster = clusterAssignments.filter(a => Number(a.vehicle_id) === vehicleId).length;
      const totalClusterDistance = itemsForCluster.reduce((sum, item) => sum + Number(item?.distance_km || 0), 0);
      let score = 0;
      score += currentCountInCluster * 260;
      score += sameGroupBonus;
      score += strict * 1.3 + direction * 0.9 + affinity * 0.6;
      score -= Number(month.totalDistance || 0) * 0.02;
      score -= Number(month.avgDistance || 0) * 0.15;
      score -= existingCount * 24;
      score -= totalClusterDistance * 0.05;

      if (bestVehicleId == null || score > bestScore) {
        bestVehicleId = vehicleId;
        bestScore = score;
      }
    }

    if (bestVehicleId == null) continue;
    clusterAssignments.forEach(a => {
      const vehicle = vehicleMap.get(bestVehicleId);
      a.vehicle_id = bestVehicleId;
      a.driver_name = vehicle?.driver_name || "";
    });
  }

  const evaluateBridgeVehicle = (assignment, vehicle) => {
    const item = itemMap.get(Number(assignment.item_id));
    if (!item || !vehicle) return -Infinity;
    const targetArea = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
    const vehicleId = Number(vehicle.id);
    const existingAssignments = working.filter(a => Number(a.vehicle_id) === vehicleId && Number(a.item_id) !== Number(assignment.item_id));
    const existingAreas = existingAssignments
      .map(a => {
        const row = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || "無し");
      })
      .filter(Boolean);

    if (existingAreas.length && existingAreas.some(area => hasHardReverseMix(targetArea, [area]))) return -Infinity;
    if (isHardReverseForHome(targetArea, vehicle?.home_area || "")) return -Infinity;

    const month = monthlyMap?.get(vehicleId) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
    const existingCount = existingAssignments.length;
    const strict = getStrictHomeCompatibilityScore(targetArea, vehicle?.home_area || "");
    const direction = Math.max(0, getDirectionAffinityScore(targetArea, vehicle?.home_area || ""));
    const bestBridge = existingAreas.length
      ? Math.max(...existingAreas.map(area => getSoftBridgeAreaScore(targetArea, area)))
      : 0;
    const bestAffinity = existingAreas.length
      ? Math.max(...existingAreas.map(area => getAreaAffinityScore(targetArea, area)))
      : getAreaAffinityScore(targetArea, vehicle?.home_area || "");
    const bundleCount = existingAreas.filter(area => getSoftBridgeAreaScore(targetArea, area) >= 90).length;
    const isLooseOvernight = getOvernightLooseHourBucket(item) === "overnight";
    let score = 0;
    score += bestBridge * 3.6;
    score += bundleCount * 110;
    score += strict * 1.15 + direction * 0.7 + bestAffinity * 0.4;
    score -= Number(month.totalDistance || 0) * 0.018;
    score -= Number(month.avgDistance || 0) * 0.11;
    score -= existingCount * 14;
    if (isLooseOvernight) score += 44;
    if (!existingAreas.length) score -= 120;
    return score;
  };

  for (let pass = 0; pass < 2; pass += 1) {
    for (const assignment of [...working]) {
      const item = itemMap.get(Number(assignment.item_id));
      if (!item) continue;
      if (getOvernightLooseHourBucket(item) !== "overnight") continue;

      const targetArea = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || "無し");
      const currentVehicle = vehicleMap.get(Number(assignment.vehicle_id));
      const currentScore = evaluateBridgeVehicle(assignment, currentVehicle);
      let bestVehicle = currentVehicle;
      let bestScore = currentScore;

      for (const vehicle of vehicles) {
        const score = evaluateBridgeVehicle(assignment, vehicle);
        if (score > bestScore) {
          bestScore = score;
          bestVehicle = vehicle;
        }
      }

      if (!bestVehicle || Number(bestVehicle.id) === Number(assignment.vehicle_id)) continue;
      if (bestScore < currentScore + 42) continue;

      const destinationAssignments = working.filter(a => Number(a.vehicle_id) === Number(bestVehicle.id) && Number(a.item_id) !== Number(assignment.item_id));
      const destinationAreas = destinationAssignments.map(a => {
        const row = itemMap.get(Number(a.item_id));
        return normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || "無し");
      }).filter(Boolean);
      if (destinationAreas.length && destinationAreas.every(area => getSoftBridgeAreaScore(targetArea, area) < 70)) continue;

      assignment.vehicle_id = Number(bestVehicle.id);
      assignment.driver_name = bestVehicle?.driver_name || "";
    }
  }

  return working;
}


function getOverflowDisplayRowsFromCache() {
  const meta = window.__THEMIS_LAST_OVERFLOW__ || null;
  if (!meta || !Array.isArray(meta.overflowGroups) || !meta.overflowGroups.length) return [];

  const overflowIds = new Set(
    meta.overflowGroups.flatMap(group => Array.isArray(group?.itemIds) ? group.itemIds.map(id => Number(id || 0)) : []).filter(Boolean)
  );
  if (!overflowIds.size) return [];

  const groupById = new Map();
  meta.overflowGroups.forEach(group => {
    (Array.isArray(group?.itemIds) ? group.itemIds : []).forEach(id => {
      groupById.set(Number(id || 0), String(group?.group || '無し'));
    });
  });

  return currentActualsCache
    .filter(item => overflowIds.has(Number(item?.id || 0)))
    .map(item => ({
      ...item,
      overflow_group: groupById.get(Number(item?.id || 0)) || normalizeAreaLabel(item?.destination_area || item?.casts?.area || '無し')
    }))
    .sort((a, b) => Number(a?.actual_hour || 0) - Number(b?.actual_hour || 0) || Number(b?.distance_km || 0) - Number(a?.distance_km || 0));
}




function getCapacityOverflowRowsFromCache() {
  const meta = window.__THEMIS_LAST_OVERFLOW__ || null;
  const rows = Array.isArray(meta?.capacityOverflowItems) ? meta.capacityOverflowItems : [];
  if (!rows.length) return [];
  const byId = new Map(currentActualsCache.map(item => [Number(item?.id || 0), item]));
  return rows.map(row => {
    const item = byId.get(Number(row?.itemId || 0)) || null;
    return {
      ...(item || {}),
      actual_hour: Number(row?.hour || item?.actual_hour || 0),
      distance_km: Number(row?.distanceKm || item?.distance_km || item?.casts?.distance_km || 0),
      capacity_group: String(row?.group || row?.area || normalizeAreaLabel(item?.destination_area || item?.casts?.area || '無し')),
      capacity_reason: '定員超過'
    };
  }).sort((a, b) => Number(a?.actual_hour || 0) - Number(b?.actual_hour || 0) || Number(a?.distance_km || 0) - Number(b?.distance_km || 0));
}

function buildCapacityOverflowHtml() {
  const meta = window.__THEMIS_LAST_OVERFLOW__ || null;
  const rows = getCapacityOverflowRowsFromCache();
  const overflowCount = Number(meta?.capacityOverflowCount || 0);
  if (!overflowCount || !rows.length) return '';
  const totalSeatCapacity = Number(meta?.totalSeatCapacity || 0);
  const totalCastCount = Number(meta?.totalCastCount || 0);
  return `
    <div class="vehicle-result-card" style="border:1px solid rgba(239,68,68,.35); margin-bottom:16px;">
      <div class="vehicle-result-head">
        <div class="vehicle-result-title">
          <h4>定員オーバー</h4>
          <div class="vehicle-result-meta">方面確定後に、起点から近い順で超過分を未配車へ退避しています</div>
        </div>
        <div class="vehicle-result-badges">
          <span class="metric-badge">総定員 ${totalSeatCapacity}人</span>
          <span class="metric-badge">対象 ${totalCastCount}人</span>
          <span class="metric-badge">超過 ${overflowCount}人</span>
        </div>
      </div>
      <div class="vehicle-result-body">
        ${rows.map((row, index) => `
          <div class="dispatch-row">
            <div class="dispatch-left">
              <span class="badge-time">${escapeHtml(getHourLabel(row.actual_hour))}</span>
              <span class="badge-order">未配車 ${index + 1}</span>
              <span class="dispatch-name">${buildMapLinkHtml({
                name: row.casts?.name,
                address: row.destination_address || row.casts?.address,
                lat: row.casts?.latitude,
                lng: row.casts?.longitude,
                className: 'dispatch-name-link'
              })}</span>
              <span class="dispatch-area">${escapeHtml(row.capacity_group || normalizeAreaLabel(row.destination_area || '-'))}</span>
              <span class="badge-status canceled">定員超過</span>
            </div>
            <div class="dispatch-right">
              <div class="dispatch-distance">${Number(row.distance_km || row.casts?.distance_km || 0).toFixed(1)}km</div>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
}
function buildOverflowEvaluationHtml() {
  return "";

  /* hidden for operations
    <div class="vehicle-result-card" style="border:1px solid rgba(59,130,246,.28); margin-bottom:16px;">
      <div class="vehicle-result-head">
        <div class="vehicle-result-title">
          <h4>あぶれ仮投入評価</h4>
          <div class="vehicle-result-meta">判定順: 戻り時間 → 距離 → 人数</div>
        </div>
        <div class="vehicle-result-badges">
          <span class="metric-badge">対象 ${evaluations.length}件</span>
        </div>
      </div>
      <div class="vehicle-result-body">
        ${evaluations.map((evalRow, index) => {
          const candidates = Array.isArray(evalRow?.candidates) ? evalRow.candidates : [];
          const selectedVehicleId = Number(evalRow?.selectedVehicleId || 0);
          return `
            <div style="padding:12px 0; border-bottom:${index === evaluations.length - 1 ? 'none' : '1px solid rgba(148,163,184,.14)'};">
              <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-bottom:8px;">
                <span class="badge-time">${escapeHtml(getHourLabel(evalRow?.hour || 0))}</span>
                <span class="badge-order">あぶれ ${index + 1}</span>
                <span class="dispatch-name">${escapeHtml(evalRow?.castName || '-')}</span>
                <span class="dispatch-area">${escapeHtml(evalRow?.group || evalRow?.area || '-')}</span>
                <span class="metric-badge">距離 ${Number(evalRow?.distanceKm || 0).toFixed(1)}km</span>
              </div>
              <div style="display:grid; gap:8px;">
                ${candidates.map(candidate => {
                  const isSelected = Number(candidate?.vehicleId || 0) === selectedVehicleId;
                  if (candidate?.excluded) {
                    return `
                      <div style="padding:10px 12px; border-radius:12px; background:rgba(15,23,42,.35); border:1px solid rgba(148,163,184,.18); color:#94a3b8;">
                        <strong>${escapeHtml(candidate?.vehicleName || '-')}</strong>
                        / 除外: ${escapeHtml(candidate?.reason || '-')}
                      </div>
                    `;
                  }
                  return `
                    <div style="padding:10px 12px; border-radius:12px; background:${isSelected ? 'rgba(34,197,94,.10)' : 'rgba(15,23,42,.35)'}; border:1px solid ${isSelected ? 'rgba(34,197,94,.38)' : 'rgba(148,163,184,.18)'};">
                      <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-bottom:4px;">
                        <strong>${escapeHtml(candidate?.vehicleName || '-')}</strong>
                        ${isSelected ? '<span class="badge-status assigned">採用</span>' : ''}
                      </div>
                      <div class="dispatch-meta" style="font-size:12px; color:#9aa3b2; line-height:1.8;">
                        最大戻り ${escapeHtml(formatMinutesAsJa(candidate?.maxReturnMinutes || 0))}
                        / 全体合計往復 ${Number(candidate?.totalRoundTripKm || 0).toFixed(1)}km
                        / 人数 ${Number(candidate?.seatCountAfter || 0)}人
                        / 対象車戻り ${escapeHtml(formatMinutesAsJa(candidate?.routeMinutes || 0))} (往路 ${escapeHtml(formatMinutesAsJa(candidate?.outboundMinutes || 0))} / 復路 ${escapeHtml(formatMinutesAsJa(candidate?.returnMinutes || 0))})
                        / 対象車往復 ${Number(candidate?.routeKm || 0).toFixed(1)}km (往路 ${Number(candidate?.outboundKm || 0).toFixed(1)} / 復路 ${Number(candidate?.returnKm || 0).toFixed(1)})
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>
            </div>
          `;
        }).join('')}
      </div>
    </div>
  `;
  */
}

function renderDailyDispatchResult() {
  if (!els.dailyDispatchResult) return;

  const vehicles = getSelectedVehiclesForToday();
  if (!vehicles.length) {
    els.dailyDispatchResult.innerHTML = `<div class="muted">使用車両が未選択です</div>`;
    renderOperationAndSimulationUI();
    return;
  }

  const activeItems = currentActualsCache.filter(
    x => normalizeStatus(x.status) !== "done" && normalizeStatus(x.status) !== "cancel"
  );

  try {
    const monthlyMap = buildMonthlyDistanceMapForCurrentMonth();
    const timelineHtml = buildRotationTimelineHtmlSafe(vehicles, activeItems);
    const cards = buildDailyDispatchVehicleCards(vehicles, activeItems, monthlyMap);

    const overviewHtml = `
      <div class="vehicle-result-card" style="margin-bottom:16px;">
        <div class="vehicle-result-head" style="align-items:flex-start;gap:12px;">
          <div class="vehicle-result-title">
            <h4>全配車 俯瞰マップ</h4>
            <div class="vehicle-result-meta">各車両は色分け表示。凡例クリックで表示ON/OFFできます。</div>
          </div>
          <div class="vehicle-result-badges">
            <span class="metric-badge">起点 黒</span>
            <span class="metric-badge">車両別 色分け</span>
            <span class="metric-badge">ピン番号=降車順</span>
          </div>
        </div>
        <div id="dispatchOverviewLegend" style="display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 14px;"></div>
        <div id="dispatchOverviewMap" style="width:100%;height:420px;border-radius:18px;overflow:hidden;background:#0f172a;"></div>
      </div>
    `;

    const overflowRows = getOverflowDisplayRowsFromCache();
    const capacityOverflowHtml = buildCapacityOverflowHtml();
    const overflowEvaluationHtml = buildOverflowEvaluationHtml();

    const cardsHtml = cards
      .map(({ vehicle, rows, orderedRows }) => {
        const summary = getVehicleDailySummary(vehicle, orderedRows);
        const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
        const lineLabel = buildVehicleLineLabel(vehicle);
        const projectedMonthly = getVehicleProjectedMonthlyDistance(vehicle.id, monthlyMap, orderedRows);

        const body = orderedRows.length
          ? orderedRows
              .map(
                (row, index) => `
                  <div class="dispatch-row">
                    <div class="dispatch-left">
                      <span class="badge-time">${escapeHtml(getHourLabel(row.actual_hour))}</span>
                      <span class="badge-order">順番 ${index + 1}</span>
                      <span class="dispatch-name">${buildMapLinkHtml({
                        name: row.casts?.name,
                        address: row.destination_address || row.casts?.address,
                        lat: row.casts?.latitude,
                        lng: row.casts?.longitude,
                        className: "dispatch-name-link"
                      })}</span>
                      <span class="dispatch-area">${escapeHtml(normalizeAreaLabel(row.destination_area || "-"))}</span>
                      ${isManualLastTripItem(row) ? `<span class="badge-status assigned">ラスト便</span>` : ""}
                    </div>
                    <div class="dispatch-right">
                      <div class="dispatch-distance">${Number(row.distance_km || 0).toFixed(1)}km</div>
                      <select class="dispatch-vehicle-select" data-item-id="${row.id}">
                        ${vehicles
                          .map(
                            v => `
                              <option value="${v.id}" ${Number(v.id) === Number(vehicle.id) ? "selected" : ""}>
                                 ${escapeHtml(v.driver_name || v.plate_number || "-")}
                              </option>
                            `
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                `
              )
              .join("")
          : `<div class="empty-vehicle-text">送りなし</div>`;

        return `
          <div class="vehicle-result-card">
            <div class="vehicle-result-head">
              <div class="vehicle-result-title">
                <h4>
                  ${escapeHtml(vehicle.driver_name || vehicle.plate_number || "-")}
                  ${isManualLastVehicle(vehicle.id) ? `<span class="badge-status assigned" style="margin-left:8px;">手動ラスト便車両</span>` : ""}
                  ${isDriverLastTripChecked(vehicle.id) ? `<span class="badge-status assigned" style="margin-left:8px;">ラスト便チェック</span>` : ""}
                </h4>
                <div class="vehicle-result-meta">
                  ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || "-"))}
                  / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
                  / 定員${vehicle.seat_capacity ?? "-"}
                  ${isDriverLastTripChecked(vehicle.id) ? `/ ラスト便対象` : ""}
                </div>
              </div>
              <div class="vehicle-result-badges">
                <span class="metric-badge">人数 ${rows.length}</span>
                <span class="metric-badge">累計距離 ${summary.totalKm.toFixed(1)}km</span>
                <span class="metric-badge">累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}</span>
                <span class="metric-badge">累計件数 ${summary.jobCount}件</span>
                ${buildVehicleRouteMapUrl(vehicle, orderedRows) ? `<a href="${escapeHtml(buildVehicleRouteMapUrl(vehicle, orderedRows))}" target="_blank" rel="noopener noreferrer" class="metric-badge" style="text-decoration:none;">Google Mapsで開く</a>` : ""}
              </div>
            </div>
            <div class="vehicle-result-body">${body}</div>
            ${orderedRows.length ? `
              <div class="dispatch-meta" style="margin-top:10px; font-size:12px; color:#9aa3b2; line-height:1.8;">
                戻り ${escapeHtml(forecast.returnAfterLabel)}
                / 次便可能 ${escapeHtml(forecast.predictedReadyTime)}
                / 累計距離 ${summary.totalKm.toFixed(1)}km
                / 累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}
                / 累計件数 ${summary.jobCount}件
                / 月間見込 ${projectedMonthly.toFixed(1)}km
                ${forecast.extraSharedDelayMinutes > 0 ? `/ 同乗追加遅延 ${forecast.extraSharedDelayMinutes}分` : ""}
              </div>
            ` : `
              <div class="dispatch-meta" style="margin-top:10px; font-size:12px; color:#9aa3b2; line-height:1.8;">
                累計距離 ${summary.totalKm.toFixed(1)}km
                / 累計時間 ${escapeHtml(formatMinutesAsJa(summary.driveMinutes))}
                / 累計件数 ${summary.jobCount}件
                / 月間見込 ${projectedMonthly.toFixed(1)}km
              </div>
            `}
          </div>
        `;
      })
      .join("");

    const overflowHtml = overflowRows.length
      ? `
        <div class="vehicle-result-card" style="border:1px solid rgba(245,158,11,.35);">
          <div class="vehicle-result-head">
            <div class="vehicle-result-title">
              <h4>あぶれ方面</h4>
              <div class="vehicle-result-meta">可能台数を超えた方面は未配車のまま表示しています</div>
            </div>
            <div class="vehicle-result-badges">
              <span class="metric-badge">件数 ${overflowRows.length}件</span>
            </div>
          </div>
          <div class="vehicle-result-body">
            ${overflowRows.map((row, index) => `
              <div class="dispatch-row">
                <div class="dispatch-left">
                  <span class="badge-time">${escapeHtml(getHourLabel(row.actual_hour))}</span>
                  <span class="badge-order">あぶれ ${index + 1}</span>
                  <span class="dispatch-name">${buildMapLinkHtml({
                    name: row.casts?.name,
                    address: row.destination_address || row.casts?.address,
                    lat: row.casts?.latitude,
                    lng: row.casts?.longitude,
                    className: "dispatch-name-link"
                  })}</span>
                  <span class="dispatch-area">${escapeHtml(row.overflow_group || normalizeAreaLabel(row.destination_area || '-'))}</span>
                  <span class="badge-status canceled">未配車</span>
                </div>
                <div class="dispatch-right">
                  <div class="dispatch-distance">${Number(row.distance_km || row.casts?.distance_km || 0).toFixed(1)}km</div>
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      `
      : "";

    els.dailyDispatchResult.innerHTML = timelineHtml + overviewHtml + cardsHtml + capacityOverflowHtml + overflowEvaluationHtml + overflowHtml;
    renderOperationAndSimulationUI();
    renderDispatchOverviewMap(cards);

    els.dailyDispatchResult.querySelectorAll(".dispatch-vehicle-select").forEach(select => {
      select.addEventListener("change", async () => {
        const itemId = normalizeDispatchEntityId(select.dataset.itemId);
        const localVehicleId = Number(select.value || 0);
        const vehicle = allVehiclesCache.find(v => Number(v.id) === localVehicleId);
        const cloudVehicleId = resolveVehicleCloudRowId(localVehicleId);

        const { error } = await supabaseClient
          .from(getDispatchUnifiedTableName())
          .update({
            vehicle_id: normalizeDispatchEntityId(cloudVehicleId || localVehicleId),
            driver_name: vehicle?.driver_name || null
          })
          .eq("id", itemId);

        if (error) {
          alert(error.message);
          return;
        }

        await addHistory(currentDispatchId, itemId, "change_vehicle", "車両を変更");
        await loadActualsByDate(els.actualDate?.value || todayStr());
        renderDailyDispatchResult();
      });
    });
  } catch (error) {
    console.error("renderDailyDispatchResult error:", error);
    els.dailyDispatchResult.innerHTML = `<div class="muted">配車結果の表示でエラーが発生しました</div>`;
    renderOperationAndSimulationUI();
  }
}

async function clearAllActuals() {
  return await clearAllActualsSingle();
}

function getVehicleDailyRunsTableName() {
  return String(window?.DROP_OFF_TABLES?.vehicle_daily_runs || "dropoff_vehicle_daily_runs");
}

function rememberResolvedWorkspaceTeamId(teamId) {
  const normalized = String(teamId || '').trim();
  if (!normalized) return null;
  window.__DROP_OFF_CURRENT_TEAM_ID__ = normalized;
  try {
    window.localStorage.setItem('__DROP_OFF_LAST_TEAM_ID__', normalized);
  } catch (_) {}
  return normalized;
}

async function resolveWorkspaceTeamIdForDailyRuns() {
  const normalize = value => {
    const raw = String(value || '').trim();
    return raw || null;
  };

  try {
    if (typeof ensureDropOffWorkspaceId === "function") {
      const ensured = normalize(await ensureDropOffWorkspaceId());
      if (ensured) return rememberResolvedWorkspaceTeamId(ensured);
    }
  } catch (error) {
    console.warn('resolveWorkspaceTeamIdForDailyRuns ensure failed:', error);
  }

  const directCandidates = [
    currentUserProfile?.team_id,
    window.currentUserProfile?.team_id,
    window.currentProfile?.team_id,
    window.__DROP_OFF_CURRENT_TEAM_ID__
  ];
  for (const candidate of directCandidates) {
    const normalized = normalize(candidate);
    if (normalized) return rememberResolvedWorkspaceTeamId(normalized);
  }

  const uid = currentUser?.id || window.currentUser?.id || null;
  if (uid && typeof getDropOffWorkspaceCacheKey === 'function') {
    try {
      const cached = normalize(window.localStorage.getItem(getDropOffWorkspaceCacheKey(uid)));
      if (cached) return rememberResolvedWorkspaceTeamId(cached);
    } catch (_) {}
  }

  const caches = [allVehiclesCache, allOriginsCache, allCastsCache, currentPlansCache, currentActualsCache];
  for (const rows of caches) {
    const matched = (Array.isArray(rows) ? rows : []).find(row => normalize(row?.team_id));
    const normalized = normalize(matched?.team_id);
    if (normalized) return rememberResolvedWorkspaceTeamId(normalized);
  }

  try {
    const lastTeamId = normalize(window.localStorage.getItem('__DROP_OFF_LAST_TEAM_ID__'));
    if (lastTeamId) return rememberResolvedWorkspaceTeamId(lastTeamId);
  } catch (_) {}

  if (!window.supabaseClient || !uid) return null;

  try {
    const membersTable = getTableName('team_members');
    const memberRes = await supabaseClient
      .from(membersTable)
      .select('team_id,user_id,member_email')
      .eq('user_id', uid)
      .limit(1);
    if (!memberRes.error && Array.isArray(memberRes.data) && memberRes.data[0]?.team_id) {
      const teamId = normalize(memberRes.data[0].team_id);
      if (teamId) return rememberResolvedWorkspaceTeamId(teamId);
    }

    const email = String(currentUser?.email || window.currentUser?.email || '').trim();
    if (email) {
      const emailRes = await supabaseClient
        .from(membersTable)
        .select('team_id,member_email')
        .eq('member_email', email)
        .limit(1);
      if (!emailRes.error && Array.isArray(emailRes.data) && emailRes.data[0]?.team_id) {
        const teamId = normalize(emailRes.data[0].team_id);
        if (teamId) return rememberResolvedWorkspaceTeamId(teamId);
      }
    }
  } catch (error) {
    console.warn('resolveWorkspaceTeamIdForDailyRuns member fallback failed:', error);
  }

  return null;
}

function mergeDailyRunPayloadWithExistingRow(nextPayload, existingRow) {
  const nextDistance = Number(nextPayload?.reference_distance_km || 0);
  const nextTripCount = Number(nextPayload?.trip_count || 0);
  const nextDriveMinutes = Number(nextPayload?.drive_minutes || 0);

  if (!existingRow) {
    return {
      ...nextPayload,
      reference_distance_km: Number(nextDistance.toFixed(1)),
      trip_count: nextTripCount,
      drive_minutes: nextDriveMinutes,
      is_workday: nextPayload?.is_workday !== false
    };
  }

  const existingDistance = Number(existingRow?.reference_distance_km || 0);
  const existingTripCount = Number(existingRow?.trip_count || 0);
  const existingDriveMinutes = Number(existingRow?.drive_minutes || 0);

  return {
    ...nextPayload,
    reference_distance_km: Number(Math.max(existingDistance, nextDistance).toFixed(1)),
    trip_count: Math.max(existingTripCount, nextTripCount),
    drive_minutes: Math.max(existingDriveMinutes, nextDriveMinutes),
    is_workday: existingRow?.is_workday !== false || nextPayload?.is_workday !== false
  };
}

async function confirmDailyToMonthly() {
  const reportDate = els.dispatchDate?.value || todayStr();
  const workspaceTeamId = await resolveWorkspaceTeamIdForDailyRuns();

  if (!workspaceTeamId) {
    alert("team_id を取得できないため保存できません");
    return;
  }

  const selectedVehicles = getSelectedVehiclesForToday();
  if (!selectedVehicles.length) {
    alert("先に可能車両を選択してください");
    return;
  }

  const payloads = [];

  selectedVehicles.forEach(vehicle => {
    const localVehicleId = Number(vehicle?.id || 0);
    if (!(localVehicleId > 0)) return;

    const cloudVehicleId = resolveVehicleCloudRowId(vehicle?.id);
    if (!cloudVehicleId) return;

    const assignedRows = Array.isArray(currentActualsCache)
      ? currentActualsCache.filter(item =>
          sameVehicleAssignmentId(item?.vehicle_id, localVehicleId) &&
          normalizeStatus(item?.status) !== "cancel"
        )
      : [];

    const summary = getVehicleLiveDailySummary(vehicle, assignedRows);
    const totalKm = Number(summary?.totalKm || 0);
    const tripCount = Number(summary?.jobCount || 0);
    const driveMinutes = Math.round(Number(summary?.driveMinutes || 0));

    if (!(tripCount > 0 || totalKm > 0 || driveMinutes > 0)) return;

    payloads.push({
      team_id: workspaceTeamId,
      vehicle_id: cloudVehicleId,
      run_date: reportDate,
      reference_distance_km: Number(totalKm.toFixed(1)),
      trip_count: tripCount,
      drive_minutes: driveMinutes,
      is_workday: true
    });
  });

  if (!payloads.length) {
    alert("保存対象の配車結果がありません");
    return;
  }

  const tableName = getVehicleDailyRunsTableName();
  const vehicleIds = [...new Set(payloads.map(row => String(row?.vehicle_id || "").trim()).filter(Boolean))];
  let mergedPayloads = payloads;

  if (vehicleIds.length) {
    let existingQuery = supabaseClient
      .from(tableName)
      .select("team_id,vehicle_id,run_date,reference_distance_km,trip_count,drive_minutes,is_workday")
      .eq("team_id", workspaceTeamId)
      .eq("run_date", reportDate)
      .in("vehicle_id", vehicleIds);

    let existingRes = await existingQuery;
    if (existingRes.error && typeof isMissingColumnError === "function" && isMissingColumnError(existingRes.error) && /team_id/i.test(String(existingRes.error?.message || ""))) {
      existingRes = await supabaseClient
        .from(tableName)
        .select("vehicle_id,run_date,reference_distance_km,trip_count,drive_minutes,is_workday")
        .eq("run_date", reportDate)
        .in("vehicle_id", vehicleIds);
    }

    if (existingRes.error) {
      if (typeof isMissingTableError === "function" && isMissingTableError(existingRes.error)) {
        console.error(existingRes.error);
        alert("dropoff_vehicle_daily_runs テーブルが未作成です。先にSQLを実行してください。");
        return;
      }
      console.error(existingRes.error);
      alert("既存の日次実績確認に失敗しました: " + existingRes.error.message);
      return;
    }

    const existingMap = new Map((Array.isArray(existingRes.data) ? existingRes.data : []).map(row => [String(row?.vehicle_id || "").trim(), row]));
    mergedPayloads = payloads.map(row => mergeDailyRunPayloadWithExistingRow(row, existingMap.get(String(row?.vehicle_id || "").trim())));
  }

  const { error } = await supabaseClient
    .from(tableName)
    .upsert(mergedPayloads, { onConflict: "team_id,vehicle_id,run_date" });

  if (error) {
    if (typeof isMissingTableError === "function" && isMissingTableError(error)) {
      console.error(error);
      alert("dropoff_vehicle_daily_runs テーブルが未作成です。先にSQLを実行してください。");
      return;
    }
    console.error(error);
    alert("日次実績の保存に失敗しました: " + error.message);
    return;
  }

  await addHistory(currentDispatchId, null, "confirm_daily", "配車結果を日次実績へ保存");
  alert("日次実績を保存しました");
  await loadVehicles();
  await loadHomeAndAll();
}

async function resetMonthlySummary() {
  if (!window.confirm("今月の走行記録を削除しますか？")) return;

  const dateStr = els.dispatchDate?.value || todayStr();
  const monthKey = getMonthKey(dateStr);
  const monthStart = `${monthKey}-01`;
  const start = new Date(`${monthStart}T00:00:00`);
  const next = new Date(start.getFullYear(), start.getMonth() + 1, 1);
  const nextStr = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-01`;
  const workspaceTeamId = typeof resolveWorkspaceTeamIdForDailyRuns === "function"
    ? await resolveWorkspaceTeamIdForDailyRuns()
    : null;

  let query = supabaseClient
    .from(getVehicleDailyRunsTableName())
    .delete()
    .gte("run_date", monthStart)
    .lt("run_date", nextStr);

  if (workspaceTeamId) query = query.eq("team_id", workspaceTeamId);

  const { error } = await query;

  if (error) {
    alert(error.message);
    return;
  }

  await addHistory(null, null, "reset_monthly_reports", `${monthKey} の月間距離/出勤日数をリセット`);
  if (typeof window.__dropoffRefreshMonthlyUi === 'function') {
    await window.__dropoffRefreshMonthlyUi(dateStr);
  }
  renderManualLastVehicleInfo();
}

function getDropOffLightHistoryStorageKey() {
  return "__DROP_OFF_LIGHT_HISTORY_V1__";
}

function readDropOffLightHistoryEntries() {
  try {
    const raw = window.localStorage.getItem(getDropOffLightHistoryStorageKey());
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function writeDropOffLightHistoryEntries(entries) {
  const safeEntries = (Array.isArray(entries) ? entries : []).slice(0, 200);
  try {
    window.localStorage.setItem(getDropOffLightHistoryStorageKey(), JSON.stringify(safeEntries));
  } catch (error) {
    console.warn("light history save failed:", error);
  }
  return safeEntries;
}

async function resolveWorkspaceTeamIdForHistory() {
  try {
    if (typeof resolveWorkspaceTeamIdForDailyRuns === "function") {
      const resolved = await resolveWorkspaceTeamIdForDailyRuns();
      if (resolved) return resolved;
    }
  } catch (error) {
    console.warn("resolveWorkspaceTeamIdForHistory failed:", error);
  }

  const candidates = [
    window.__DROP_OFF_CURRENT_TEAM_ID__,
    currentUserProfile?.team_id,
    window.currentUserProfile?.team_id,
    window.currentProfile?.team_id
  ];
  for (const candidate of candidates) {
    const normalized = String(candidate || "").trim();
    if (normalized) return normalized;
  }

  try {
    const cached = String(window.localStorage.getItem("__DROP_OFF_LAST_TEAM_ID__") || "").trim();
    if (cached) return cached;
  } catch (_) {}

  return null;
}

async function appendDropOffLightHistoryEntry(entry) {
  const teamId = entry?.team_id || await resolveWorkspaceTeamIdForHistory();
  const createdAt = entry?.created_at || new Date().toISOString();
  const safeEntry = {
    id: entry?.id || `light_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    dispatch_id: entry?.dispatch_id ?? null,
    item_id: entry?.item_id ?? null,
    action: String(entry?.action || ""),
    message: String(entry?.message || ""),
    acted_by: entry?.acted_by || getCurrentUserIdSafe() || null,
    team_id: teamId || null,
    created_at: createdAt,
    source: "light"
  };

  const entries = readDropOffLightHistoryEntries();
  entries.unshift(safeEntry);
  writeDropOffLightHistoryEntries(entries);

  try {
    if (typeof window.__DROP_OFF_REFRESH_HISTORY__ === "function") {
      await window.__DROP_OFF_REFRESH_HISTORY__();
    }
  } catch (error) {
    console.warn("light history refresh failed:", error);
  }

  return safeEntry;
}

async function getDropOffLightHistoryEntries(limit = 100) {
  const teamId = await resolveWorkspaceTeamIdForHistory();
  const entries = readDropOffLightHistoryEntries()
    .filter(row => !teamId || String(row?.team_id || "").trim() === teamId)
    .sort((a, b) => String(b?.created_at || "").localeCompare(String(a?.created_at || "")));
  return entries.slice(0, Math.max(1, Number(limit) || 100));
}

window.__DROP_OFF_HISTORY_MODE__ = "light";
window.__DROP_OFF_GET_LIGHT_HISTORY__ = getDropOffLightHistoryEntries;
window.__DROP_OFF_APPEND_LIGHT_HISTORY__ = appendDropOffLightHistoryEntry;

async function addHistory(dispatchId, itemId, action, message) {
  const safeDispatchId = Number.isFinite(Number(dispatchId)) && Number(dispatchId) > 0 ? Number(dispatchId) : null;
  const safeItemId = normalizeDispatchEntityId(itemId) || null;

  await appendDropOffLightHistoryEntry({
    dispatch_id: safeDispatchId,
    item_id: safeItemId,
    action,
    message,
    acted_by: getCurrentUserIdSafe()
  });
}


function applyFetchFilters(query, filters = []) {
  (Array.isArray(filters) ? filters : []).forEach(filter => {
    if (!filter || !filter.column || !filter.op) return;
    if (filter.op === "eq") query = query.eq(filter.column, filter.value);
    if (filter.op === "gte") query = query.gte(filter.column, filter.value);
    if (filter.op === "lte") query = query.lte(filter.column, filter.value);
    if (filter.op === "lt") query = query.lt(filter.column, filter.value);
    if (filter.op === "gt") query = query.gt(filter.column, filter.value);
  });
  return query;
}

function isLikelyUuid(value) {
  const str = String(value || "").trim();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(str);
}

function collectWorkspaceScopedIds(kind) {
  const asStrings = values => [...new Set((Array.isArray(values) ? values : [])
    .filter(value => value !== null && value !== undefined && String(value).trim() !== "")
    .map(value => String(value).trim()))];
  const asUuidStrings = values => asStrings(values).filter(isLikelyUuid);

  if (kind === "vehicles") {
    return asUuidStrings((Array.isArray(allVehiclesCache) ? allVehiclesCache : []).flatMap(row => [row?.cloud_row_id, row?.db_id, row?.id]));
  }
  if (kind === "origins") {
    return asUuidStrings((Array.isArray(allOriginsCache) ? allOriginsCache : []).flatMap(row => [row?.id]));
  }
  if (kind === "casts") {
    return asUuidStrings((Array.isArray(allCastsCache) ? allCastsCache : []).flatMap(row => [row?.id, row?.cast_id]));
  }
  if (kind === "dispatches") {
    return asUuidStrings([
      ...(Array.isArray(currentPlansCache) ? currentPlansCache : []).map(row => row?.id),
      ...(Array.isArray(currentActualsCache) ? currentActualsCache : []).map(row => row?.id)
    ]);
  }
  if (kind === "vehicle_daily_runs") {
    return asUuidStrings((Array.isArray(currentDailyReportsCache) ? currentDailyReportsCache : []).flatMap(row => [row?.id]));
  }
  return [];
}

async function deleteRowsByIds(tableName, ids = [], idColumn = "id") {
  const normalizedIds = [...new Set((Array.isArray(ids) ? ids : [])
    .filter(value => value !== null && value !== undefined && String(value).trim() !== "")
    .map(value => String(value).trim()))];
  const safeIds = idColumn === "id" ? normalizedIds.filter(isLikelyUuid) : normalizedIds;

  if (idColumn === "id" && normalizedIds.length !== safeIds.length) {
    console.warn(`deleteRowsByIds skipped non-UUID ids for ${tableName}:`, normalizedIds.filter(v => !isLikelyUuid(v)));
  }

  if (!safeIds.length) return { data: [], error: null, deletedByIds: false };

  const pageSize = 200;
  for (let i = 0; i < safeIds.length; i += pageSize) {
    const chunk = safeIds.slice(i, i + pageSize);
    const { error } = await supabaseClient
      .from(tableName)
      .delete()
      .in(idColumn, chunk);
    if (error) return { data: null, error, deletedByIds: true };
  }

  return { data: [], error: null, deletedByIds: true };
}

async function deleteWorkspaceScopedRows(tableName, workspaceTeamId, options = {}) {
  const idColumn = options.idColumn || "id";
  const fallbackIds = Array.isArray(options.fallbackIds) ? options.fallbackIds : [];

  let query = supabaseClient.from(tableName).delete();
  if (workspaceTeamId) query = query.eq("team_id", workspaceTeamId);
  const result = await query;

  if (!result.error) return result;
  if (!(typeof isMissingColumnError === "function" && isMissingColumnError(result.error) && /team_id/i.test(String(result.error?.message || "")))) {
    return result;
  }

  console.warn(`deleteWorkspaceScopedRows fallback without team_id: ${tableName}`);
  return await deleteRowsByIds(tableName, fallbackIds, idColumn);
}

async function fetchAllTableRows(tableName, orderColumn = "id", options = {}) {
  const pageSize = 1000;
  const teamId = options.useTeamScope === false
    ? null
    : String(options.teamId || await resolveWorkspaceTeamIdForDailyRuns() || "").trim() || null;
  const filters = Array.isArray(options.filters) ? options.filters : [];
  let from = 0;
  let allRows = [];
  let useTeamScope = Boolean(teamId);
  let warnedMissingTeamScope = false;

  while (true) {
    let query = supabaseClient
      .from(tableName)
      .select("*")
      .order(orderColumn, { ascending: true });

    if (useTeamScope && teamId) {
      query = query.eq("team_id", teamId);
    }

    query = applyFetchFilters(query, filters);

    let { data, error } = await query.range(from, from + pageSize - 1);

    if (error && useTeamScope && typeof isMissingColumnError === "function" && isMissingColumnError(error) && /team_id/i.test(String(error?.message || ""))) {
      useTeamScope = false;
      if (!warnedMissingTeamScope) {
        console.warn(`fetchAllTableRows fallback without team_id: ${tableName}`);
        warnedMissingTeamScope = true;
      }
      let fallbackQuery = supabaseClient
        .from(tableName)
        .select("*")
        .order(orderColumn, { ascending: true });
      fallbackQuery = applyFetchFilters(fallbackQuery, filters);
      ({ data, error } = await fallbackQuery.range(from, from + pageSize - 1));
    }

    if (error) throw error;

    const rows = data || [];
    allRows = allRows.concat(rows);

    if (rows.length < pageSize) break;
    from += pageSize;
  }

  return allRows;
}

function stripMetaForInsert(row, extraRemoveKeys = []) {
  const clone = { ...row };
  const removeKeys = [
    "id",
    "created_at",
    "updated_at",
    ...extraRemoveKeys
  ];

  removeKeys.forEach(key => {
    if (key in clone) delete clone[key];
  });

  return clone;
}

function normalizeBackupIdKey(value) {
  const raw = String(value ?? "").trim();
  return raw || null;
}

function remapBackupRelationId(value, idMap) {
  const key = normalizeBackupIdKey(value);
  if (!key) return null;
  return idMap.get(key) || null;
}

function buildLegacyPlanDispatchRow(oldRow = {}) {
  const hour = Number(oldRow?.dispatch_hour ?? oldRow?.plan_hour ?? 0);
  const dispatchDate = oldRow?.dispatch_date || oldRow?.plan_date || todayStr();
  const castId = oldRow?.cast_id ?? oldRow?.person_id ?? null;
  return {
    ...oldRow,
    dispatch_kind: "plan",
    dispatch_date: dispatchDate,
    plan_date: dispatchDate,
    dispatch_hour: hour,
    plan_hour: hour,
    cast_id: castId,
    person_id: castId,
    destination_area: oldRow?.destination_area || oldRow?.planned_area || null,
    status: oldRow?.status || "planned"
  };
}

function buildLegacyActualDispatchRow(oldRow = {}) {
  const hour = Number(oldRow?.dispatch_hour ?? oldRow?.actual_hour ?? oldRow?.plan_hour ?? 0);
  const dispatchDate = oldRow?.dispatch_date || oldRow?.plan_date || todayStr();
  const castId = oldRow?.cast_id ?? oldRow?.person_id ?? null;
  return {
    ...oldRow,
    dispatch_kind: "actual",
    dispatch_date: dispatchDate,
    plan_date: dispatchDate,
    dispatch_hour: hour,
    actual_hour: Number(oldRow?.actual_hour ?? hour),
    cast_id: castId,
    person_id: castId,
    destination_area: oldRow?.destination_area || oldRow?.planned_area || null,
    status: oldRow?.status || "pending"
  };
}

function normalizeDispatchBackupRows(backup = {}) {
  const directRows = (Array.isArray(backup?.dispatches) ? backup.dispatches : []).filter(row => {
    const kind = String(row?.dispatch_kind || "").trim();
    return kind === "plan" || kind === "actual";
  });

  if (directRows.length) {
    return directRows;
  }

  const rows = [];
  (Array.isArray(backup?.dispatch_plans) ? backup.dispatch_plans : []).forEach(row => {
    rows.push(buildLegacyPlanDispatchRow(row));
  });
  (Array.isArray(backup?.dispatch_items) ? backup.dispatch_items : []).forEach(row => {
    rows.push(buildLegacyActualDispatchRow(row));
  });
  return rows;
}

function normalizeDailyRunsBackupRows(backup = {}) {
  const directRows = Array.isArray(backup?.vehicle_daily_runs) ? backup.vehicle_daily_runs : [];
  if (directRows.length) return directRows;

  return (Array.isArray(backup?.vehicle_daily_reports) ? backup.vehicle_daily_reports : []).map(row => ({
    ...row,
    run_date: row?.run_date || row?.report_date || todayStr(),
    reference_distance_km: Number(row?.reference_distance_km ?? row?.distance_km ?? 0),
    trip_count: Number(row?.trip_count ?? 0),
    drive_minutes: Number(row?.drive_minutes ?? 0),
    is_workday: row?.is_workday ?? true
  }));
}

function buildWorkspaceScopedRow(row, workspaceTeamId, extraRemoveKeys = []) {
  const clone = stripMetaForInsert(row, extraRemoveKeys);
  clone.team_id = workspaceTeamId;
  return clone;
}

function extractMissingSchemaColumnName(error) {
  const message = String(error?.message || "");
  const patterns = [
    /Could not find the '([^']+)' column/i,
    /column\s+"?([^"\s]+)"?\s+does not exist/i,
    /Could not find the ([^\s]+) column/i
  ];
  for (const pattern of patterns) {
    const match = message.match(pattern);
    if (match?.[1]) return String(match[1]).trim();
  }
  return null;
}

async function insertRowWithSchemaFallback(tableName, row, options = {}) {
  let payload = { ...(row || {}) };
  const removedColumns = [];
  const select = String(options.select || "").trim();
  const single = options.single !== false;

  for (let attempt = 0; attempt < 10; attempt += 1) {
    let query = supabaseClient.from(tableName).insert(payload);
    if (select) {
      query = query.select(select);
      if (single) query = query.single();
    }

    const result = await query;
    if (!result.error) {
      return { ...result, removedColumns };
    }

    const missingColumn = extractMissingSchemaColumnName(result.error);
    if (!missingColumn || !(missingColumn in payload)) {
      return { ...result, removedColumns };
    }

    removedColumns.push(missingColumn);
    delete payload[missingColumn];
  }

  return {
    data: null,
    error: new Error(`${tableName} insert fallback exceeded retry limit`),
    removedColumns
  };
}

async function exportWorkspaceBackupByTeamId(workspaceTeamId, options = {}) {
  const safeTeamId = String(workspaceTeamId || '').trim() || null;
  const teamName = String(options.teamName || getPlatformAdminTeamLabel(safeTeamId, 'workspace')).trim() || 'workspace';
  const filePrefix = String(options.filePrefix || 'dropoff_workspace_backup').trim() || 'dropoff_workspace_backup';
  const successStatus = String(options.successStatus || '').trim();
  const writeHistory = options.writeHistory === true;

  if (!safeTeamId) {
    alert('対象チームを特定できないため、バックアップできません');
    return false;
  }

  try {
    const [
      origins,
      casts,
      vehicles,
      dispatches,
      vehicleDailyRuns
    ] = await Promise.all([
      fetchAllTableRows(getTableName("origins"), "id", { teamId: safeTeamId }),
      fetchAllTableRows(getTableName("casts"), "id", { teamId: safeTeamId }),
      fetchAllTableRows(getTableName("vehicles"), "id", { teamId: safeTeamId }),
      fetchAllTableRows(getTableName("dispatches"), "id", { teamId: safeTeamId }),
      fetchAllTableRows(getVehicleDailyRunsTableName(), "run_date", { teamId: safeTeamId })
    ]);

    const payload = {
      app: "DROP OFF",
      version: 3,
      scope: "workspace",
      exported_at: new Date().toISOString(),
      workspace_team_id: safeTeamId,
      team_id: safeTeamId,
      team_name: teamName,
      origin: isSameTeamId(window.currentWorkspaceTeamId, safeTeamId)
        ? {
            label: ORIGIN_LABEL,
            lat: ORIGIN_LAT,
            lng: ORIGIN_LNG
          }
        : null,
      data: {
        origins,
        casts,
        vehicles,
        dispatches,
        vehicle_daily_runs: vehicleDailyRuns
      }
    };

    downloadTextFile(
      `${filePrefix}_${sanitizeDownloadLabel(teamName, 'team')}_${todayStr()}.json`,
      JSON.stringify(payload, null, 2),
      "application/json;charset=utf-8"
    );

    if (writeHistory && isSameTeamId(window.currentWorkspaceTeamId, safeTeamId)) {
      await addHistory(null, null, "export_all", "ワークスペースバックアップを出力");
    }
    if (successStatus) setPlatformAdminActionStatus(successStatus, false);
    return true;
  } catch (error) {
    console.error("exportWorkspaceBackupByTeamId error:", error);
    alert("バックアップの出力に失敗しました: " + error.message);
    if (successStatus) setPlatformAdminActionStatus('バックアップに失敗しました: ' + (error.message || error), true);
    return false;
  }
}

async function importWorkspaceDataFromInput(inputEl, options = {}) {
  const file = inputEl?.files?.[0];
  const explicitTeamId = String(options.teamId || '').trim() || null;
  const teamName = String(options.teamName || getPlatformAdminTeamLabel(explicitTeamId, 'このチーム')).trim() || 'このチーム';
  const writeHistory = options.writeHistory === true;
  const successMessage = String(options.successMessage || 'ワークスペースのインポートが完了しました').trim();
  const successStatus = String(options.successStatus || '').trim();

  if (!file) {
    alert('JSONファイルを選択してください');
    return false;
  }

  try {
    const text = await file.text();
    const json = JSON.parse(text);

    if (!json?.data) {
      alert('バックアップJSONの形式が正しくありません');
      return false;
    }

    let activeWorkspaceTeamId = explicitTeamId || String(await resolveWorkspaceTeamIdForDailyRuns() || '').trim() || null;
    if (!activeWorkspaceTeamId) {
      const backupTeamId = rememberResolvedWorkspaceTeamId(json?.workspace_team_id || json?.team_id || null);
      if (backupTeamId) activeWorkspaceTeamId = backupTeamId;
    }
    if (!activeWorkspaceTeamId) {
      alert('現在のワークスペースを特定できないため、復元できません');
      return false;
    }

    const backup = json.data;
    const backupDispatchRows = normalizeDispatchBackupRows(backup);
    const backupDailyRuns = normalizeDailyRunsBackupRows(backup);
    const backupOrigins = Array.isArray(backup?.origins) ? backup.origins : [];
    const backupCasts = Array.isArray(backup?.casts) ? backup.casts : [];
    const backupVehicles = Array.isArray(backup?.vehicles) ? backup.vehicles : [];

    const proceed = window.confirm(
      `${teamName} の現在データを消去して、バックアップJSONから復元しますか？`
    );
    if (!proceed) {
      if (inputEl) inputEl.value = '';
      return false;
    }

    const proceed2 = window.confirm(
      `本当に実行しますか？ ${teamName} のデータは、一度削除してから復元します。

復元対象
起点: ${backupOrigins.length}件
送り先: ${backupCasts.length}件
車両: ${backupVehicles.length}件
配車: ${backupDispatchRows.length}件
日次実績: ${backupDailyRuns.length}件`
    );
    if (!proceed2) {
      if (inputEl) inputEl.value = '';
      return false;
    }

    const deleteTargets = [
      { table: getTableName("dispatches"), fallbackIds: isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId) ? collectWorkspaceScopedIds("dispatches") : [] },
      { table: getVehicleDailyRunsTableName(), fallbackIds: isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId) ? collectWorkspaceScopedIds("vehicle_daily_runs") : [] },
      { table: getTableName("casts"), fallbackIds: isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId) ? collectWorkspaceScopedIds("casts") : [] },
      { table: getTableName("vehicles"), fallbackIds: isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId) ? collectWorkspaceScopedIds("vehicles") : [] },
      { table: getTableName("origins"), fallbackIds: isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId) ? collectWorkspaceScopedIds("origins") : [] }
    ];

    for (const target of deleteTargets) {
      const { error } = await deleteWorkspaceScopedRows(target.table, activeWorkspaceTeamId, { fallbackIds: target.fallbackIds });
      if (error) {
        console.error(`${target.table} delete error:`, error);
        alert(`${target.table} の削除に失敗しました: ${error.message}`);
        return false;
      }
    }

    const castIdMap = new Map();
    const vehicleIdMap = new Map();

    for (const oldRow of backupOrigins) {
      const row = buildWorkspaceScopedRow(oldRow, activeWorkspaceTeamId);
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();
      const { error, removedColumns } = await insertRowWithSchemaFallback(getTableName("origins"), row);
      reportImportRemovedColumns('origins', removedColumns);
      if (error) throw error;
    }

    for (const oldRow of backupCasts) {
      const row = buildWorkspaceScopedRow(oldRow, activeWorkspaceTeamId);
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();
      const { data, error, removedColumns } = await insertRowWithSchemaFallback(getTableName("casts"), row, { select: "id", single: true });
      reportImportRemovedColumns('casts', removedColumns);
      if (error) throw error;
      const oldIdKey = normalizeBackupIdKey(oldRow?.id || oldRow?.cast_id);
      const newIdKey = normalizeBackupIdKey(data?.id);
      if (oldIdKey && newIdKey) castIdMap.set(oldIdKey, newIdKey);
    }

    for (const oldRow of backupVehicles) {
      const row = buildWorkspaceScopedRow(oldRow, activeWorkspaceTeamId);
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();
      const { data, error, removedColumns } = await insertRowWithSchemaFallback(getTableName("vehicles"), row, { select: "id", single: true });
      reportImportRemovedColumns('vehicles', removedColumns);
      if (error) throw error;
      const oldIdKey = normalizeBackupIdKey(oldRow?.id || oldRow?.vehicle_id);
      const newIdKey = normalizeBackupIdKey(data?.id);
      if (oldIdKey && newIdKey) vehicleIdMap.set(oldIdKey, newIdKey);
    }

    for (const oldRow of backupDispatchRows) {
      const row = buildWorkspaceScopedRow(oldRow, activeWorkspaceTeamId);
      const mappedCastId = remapBackupRelationId(oldRow?.cast_id ?? oldRow?.person_id, castIdMap);
      const mappedVehicleId = remapBackupRelationId(oldRow?.vehicle_id, vehicleIdMap);
      row.cast_id = mappedCastId;
      row.person_id = mappedCastId;
      row.vehicle_id = mappedVehicleId;
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();
      const { error, removedColumns } = await insertRowWithSchemaFallback(getTableName("dispatches"), row);
      reportImportRemovedColumns('dispatches', removedColumns);
      if (error) throw error;
    }

    for (const oldRow of backupDailyRuns) {
      const row = buildWorkspaceScopedRow(oldRow, activeWorkspaceTeamId, ["report_date"]);
      row.vehicle_id = remapBackupRelationId(oldRow?.vehicle_id, vehicleIdMap);
      row.run_date = row.run_date || oldRow?.report_date || todayStr();
      row.reference_distance_km = Number(oldRow?.reference_distance_km ?? oldRow?.distance_km ?? row.reference_distance_km ?? 0);
      row.trip_count = Number(oldRow?.trip_count ?? row.trip_count ?? 0);
      row.drive_minutes = Number(oldRow?.drive_minutes ?? row.drive_minutes ?? 0);
      row.is_workday = typeof row.is_workday === "boolean" ? row.is_workday : true;
      if (!row.created_by) row.created_by = getCurrentUserIdSafe();
      const { error, removedColumns } = await insertRowWithSchemaFallback(getVehicleDailyRunsTableName(), row);
      reportImportRemovedColumns('vehicle_daily_runs', removedColumns);
      if (error) throw error;
    }

    if (inputEl) inputEl.value = '';

    if (writeHistory && isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId)) {
      await addHistory(null, null, "import_all", "ワークスペースバックアップから復元");
    }

    if (successStatus) setPlatformAdminActionStatus(successStatus, false);
    alert(successMessage);

    if (isSameTeamId(window.currentWorkspaceTeamId, activeWorkspaceTeamId)) {
      rememberResolvedWorkspaceTeamId(activeWorkspaceTeamId);
      await loadHomeAndAll();
    syncHomeSetupCardVisibility();
      renderManualLastVehicleInfo();
    }
    await loadPlatformAdminTeams();
    return true;
  } catch (error) {
    console.error("importWorkspaceDataFromInput error:", error);
    alert("ワークスペースのインポートに失敗しました: " + error.message);
    if (successStatus) setPlatformAdminActionStatus('復元に失敗しました: ' + (error.message || error), true);
    return false;
  } finally {
    if (inputEl) inputEl.value = '';
  }
}

async function exportAllData() {
  const workspaceTeamId = await resolveWorkspaceTeamIdForDailyRuns();
  if (!workspaceTeamId) {
    alert("現在のワークスペースを特定できないため、バックアップできません");
    return;
  }
  rememberResolvedWorkspaceTeamId(workspaceTeamId);
  await exportWorkspaceBackupByTeamId(workspaceTeamId, {
    teamName: String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || 'workspace').trim() || 'workspace',
    filePrefix: 'dropoff_workspace_backup',
    writeHistory: true
  });
}

function triggerImportAll(inputEl = els.importAllFileInput) {
  const resolvedInput = (
    inputEl && typeof inputEl.click === "function" && typeof inputEl.value !== "undefined"
      ? inputEl
      : els.importAllFileInput
  );

  if (!resolvedInput) {
    alert("インポート入力を初期化できませんでした。");
    return;
  }

  try {
    resolvedInput.value = "";
  } catch (_) {}

  resolvedInput.click();
}

async function importAllDataFromFile() {
  const workspaceTeamId = await resolveWorkspaceTeamIdForDailyRuns();
  if (workspaceTeamId) rememberResolvedWorkspaceTeamId(workspaceTeamId);
  await importWorkspaceDataFromInput(els.importAllFileInput, {
    teamId: workspaceTeamId,
    teamName: String(window.currentWorkspaceInfo?.name || window.currentWorkspaceInfo?.team_name || '現在のワークスペース').trim() || '現在のワークスペース',
    writeHistory: true,
    successMessage: 'ワークスペースのインポートが完了しました'
  });
}

async function resetAllDataDanger() {
  if (!window.confirm("本当に現在のワークスペースの全データを消去しますか？この操作は元に戻せません。")) return;

  try {
    const workspaceTeamId = await resolveWorkspaceTeamIdForDailyRuns();
    if (!workspaceTeamId) {
      alert("現在のワークスペースを特定できないため、全消去できません");
      return;
    }
    rememberResolvedWorkspaceTeamId(workspaceTeamId);

    const deleteTargets = [
      { table: getTableName("dispatches"), fallbackIds: collectWorkspaceScopedIds("dispatches") },
      { table: getVehicleDailyRunsTableName(), fallbackIds: collectWorkspaceScopedIds("vehicle_daily_runs") },
      { table: getTableName("casts"), fallbackIds: collectWorkspaceScopedIds("casts") },
      { table: getTableName("vehicles"), fallbackIds: collectWorkspaceScopedIds("vehicles") },
      { table: getTableName("origins"), fallbackIds: collectWorkspaceScopedIds("origins") }
    ];

    for (const target of deleteTargets) {
      const { error } = await deleteWorkspaceScopedRows(target.table, workspaceTeamId, { fallbackIds: target.fallbackIds });
      if (error) {
        console.error(`${target.table} delete error:`, error);
        alert(`${target.table} の削除でエラー: ${error.message}`);
        return;
      }
    }

    currentDispatchId = null;
    activeVehicleIdsForToday = new Set();

    resetCastForm();
    syncCastBlankMetricsUi();
    resetVehicleForm();
    resetPlanForm();
    resetActualForm();

    alert("現在のワークスペースの全データを削除しました");
    await loadHomeAndAll();
    renderManualLastVehicleInfo();
  } catch (err) {
    console.error("resetAllDataDanger error:", err);
    alert("全消去中にエラーが発生しました");
  }
}
async function resetAllCastsDanger() {
  if (!window.confirm("本当に送り先全データを消去しますか？この操作は元に戻せません。")) return;

  let { error } = await supabaseClient
    .from(getTableName("casts"))
    .update({ is_active: false })
    .eq("is_active", true);

  if (error && typeof isMissingColumnError === "function" && isMissingColumnError(error)) {
    ({ error } = await supabaseClient
      .from(getTableName("casts"))
      .delete()
      .neq("id", 0));
  }

  if (error) {
    console.error(error);
    alert("送り先全データ消去に失敗しました: " + error.message);
    return;
  }

  await addHistory(null, null, "reset_casts", "送り先全データを消去");
  await loadCasts();
  alert("送り先全データを消去しました");
}

async function resetAllVehiclesDanger() {
  if (!window.confirm("本当に車両全データを消去しますか？この操作は元に戻せません。")) return;

  let { error } = await supabaseClient
    .from(getTableName("vehicles"))
    .update({ is_active: false })
    .eq("is_active", true);

  if (error && typeof isMissingColumnError === "function" && isMissingColumnError(error)) {
    ({ error } = await supabaseClient
      .from(getTableName("vehicles"))
      .delete()
      .neq("id", 0));
  }

  if (error) {
    console.error(error);
    alert("車両全データ消去に失敗しました: " + error.message);
    return;
  }

  await addHistory(null, null, "reset_vehicles", "車両全データを消去");
  await loadVehicles();
  alert("車両全データを消去しました");
}

async function resetAllDataDangerLegacy() {
  console.warn("resetAllDataDangerLegacy is redirected to resetAllDataDanger");
  return await resetAllDataDanger();
}

function renderDailyMileageInputs() {
  if (!els.dailyMileageInputs) return;

  const defaultDate = els.dispatchDate?.value || todayStr();
  const selectedVehicles = getSelectedVehiclesForToday();

  els.dailyMileageInputs.innerHTML = "";

  if (!selectedVehicles.length) {
    els.dailyMileageInputs.innerHTML = `<div class="muted">可能車両を選択すると入力欄が表示されます</div>`;
    return;
  }

  selectedVehicles.forEach(vehicle => {
    const existing = currentDailyReportsCache.find(
      r =>
        Number(r.vehicle_id) === Number(vehicle.id) &&
        String(r.report_date || "") === String(defaultDate || "")
    ) || currentDailyReportsCache.find(
      r => Number(r.vehicle_id) === Number(vehicle.id)
    );

    const row = document.createElement("div");
    row.className = "daily-mileage-row";
    row.innerHTML = `
      <div>
        <div class="daily-mileage-label">${escapeHtml(vehicle.plate_number || "-")}</div>
        <div class="daily-mileage-sub">
          ${escapeHtml(vehicle.driver_name || "-")} / 帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || "-"))}
        </div>
      </div>

      <div class="field">
        <label>入力日</label>
        <input
          type="date"
          class="daily-mileage-date-input"
          data-vehicle-id="${vehicle.id}"
          value="${existing?.report_date || defaultDate}"
        />
      </div>

      <div class="field">
        <label>実績走行距離(km)</label>
        <input
          type="number"
          step="0.1"
          min="0"
          class="daily-mileage-input"
          data-vehicle-id="${vehicle.id}"
          value="${existing?.distance_km ?? ""}"
          placeholder="例：72.5"
        />
      </div>

      <div class="field">
        <label>メモ</label>
        <input
          type="text"
          class="daily-mileage-note-input"
          data-vehicle-id="${vehicle.id}"
          value="${escapeHtml(existing?.note || "")}"
          placeholder="任意"
        />
      </div>
    `;
    els.dailyMileageInputs.appendChild(row);
  });
}

async function saveDailyMileageReports() {
  const workspaceTeamId = await resolveWorkspaceTeamIdForDailyRuns();
  if (!workspaceTeamId) {
    alert("team_id を取得できないため保存できません");
    return;
  }

  const selectedVehicles = getSelectedVehiclesForToday();

  if (!selectedVehicles.length) {
    alert("先に可能車両を選択してください");
    return;
  }

  const mileageInputs = [...document.querySelectorAll(".daily-mileage-input")];
  const noteInputs = [...document.querySelectorAll(".daily-mileage-note-input")];
  const dateInputs = [...document.querySelectorAll(".daily-mileage-date-input")];
  const tableName = getVehicleDailyRunsTableName();

  for (const vehicle of selectedVehicles) {
    const mileageInput = mileageInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );
    const noteInput = noteInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );
    const dateInput = dateInputs.find(
      input => Number(input.dataset.vehicleId) === Number(vehicle.id)
    );

    const reportDate = dateInput?.value || (els.dispatchDate?.value || todayStr());
    const distanceKm = toNullableNumber(mileageInput?.value);
    const note = noteInput?.value.trim() || "日次報告入力";
    const cloudVehicleId = normalizeDispatchEntityId(resolveVehicleCloudRowId(vehicle?.id) || vehicle?.cloud_row_id || vehicle?.id);

    if (!reportDate) continue;
    if (distanceKm === null) continue;
    if (!cloudVehicleId) continue;

    const corePayload = {
      team_id: workspaceTeamId,
      vehicle_id: cloudVehicleId,
      run_date: reportDate,
      reference_distance_km: Number(distanceKm),
      trip_count: 1,
      drive_minutes: 0,
      is_workday: true
    };

    const optionalPayload = {
      driver_name: vehicle.driver_name || null,
      note,
      created_by: getCurrentUserIdSafe()
    };

    let existingId = null;
    let lookupQuery = supabaseClient
      .from(tableName)
      .select("id")
      .eq("team_id", workspaceTeamId)
      .eq("vehicle_id", cloudVehicleId)
      .eq("run_date", reportDate)
      .limit(1)
      .maybeSingle();

    let lookupRes = await lookupQuery;
    if (lookupRes.error && typeof isMissingColumnError === "function" && isMissingColumnError(lookupRes.error) && /team_id/i.test(String(lookupRes.error?.message || ""))) {
      lookupRes = await supabaseClient
        .from(tableName)
        .select("id")
        .eq("vehicle_id", cloudVehicleId)
        .eq("run_date", reportDate)
        .limit(1)
        .maybeSingle();
    }

    if (lookupRes.error) {
      console.error(lookupRes.error);
      alert("日次報告の保存前確認に失敗しました: " + lookupRes.error.message);
      return;
    }

    existingId = lookupRes.data?.id || null;

    if (existingId) {
      const updateRes = await supabaseClient
        .from(tableName)
        .update(corePayload)
        .eq("id", existingId);

      if (updateRes.error) {
        console.error(updateRes.error);
        alert("日次報告の更新に失敗しました: " + updateRes.error.message);
        return;
      }

      const optionalUpdate = await supabaseClient
        .from(tableName)
        .update(optionalPayload)
        .eq("id", existingId);

      if (optionalUpdate.error) {
        const optionalCols = Object.keys(optionalPayload);
        const missingOptionalColumn = extractMissingSchemaColumnName(optionalUpdate.error);
        if (!missingOptionalColumn || !optionalCols.includes(missingOptionalColumn)) {
          console.warn("vehicle_daily_runs optional update skipped:", optionalUpdate.error.message);
        }
      }
      continue;
    }

    const insertRes = await insertRowWithSchemaFallback(tableName, { ...corePayload, ...optionalPayload });
    if (insertRes.error) {
      console.error(insertRes.error);
      alert("日次報告の保存に失敗しました: " + insertRes.error.message);
      return;
    }
    if (insertRes.removedColumns?.length) {
      console.warn("vehicle_daily_runs manual save removed columns:", insertRes.removedColumns);
    }
  }

  await addHistory(null, null, "save_daily_mileage", `日次走行距離を保存`);
  alert("日次走行距離を保存しました");

  await loadDailyReports(els.dispatchDate?.value || todayStr());
  renderDailyMileageInputs();
  renderHomeMonthlyVehicleList();
  renderVehiclesTable();
}

async function syncDateAndReloadFromDispatchDate() {
  const dateStr = els.dispatchDate?.value || todayStr();
  if (els.planDate) els.planDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;
  syncMileageReportRange(dateStr, true);

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
  renderDailyDispatchResult();
}

async function syncDateAndReloadFromPlanDate() {
  const dateStr = els.planDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.actualDate) els.actualDate.value = dateStr;
  syncMileageReportRange(dateStr, true);

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
  renderManualLastVehicleInfo();
}

async function syncDateAndReloadFromActualDate() {
  const dateStr = els.actualDate?.value || todayStr();
  if (els.dispatchDate) els.dispatchDate.value = dateStr;
  if (els.planDate) els.planDate.value = dateStr;
  syncMileageReportRange(dateStr, true);

  await loadPlansByDate(dateStr);
  await loadActualsByDate(dateStr);
  await loadDailyReports(dateStr);
}

function bindPlanAndActualFormEvents() {
  if (els.planCastSelect) {
    els.planCastSelect.addEventListener("change", () => syncPlanFieldsFromCastInput(true));
    els.planCastSelect.addEventListener("input", () => {
      if (!findCastByInputValue(els.planCastSelect?.value || "")) clearLinkedCastSelection(els.planCastSelect);
      syncPlanFieldsFromCastInput(false);
    });
  }
  if (els.castSelect) {
    els.castSelect.addEventListener("change", () => syncActualFieldsFromCastInput(true));
    els.castSelect.addEventListener("input", () => {
      if (!findCastByInputValue(els.castSelect?.value || "")) clearLinkedCastSelection(els.castSelect);
      syncActualFieldsFromCastInput(false);
    });
  }
  if (els.planSelect) els.planSelect.addEventListener("change", fillActualFormFromSelectedPlan);
  if (els.cancelPlanEditBtn) els.cancelPlanEditBtn.addEventListener("click", resetPlanForm);
  if (els.cancelActualEditBtn) els.cancelActualEditBtn.addEventListener("click", resetActualForm);
  if (els.addSelectedPlanBtn) els.addSelectedPlanBtn.addEventListener("click", addPlanToActual);
  if (els.addFromPlansCloseBtn) els.addFromPlansCloseBtn.addEventListener("click", closeAddFromPlansDialog);
  if (els.addFromPlansCancelBtn) els.addFromPlansCancelBtn.addEventListener("click", closeAddFromPlansDialog);
  if (els.addFromPlansConfirmBtn) els.addFromPlansConfirmBtn.addEventListener("click", async () => {
    const checkedIds = Array.from(els.addFromPlansList?.querySelectorAll('input[data-plan-id]:checked') || [])
      .map(input => String(input.dataset.planId || "").trim())
      .filter(Boolean);

    if (!checkedIds.length) {
      alert("追加する予定を選択してください");
      return;
    }

    let addedCount = 0;
    for (const planId of checkedIds) {
      const plan = addFromPlansDialogCache.find(row => String(row?.id || "") === planId);
      const result = await addPlanToActualBySource(plan, { skipReload: true, silent: true });
      if (result?.ok) addedCount += 1;
    }

    closeAddFromPlansDialog();
    await loadActualsByDate(els.actualDate?.value || todayStr());
    await loadPlansByDate(els.planDate?.value || todayStr());
    if (els.planSelect) els.planSelect.value = "";
    renderPlanSelect();

    if (!addedCount) alert("追加できる予定がありませんでした");
  });
  if (els.addFromPlansDialog) {
    els.addFromPlansDialog.addEventListener("click", event => {
      if (event.target === els.addFromPlansDialog) closeAddFromPlansDialog();
    });
  }
}

function bindDispatchEvents() {
  if (els.optimizeBtn) els.optimizeBtn.addEventListener("click", () => (typeof window.runAutoDispatchV1 === "function" ? window.runAutoDispatchV1() : runAutoDispatch()));
  if (els.simulationSlotSelect) els.simulationSlotSelect.addEventListener("change", () => {
    if (suppressSimulationSlotChange || isRefreshingHybridUI) return;
    simulationSlotHour = Number(els.simulationSlotSelect.value || getOperationBaseHour());
    runSlotDiagnosisPreview();
  });
  if (els.simulationIncludePlanInflow) els.simulationIncludePlanInflow.addEventListener("change", () => {
    if (isRefreshingHybridUI) return;
    runSlotDiagnosisPreview();
  });
  if (els.runSimulationBtn) els.runSimulationBtn.addEventListener("click", runSlotDiagnosisPreview);
  if (els.runSimulationDispatchBtn) els.runSimulationDispatchBtn.addEventListener("click", runSimulationDispatchPreview);
}

let postDispatchEventsBound = false;
let dashboardInitialized = false;
let mileageSyncListenersBound = false;

function bindMileageReportSyncListeners() {
  if (mileageSyncListenersBound) return;
  mileageSyncListenersBound = true;

  const syncMileageReport = () => {
    syncMileageReportRange(els.dispatchDate?.value || todayStr(), true);
  };

  window.addEventListener("load", syncMileageReport, { once: true });
  window.addEventListener("pageshow", syncMileageReport, { once: true });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      syncMileageReport();
    }
  });
}

function bindPostDispatchEvents() {
  if (postDispatchEventsBound) return;
  postDispatchEventsBound = true;

  if (els.copyResultBtn) els.copyResultBtn.addEventListener("click", copyDispatchResult);
  if (els.confirmDailyBtn) els.confirmDailyBtn.addEventListener("click", confirmDailyToMonthly);
  if (els.clearActualBtn) els.clearActualBtn.addEventListener("click", clearAllActuals);
}


function getAdminTeamStatusLabel(status) {
  return String(status || 'active').trim() === 'suspended' ? '停止中' : '稼働中';
}

function toPlatformAdminMetricNumber(...candidates) {
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate.length;
    const numeric = Number(candidate);
    if (Number.isFinite(numeric)) return numeric;
  }
  return 0;
}

function getPlatformAdminRowMembers(row) {
  if (Array.isArray(row?.members)) return row.members;
  if (Array.isArray(row?.team_members)) return row.team_members;
  if (Array.isArray(row?.profiles)) return row.profiles;
  return [];
}

function getPlatformAdminMetricSnapshot(row = {}) {
  const members = getPlatformAdminRowMembers(row);
  return {
    membersCount: toPlatformAdminMetricNumber(row?.members_count, row?.member_count, row?.users_count, row?.profiles_count, members),
    vehiclesCount: toPlatformAdminMetricNumber(row?.vehicles_count, row?.vehicle_count, row?.cars_count, row?.total_vehicles, row?.vehicles),
    castsCount: toPlatformAdminMetricNumber(row?.casts_count, row?.cast_count, row?.people_count, row?.total_casts, row?.casts, row?.people),
    originsCount: toPlatformAdminMetricNumber(row?.origins_count, row?.origin_count, row?.bases_count, row?.origins),
    dispatchesCount: toPlatformAdminMetricNumber(row?.dispatches_count, row?.dispatch_count, row?.plans_count, row?.actuals_count, row?.dispatches),
    members
  };
}

function normalizePlatformAdminMetricLabel(value) {
  return String(value || '').replace(/\s+/g, '').trim().toLowerCase();
}

function findPlatformAdminMetricValueElement(root, label, options = {}) {
  if (!root) return null;
  const target = normalizePlatformAdminMetricLabel(label);
  if (!target) return null;
  const preferText = options.preferText === true;
  const leaves = Array.from(root.querySelectorAll('*')).filter(node => node && node.children.length === 0);
  const labelEl = leaves.find(node => normalizePlatformAdminMetricLabel(node.textContent) === target);
  if (!labelEl) return null;

  let current = labelEl.parentElement;
  for (let depth = 0; depth < 5 && current; depth += 1) {
    const branchLeaves = Array.from(current.querySelectorAll('*')).filter(node => node && node.children.length === 0);
    const candidates = branchLeaves.filter(node => node !== labelEl && normalizePlatformAdminMetricLabel(node.textContent) !== target);
    if (preferText) {
      const textCandidate = candidates.find(node => String(node.textContent || '').trim());
      if (textCandidate) return textCandidate;
    } else {
      const numericCandidate = candidates.find(node => /^[-+]?\d+(?:\.\d+)?(?:km|日|件|人|台)?$/i.test(String(node.textContent || '').trim()));
      if (numericCandidate) return numericCandidate;
      const fallbackCandidate = candidates.find(node => String(node.textContent || '').trim());
      if (fallbackCandidate) return fallbackCandidate;
    }
    current = current.parentElement;
  }
  return null;
}

function setPlatformAdminMetricByLabel(root, label, value, options = {}) {
  const node = findPlatformAdminMetricValueElement(root, label, options);
  if (node) node.textContent = String(value ?? '');
}

function updatePlatformAdminSummaryCards(rows = []) {
  const root = document.getElementById('platformAdminTab') || document;
  const list = Array.isArray(rows) ? rows : [];
  const snapshots = list.map(getPlatformAdminMetricSnapshot);
  const totalTeams = list.length;
  const visibleTeams = list.length;
  const suspendedTeams = list.filter(row => String(row?.status || 'active').trim() === 'suspended').length;
  const visibleMembers = snapshots.reduce((sum, item) => sum + Number(item.membersCount || 0), 0);
  const visibleVehicles = snapshots.reduce((sum, item) => sum + Number(item.vehiclesCount || 0), 0);
  const visibleCasts = snapshots.reduce((sum, item) => sum + Number(item.castsCount || 0), 0);

  setPlatformAdminMetricByLabel(root, '全チーム', totalTeams);
  setPlatformAdminMetricByLabel(root, '表示中', visibleTeams);
  setPlatformAdminMetricByLabel(root, '停止中', suspendedTeams);
  setPlatformAdminMetricByLabel(root, '表示メンバー', visibleMembers);
  setPlatformAdminMetricByLabel(root, '表示車両', visibleVehicles);
  setPlatformAdminMetricByLabel(root, '表示送り先', visibleCasts);
}

function updatePlatformAdminRoleCards(data = {}) {
  const root = els.platformAdminDetailContent || document.getElementById('platformAdminTab') || document;
  const members = getPlatformAdminRowMembers(data);
  const ownerCount = members.filter(row => String(row?.role || '').trim() === 'owner').length;
  const adminCount = members.filter(row => String(row?.role || '').trim() === 'admin').length;
  const userCount = members.filter(row => String(row?.role || '').trim() === 'user').length;
  const inactiveCount = members.filter(row => row?.is_active === false).length;
  const updatedLabel = formatDateTimeLabel(data?.updated_at || data?.created_at || '-');

  setPlatformAdminMetricByLabel(root, 'owner', ownerCount);
  setPlatformAdminMetricByLabel(root, 'admin', adminCount);
  setPlatformAdminMetricByLabel(root, 'user', userCount);
  setPlatformAdminMetricByLabel(root, '無効', inactiveCount);
  setPlatformAdminMetricByLabel(root, '更新', updatedLabel, { preferText: true });
}

async function hydratePlatformAdminTeamMetrics(rows = []) {
  const list = Array.isArray(rows) ? rows : [];
  const summaryFn = window.getTeamSummaryForAdmin;
  if (typeof summaryFn !== 'function') return list;

  const hydrated = [];
  for (const row of list) {
    const snapshot = getPlatformAdminMetricSnapshot(row);
    const hasCounts = snapshot.membersCount > 0 || snapshot.vehiclesCount > 0 || snapshot.castsCount > 0 || snapshot.originsCount > 0 || snapshot.dispatchesCount > 0;
    const teamId = String(row?.id || '').trim();
    if (!teamId || hasCounts) {
      hydrated.push(row);
      continue;
    }
    try {
      const { data, error } = await summaryFn(teamId);
      if (!error && data) hydrated.push({ ...row, ...data });
      else hydrated.push(row);
    } catch (_) {
      hydrated.push(row);
    }
  }
  return hydrated;
}

function renderPlatformAdminMembers(rows) {
  if (!els.platformAdminMembersTableBody) return;
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    els.platformAdminMembersTableBody.innerHTML = '<tr><td colspan="4" class="soft-text">メンバーはまだいません。</td></tr>';
    return;
  }
  els.platformAdminMembersTableBody.innerHTML = list.map(row => `
    <tr>
      <td>${escapeHtml(row.display_name || row.email || 'ユーザー')}</td>
      <td>${escapeHtml(row.email || '-')}</td>
      <td>${escapeHtml(getRoleLabel(row.role))}</td>
      <td>${row.is_active === false ? '無効' : '有効'}</td>
    </tr>
  `).join('');
}


let orphanAuthCandidatesCache = [];
let orphanAuthCandidatesLoadingToken = 0;

function getOrphanCleanupDomRefs() {
  const refreshBtn = document.getElementById('refreshPlatformAdminOrphansBtn')
    || Array.from(document.querySelectorAll('button')).find(btn => String(btn?.textContent || '').replace(/\s+/g, '').includes('候補更新'))
    || null;
  const statusEl = document.getElementById('platformAdminOrphanStatusText') || null;
  const tbody = document.getElementById('platformAdminOrphanTableBody') || null;

  if (refreshBtn && tbody) {
    const table = tbody.closest('table') || null;
    const section = document.getElementById('platformAdminOrphanWrap') || table?.closest('section,div') || null;
    return { section, refreshBtn, table, tbody, statusEl };
  }

  if (!refreshBtn) return null;

  let section = refreshBtn.parentElement;
  while (section && section !== document.body) {
    const text = String(section.textContent || '');
    if (text.includes('所属なしユーザー整理') && section.querySelector('table')) break;
    section = section.parentElement;
  }
  if (!section || section === document.body) return null;

  const table = section.querySelector('table');
  const fallbackTbody = table?.querySelector('tbody') || null;
  if (!fallbackTbody) return null;

  let fallbackStatus = Array.from(section.querySelectorAll('p, div, span')).find(node => {
    if (!node || node === section) return false;
    if (node.contains(table) || node.contains(refreshBtn)) return false;
    const text = String(node.textContent || '').trim();
    return text.includes('候補を') || text.includes('所属なしユーザー') || text.includes('完全削除');
  }) || null;

  if (!fallbackStatus) {
    fallbackStatus = document.createElement('div');
    fallbackStatus.className = 'soft-text';
    table.parentElement?.insertBefore(fallbackStatus, table);
  }

  return { section, refreshBtn, table, tbody: fallbackTbody, statusEl: fallbackStatus };
}

function renderOrphanCleanupCandidates(rows = [], message = '') {
  const refs = getOrphanCleanupDomRefs();
  if (!refs) return;
  const { tbody, statusEl } = refs;
  const list = Array.isArray(rows) ? rows : [];

  if (statusEl) {
    statusEl.textContent = message || (list.length
      ? `所属なしユーザー候補 ${list.length}件を表示しています。`
      : '所属なしユーザーは見つかりませんでした。');
  }

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="soft-text">所属なしユーザーは見つかりませんでした。</td></tr>';
    return;
  }

  tbody.innerHTML = list.map(row => {
    const userId = String(row?.user_id || '').trim();
    const email = String(row?.email || '').trim();
    const disabled = row?.is_current_user || row?.is_platform_admin ? 'disabled' : '';
    return `
      <tr data-orphan-user-id="${escapeHtml(userId)}" data-orphan-email="${escapeHtml(email)}">
        <td>${escapeHtml(row?.display_name || email || 'ユーザー')}</td>
        <td>${escapeHtml(email || '-')}</td>
        <td>${escapeHtml(row?.reason || '所属チームなし')}</td>
        <td>${escapeHtml(formatDateTimeLabel(row?.updated_at || row?.created_at))}</td>
        <td class="action-row">
          <button class="btn danger orphan-user-delete-btn" ${disabled}>完全削除</button>
        </td>
      </tr>`;
  }).join('');
}

async function fetchAllRowsWithoutTeamScope(tableName) {
  const safeTable = String(tableName || '').trim();
  if (!safeTable) return [];

  const pageSize = 1000;
  let from = 0;
  const rows = [];

  while (true) {
    const { data, error } = await supabaseClient
      .from(safeTable)
      .select('*')
      .range(from, from + pageSize - 1);

    if (error) {
      if ((typeof isMissingTableError === 'function' && isMissingTableError(error)) || String(error?.code || '') === 'PGRST205') {
        return [];
      }
      throw error;
    }

    const chunk = Array.isArray(data) ? data : [];
    rows.push(...chunk);
    if (chunk.length < pageSize) break;
    from += pageSize;
  }

  return rows;
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 12000) {
  const safeTimeout = Math.max(3000, Number(timeoutMs || 12000));
  const controller = typeof AbortController === 'function' ? new AbortController() : null;
  const timer = controller ? window.setTimeout(() => controller.abort(), safeTimeout) : null;
  try {
    return await fetch(url, { ...(options || {}), signal: controller?.signal });
  } catch (error) {
    if (error?.name === 'AbortError') {
      throw new Error(`通信がタイムアウトしました (${Math.round(safeTimeout / 1000)}秒)`);
    }
    throw error;
  } finally {
    if (timer) window.clearTimeout(timer);
  }
}

async function listAuthUsersViaQuickHandler() {
  const sessionResult = await supabaseClient.auth.getSession();
  const accessToken = String(sessionResult?.data?.session?.access_token || '').trim();
  if (!accessToken) {
    throw new Error('ログインセッションを取得できませんでした。');
  }

  const response = await fetchWithTimeout(`${SUPABASE_URL}/functions/v1/quick-handler`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${accessToken}`,
      apikey: SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ action: 'list' })
  }, 12000);

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(String(payload?.error || `Auth候補一覧の取得に失敗しました (${response.status})`));
  }

  return Array.isArray(payload?.users) ? payload.users : [];
}

async function buildOrphanAuthCandidates() {
  const [authUsers, teamMembers, profiles, teams] = await Promise.all([
    listAuthUsersViaQuickHandler(),
    fetchAllRowsWithoutTeamScope(getTableName('team_members')),
    fetchAllRowsWithoutTeamScope(getTableName('profiles')),
    fetchAllRowsWithoutTeamScope(getTableName('teams'))
  ]);

  const validTeamIds = new Set(
    (Array.isArray(teams) ? teams : [])
      .map(row => String(row?.id || row?.team_id || '').trim())
      .filter(Boolean)
  );

  const activeMembershipRows = (Array.isArray(teamMembers) ? teamMembers : []).filter(row => {
    const teamId = String(row?.team_id || '').trim();
    return teamId && validTeamIds.has(teamId);
  });

  const memberUserIds = new Set(
    activeMembershipRows
      .map(row => String(row?.user_id || '').trim())
      .filter(Boolean)
  );
  const memberEmails = new Set(
    activeMembershipRows
      .flatMap(row => [row?.email, row?.member_email])
      .map(value => String(value || '').trim().toLowerCase())
      .filter(Boolean)
  );

  const profileByUserId = new Map();
  const profileByEmail = new Map();
  (Array.isArray(profiles) ? profiles : []).forEach(row => {
    const userId = String(row?.user_id || row?.id || '').trim();
    const email = String(row?.email || '').trim().toLowerCase();
    if (userId && !profileByUserId.has(userId)) profileByUserId.set(userId, row);
    if (email && !profileByEmail.has(email)) profileByEmail.set(email, row);
  });

  const currentUserId = String(currentUser?.id || currentUserProfile?.id || '').trim();
  const currentUserEmail = String(currentUser?.email || currentUserProfile?.email || '').trim().toLowerCase();

  return (Array.isArray(authUsers) ? authUsers : []).map(user => {
    const userId = String(user?.id || '').trim();
    const email = String(user?.email || '').trim().toLowerCase();
    const profile = profileByUserId.get(userId) || profileByEmail.get(email) || null;
    const isPlatformAdmin = user?.is_admin === true || profile?.role === 'platform_admin' || profile?.is_platform_admin === true;
    const hasTeamMembership = (userId && memberUserIds.has(userId)) || (email && memberEmails.has(email));
    return {
      user_id: userId,
      email,
      display_name: String(profile?.display_name || profile?.name || email || 'ユーザー').trim() || 'ユーザー',
      created_at: user?.created_at || profile?.created_at || null,
      updated_at: profile?.updated_at || profile?.created_at || user?.created_at || null,
      reason: hasTeamMembership ? '' : '所属チームなし',
      is_platform_admin: isPlatformAdmin,
      is_current_user: (userId && userId === currentUserId) || (email && email === currentUserEmail),
      has_team_membership: hasTeamMembership
    };
  }).filter(row => {
    if (!row.user_id && !row.email) return false;
    if (row.is_platform_admin) return false;
    return !row.has_team_membership;
  }).sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')));
}

async function loadOrphanCleanupCandidates(force = false) {
  const refs = getOrphanCleanupDomRefs();
  if (!refs) return;
  const token = ++orphanAuthCandidatesLoadingToken;
  refs.refreshBtn.disabled = true;
  if (refs.statusEl) refs.statusEl.textContent = '候補を読込中です...';
  refs.tbody.innerHTML = '<tr><td colspan="5" class="soft-text">候補を読込中です...</td></tr>';

  try {
    const rows = await buildOrphanAuthCandidates();
    if (token !== orphanAuthCandidatesLoadingToken) return;
    orphanAuthCandidatesCache = rows;
    renderOrphanCleanupCandidates(rows);
  } catch (error) {
    console.error('loadOrphanCleanupCandidates error:', error);
    if (token !== orphanAuthCandidatesLoadingToken) return;
    orphanAuthCandidatesCache = [];
    renderOrphanCleanupCandidates([], `候補取得に失敗しました: ${error?.message || error}`);
  } finally {
    const latestRefs = getOrphanCleanupDomRefs();
    if (latestRefs) latestRefs.refreshBtn.disabled = false;
  }
}

async function deleteOrphanCleanupCandidate(userId, email) {
  const safeUserId = String(userId || '').trim();
  const safeEmail = String(email || '').trim().toLowerCase();
  if (!safeUserId && !safeEmail) {
    alert('削除対象ユーザーを特定できませんでした。');
    return;
  }

  const label = safeEmail || safeUserId;
  if (!window.confirm(`${label} を完全削除しますか？\nこの操作は元に戻せません。`)) return;

  const sessionResult = await supabaseClient.auth.getSession();
  const accessToken = String(sessionResult?.data?.session?.access_token || '').trim();
  if (!accessToken) {
    alert('ログインセッションを取得できませんでした。');
    return;
  }

  const response = await fetch(`${SUPABASE_URL}/functions/v1/quick-handler`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${accessToken}`,
      apikey: SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ action: 'delete', user_id: safeUserId || undefined, email: safeEmail || undefined })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    alert(String(payload?.error || `完全削除に失敗しました (${response.status})`));
    return;
  }

  orphanAuthCandidatesCache = orphanAuthCandidatesCache.filter(row => {
    const sameId = safeUserId && String(row?.user_id || '').trim() === safeUserId;
    const sameEmail = safeEmail && String(row?.email || '').trim().toLowerCase() === safeEmail;
    return !(sameId || sameEmail);
  });
  renderOrphanCleanupCandidates(orphanAuthCandidatesCache, `${label} を完全削除しました。`);
}

function bindOrphanCleanupEvents() {
  const refs = getOrphanCleanupDomRefs();
  if (!refs) return;
  if (refs.refreshBtn && refs.refreshBtn.dataset.bound !== '1') {
    refs.refreshBtn.dataset.bound = '1';
    refs.refreshBtn.addEventListener('click', () => loadOrphanCleanupCandidates(true));
  }
  if (refs.tbody && refs.tbody.dataset.bound !== '1') {
    refs.tbody.dataset.bound = '1';
    refs.tbody.addEventListener('click', async event => {
      const deleteBtn = event.target.closest('.orphan-user-delete-btn');
      if (!deleteBtn) return;
      const tr = deleteBtn.closest('tr[data-orphan-user-id], tr[data-orphan-email]');
      if (!tr) return;
      await deleteOrphanCleanupCandidate(tr.dataset.orphanUserId || '', tr.dataset.orphanEmail || '');
    });
  }
}

function resetPlatformAdminDetailView() {
  selectedPlatformAdminTeamId = null;
  if (els.platformAdminDetailEmpty) els.platformAdminDetailEmpty.classList.remove('hidden');
  if (els.platformAdminDetailContent) els.platformAdminDetailContent.classList.add('hidden');
  if (els.platformAdminMembersTableBody) els.platformAdminMembersTableBody.innerHTML = '';
  if (els.switchPlatformTeamBtn) {
    els.switchPlatformTeamBtn.disabled = true;
    els.switchPlatformTeamBtn.dataset.teamId = '';
    els.switchPlatformTeamBtn.dataset.teamName = '';
    els.switchPlatformTeamBtn.textContent = 'このチームへ切替';
  }
  if (els.platformAdminTeamNameInput) {
    els.platformAdminTeamNameInput.value = '';
    els.platformAdminTeamNameInput.disabled = true;
    els.platformAdminTeamNameInput.dataset.teamId = '';
  }
  if (els.savePlatformAdminTeamNameBtn) {
    els.savePlatformAdminTeamNameBtn.disabled = true;
    els.savePlatformAdminTeamNameBtn.dataset.teamId = '';
  }
  setPlatformAdminTeamNameStatus('チームを選ぶと編集できます。', false);
  if (els.suspendPlatformTeamBtn) els.suspendPlatformTeamBtn.disabled = true;
  if (els.resumePlatformTeamBtn) els.resumePlatformTeamBtn.disabled = true;
  if (els.exportPlatformTeamBackupBtn) els.exportPlatformTeamBackupBtn.disabled = true;
  if (els.importPlatformTeamBackupBtn) els.importPlatformTeamBackupBtn.disabled = true;
  if (els.deletePlatformTeamBtn) els.deletePlatformTeamBtn.disabled = true;
  if (els.platformAdminDetailPlanType) els.platformAdminDetailPlanType.value = '-';
  if (els.platformAdminDetailPlanLimits) els.platformAdminDetailPlanLimits.value = '-';
  if (els.platformAdminDetailPlanFeatures) els.platformAdminDetailPlanFeatures.value = '-';
  if (els.platformAdminSetFreePlanBtn) {
    els.platformAdminSetFreePlanBtn.disabled = true;
    els.platformAdminSetFreePlanBtn.dataset.teamId = '';
  }
  if (els.platformAdminSetPaidPlanBtn) {
    els.platformAdminSetPaidPlanBtn.disabled = true;
    els.platformAdminSetPaidPlanBtn.dataset.teamId = '';
  }
  if (els.platformAdminPlanStatusText) els.platformAdminPlanStatusText.textContent = 'チームを選ぶと切替できます。';
  if (els.platformAdminDetailFlags) els.platformAdminDetailFlags.innerHTML = '';
}

function setPlatformAdminPlanStatus(message, isError = false) {
  if (!els.platformAdminPlanStatusText) return;
  els.platformAdminPlanStatusText.textContent = String(message || '').trim() || 'チームを選ぶと切替できます。';
  els.platformAdminPlanStatusText.style.color = isError ? '#ffb4b4' : '';
}

function formatAdminPlanLimitSummary(plan) {
  const parts = [
    `人数 ${formatPlanLimitValue(plan?.limits?.members)}名`,
    `起点 ${formatPlanLimitValue(plan?.limits?.origins)}件`,
    `車両 ${formatPlanLimitValue(plan?.limits?.vehicles)}台`,
    `送り先 ${formatPlanLimitValue(plan?.limits?.casts)}人`
  ];
  return parts.join(' / ');
}

function formatAdminPlanFeatureSummary(plan) {
  const flags = plan?.feature_flags || {};
  const yesNo = value => value ? '可' : '不可';
  return [
    `CSV ${yesNo(flags.csv === true)}`,
    `API ${yesNo(flags.google_api_dispatch === true)}`,
    `バックアップ ${yesNo(flags.backup === true || flags.restore === true)}`
  ].join(' / ');
}

function renderPlatformAdminPlanFlags(plan) {
  if (!els.platformAdminDetailFlags) return;
  const chips = [
    { label: getPlanTypeLabel(plan?.plan_type), className: String(plan?.plan_type || 'free') === 'paid' ? 'success' : 'warning' },
    { label: `CSV ${plan?.feature_flags?.csv === true ? '可' : '不可'}`, className: plan?.feature_flags?.csv === true ? 'success' : 'warning' },
    { label: `GoogleMap API ${plan?.feature_flags?.google_api_dispatch === true ? '可' : '不可'}`, className: plan?.feature_flags?.google_api_dispatch === true ? 'success' : 'warning' },
    { label: `バックアップ ${plan?.feature_flags?.backup === true || plan?.feature_flags?.restore === true ? '可' : '不可'}`, className: plan?.feature_flags?.backup === true || plan?.feature_flags?.restore === true ? 'success' : 'warning' }
  ];
  els.platformAdminDetailFlags.innerHTML = chips.map(item => `<span class="platform-admin-flag ${escapeHtml(item.className)}">${escapeHtml(item.label)}</span>`).join('');
}

function renderPlatformAdminPlanSection(teamData = {}) {
  const plan = normalizeDropOffPlanRecord(teamData || {});
  const planType = String(plan?.plan_type || 'free') === 'paid' ? 'paid' : 'free';
  if (els.platformAdminDetailPlanType) els.platformAdminDetailPlanType.value = getPlanTypeLabel(planType);
  if (els.platformAdminDetailPlanLimits) els.platformAdminDetailPlanLimits.value = formatAdminPlanLimitSummary(plan);
  if (els.platformAdminDetailPlanFeatures) els.platformAdminDetailPlanFeatures.value = formatAdminPlanFeatureSummary(plan);
  if (els.platformAdminSetFreePlanBtn) {
    els.platformAdminSetFreePlanBtn.dataset.teamId = String(teamData?.id || selectedPlatformAdminTeamId || '');
    els.platformAdminSetFreePlanBtn.disabled = !teamData?.id || planType === 'free';
    els.platformAdminSetFreePlanBtn.textContent = planType === 'free' ? 'free 適用中' : 'free にする';
  }
  if (els.platformAdminSetPaidPlanBtn) {
    els.platformAdminSetPaidPlanBtn.dataset.teamId = String(teamData?.id || selectedPlatformAdminTeamId || '');
    els.platformAdminSetPaidPlanBtn.disabled = !teamData?.id || planType === 'paid';
    els.platformAdminSetPaidPlanBtn.textContent = planType === 'paid' ? 'Paid 適用中' : 'Paid にする';
  }
  setPlatformAdminPlanStatus(`現在は${getPlanTypeLabel(planType)}です。ここで free / Paid を切り替えできます。`, false);
  renderPlatformAdminPlanFlags(plan);
}

async function renderPlatformAdminDetail(teamId) {
  if (!window.isPlatformAdminUser || !teamId) return;
  const fn = window.getTeamSummaryForAdmin;
  if (typeof fn !== 'function') return;
  const { data, error } = await fn(teamId);
  if (error || !data) {
    if (els.platformAdminStatusText) els.platformAdminStatusText.textContent = error?.message || 'チーム詳細の取得に失敗しました。';
    return;
  }
  selectedPlatformAdminTeamId = teamId;
  if (els.platformAdminDetailEmpty) els.platformAdminDetailEmpty.classList.add('hidden');
  if (els.platformAdminDetailContent) els.platformAdminDetailContent.classList.remove('hidden');
  const metrics = getPlatformAdminMetricSnapshot(data);
  if (els.platformAdminDetailTeamName) els.platformAdminDetailTeamName.textContent = data.team_name || '-';
  if (els.platformAdminDetailStatus) els.platformAdminDetailStatus.textContent = getAdminTeamStatusLabel(data.status);
  if (els.platformAdminDetailMembersCount) els.platformAdminDetailMembersCount.textContent = String(metrics.membersCount || 0);
  if (els.platformAdminDetailVehiclesCount) els.platformAdminDetailVehiclesCount.textContent = String(metrics.vehiclesCount || 0);
  if (els.platformAdminDetailCastsCount) els.platformAdminDetailCastsCount.textContent = String(metrics.castsCount || 0);
  if (els.platformAdminDetailOriginsCount) els.platformAdminDetailOriginsCount.textContent = String(metrics.originsCount || 0);
  if (els.platformAdminDetailDispatchesCount) els.platformAdminDetailDispatchesCount.textContent = String(metrics.dispatchesCount || 0);
  updatePlatformAdminRoleCards(data);
  const forcedTeamId = getAdminForcedTeamId();
  const isCurrentForced = String(forcedTeamId || '') === String(teamId || '');
  if (els.platformAdminDetailMeta) {
    const ownerEmail = Array.isArray(data.members) ? (data.members.find(x => String(x.role || '') === 'owner')?.email || '') : '';
    const reason = data.status === 'suspended' && data.suspended_reason ? ` / 理由: ${data.suspended_reason}` : '';
    const forcedLabel = isCurrentForced ? ' / 強制切替中' : '';
    els.platformAdminDetailMeta.textContent = `owner: ${ownerEmail || '-'} / 作成日: ${formatDateTimeLabel(data.created_at)}${reason}${forcedLabel}`;
  }
  if (els.platformAdminTeamNameInput) {
    els.platformAdminTeamNameInput.disabled = false;
    els.platformAdminTeamNameInput.dataset.teamId = String(teamId || '');
    els.platformAdminTeamNameInput.value = String(data.team_name || '').trim();
  }
  if (els.savePlatformAdminTeamNameBtn) {
    els.savePlatformAdminTeamNameBtn.disabled = false;
    els.savePlatformAdminTeamNameBtn.dataset.teamId = String(teamId || '');
  }
  setPlatformAdminTeamNameStatus('変更後に保存すると、このチームの表示名が全画面で更新されます。', false);
  if (els.switchPlatformTeamBtn) {
    els.switchPlatformTeamBtn.disabled = isCurrentForced;
    els.switchPlatformTeamBtn.textContent = isCurrentForced ? '切替中' : 'このチームへ切替';
    els.switchPlatformTeamBtn.dataset.teamId = String(teamId || '');
    els.switchPlatformTeamBtn.dataset.teamName = String(data.team_name || '');
  }
  if (els.suspendPlatformTeamBtn) els.suspendPlatformTeamBtn.disabled = String(data.status || 'active') === 'suspended';
  if (els.resumePlatformTeamBtn) els.resumePlatformTeamBtn.disabled = String(data.status || 'active') !== 'suspended';
  if (els.exportPlatformTeamBackupBtn) {
    els.exportPlatformTeamBackupBtn.disabled = !teamId;
    els.exportPlatformTeamBackupBtn.dataset.teamId = String(teamId || '');
    els.exportPlatformTeamBackupBtn.dataset.teamName = String(data.team_name || '');
  }
  if (els.importPlatformTeamBackupBtn) {
    els.importPlatformTeamBackupBtn.disabled = !teamId;
    els.importPlatformTeamBackupBtn.dataset.teamId = String(teamId || '');
    els.importPlatformTeamBackupBtn.dataset.teamName = String(data.team_name || '');
  }
  if (els.importPlatformTeamBackupInput) {
    els.importPlatformTeamBackupInput.dataset.teamId = String(teamId || '');
    els.importPlatformTeamBackupInput.dataset.teamName = String(data.team_name || '');
  }
  if (els.deletePlatformTeamBtn) {
    els.deletePlatformTeamBtn.disabled = !teamId;
    els.deletePlatformTeamBtn.dataset.teamId = String(teamId || '');
    els.deletePlatformTeamBtn.dataset.teamName = String(data.team_name || '');
  }
  renderPlatformAdminPlanSection({ ...(data || {}), id: teamId });
  setPlatformAdminActionStatus(isCurrentForced ? '強制切替中です。選択中チームに対して、強制切替以外の管理操作を実行できます。' : '選択中チームに対して、強制切替 / 無料有料切替 / バックアップ / 復元 / チーム削除を実行できます。', false);
  renderPlatformAdminMembers(data.members || []);
  bindOrphanCleanupEvents();
  Promise.resolve()
    .then(() => loadOrphanCleanupCandidates())
    .catch(error => console.error('loadOrphanCleanupCandidates deferred error:', error));
}

function renderPlatformAdminTeams() {
  if (!els.platformAdminTeamsTableBody) return;
  const rows = Array.isArray(platformAdminTeamsCache) ? platformAdminTeamsCache : [];
  if (!rows.length) {
    els.platformAdminTeamsTableBody.innerHTML = '<tr><td colspan="8" class="soft-text">チームはまだありません。</td></tr>';
    return;
  }
  els.platformAdminTeamsTableBody.innerHTML = rows.map(row => {
    const teamId = String(row.id || '');
    const disabledSuspend = String(row.status || 'active') === 'suspended' ? 'disabled' : '';
    const disabledResume = String(row.status || 'active') === 'suspended' ? '' : 'disabled';
    const metrics = getPlatformAdminMetricSnapshot(row);
    const plan = normalizeDropOffPlanRecord(row || {});
    return `
      <tr data-admin-team-id="${escapeHtml(teamId)}">
        <td>${escapeHtml(row.team_name || '-')}</td>
        <td>${escapeHtml(getAdminTeamStatusLabel(row.status))}</td>
        <td>${escapeHtml(getPlanTypeLabel(plan?.plan_type))}</td>
        <td>${escapeHtml(row.owner_email || row.owner_display_name || '-')}</td>
        <td>${escapeHtml(String(metrics.vehiclesCount ?? 0))}</td>
        <td>${escapeHtml(String(metrics.castsCount ?? 0))}</td>
        <td>${escapeHtml(formatDateTimeLabel(row.updated_at || row.created_at))}</td>
        <td>
          <div class="inline-actions">
            <button type="button" class="btn ghost admin-team-detail-btn">詳細</button>
            <button type="button" class="btn ghost admin-team-switch-btn" ${String(getAdminForcedTeamId() || '') === teamId ? 'disabled' : ''}>切替</button>
            <button type="button" class="btn danger admin-team-suspend-btn" ${disabledSuspend}>停止</button>
            <button type="button" class="btn primary admin-team-resume-btn" ${disabledResume}>再開</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
  updatePlatformAdminSummaryCards(rows);
}


function formatPlatformAdminAccessDateLabel(dateString) {
  const value = String(dateString || '').slice(0, 10);
  if (!value) return '-';
  return value.slice(5).replace('-', '/');
}

function renderPlatformAdminAnalyticsDailyTable(seriesHome = [], seriesDashboard = []) {
  if (!els.platformAdminAnalyticsDailyBody) return;
  const homeMap = new Map((Array.isArray(seriesHome) ? seriesHome : []).map(row => [String(row?.date || ''), Number(row?.count || 0)]));
  const dashboardMap = new Map((Array.isArray(seriesDashboard) ? seriesDashboard : []).map(row => [String(row?.date || ''), Number(row?.count || 0)]));
  const labels = Array.from(new Set([
    ...homeMap.keys(),
    ...dashboardMap.keys()
  ].filter(Boolean))).sort().slice(-7).reverse();

  if (!labels.length) {
    els.platformAdminAnalyticsDailyBody.innerHTML = '<tr><td colspan="3" class="soft-text">まだアクセスがありません。</td></tr>';
    return;
  }

  els.platformAdminAnalyticsDailyBody.innerHTML = labels.map(label => {
    const homeCount = Number(homeMap.get(label) || 0);
    const dashboardCount = Number(dashboardMap.get(label) || 0);
    return `
      <tr>
        <td>${escapeHtml(formatPlatformAdminAccessDateLabel(label))}</td>
        <td>${homeCount}</td>
        <td>${dashboardCount}</td>
      </tr>
    `;
  }).join('');
}

async function loadPlatformAdminAccessAnalytics() {
  if (!window.isPlatformAdminUser) return;
  if (!els.platformAdminAnalyticsStatusText) return;
  if (typeof window.fetchAccessAnalytics !== 'function') {
    els.platformAdminAnalyticsStatusText.textContent = 'アクセス解析機能を読み込めませんでした。';
    return;
  }

  els.platformAdminAnalyticsStatusText.textContent = 'アクセス解析を読み込んでいます...';
  const { data, error } = await window.fetchAccessAnalytics({ days: 30 });
  if (error || !data) {
    const message = typeof window.describeAccessAnalyticsError === 'function'
      ? window.describeAccessAnalyticsError(error)
      : (error?.message || 'アクセス解析の読込に失敗しました。');
    els.platformAdminAnalyticsStatusText.textContent = message;
    if (els.platformAdminHomeChart) els.platformAdminHomeChart.innerHTML = '<div class="platform-admin-chart-empty">データを表示できません</div>';
    if (els.platformAdminDashboardChart) els.platformAdminDashboardChart.innerHTML = '<div class="platform-admin-chart-empty">データを表示できません</div>';
    if (els.platformAdminAnalyticsDailyBody) els.platformAdminAnalyticsDailyBody.innerHTML = '<tr><td colspan="3" class="soft-text">データを表示できません。</td></tr>';
    return;
  }

  const totals = data.totals || {};
  if (els.platformAdminHomeTodayCount) els.platformAdminHomeTodayCount.textContent = String(totals.home_today || 0);
  if (els.platformAdminDashboardTodayCount) els.platformAdminDashboardTodayCount.textContent = String(totals.dashboard_today || 0);
  if (els.platformAdminHome30dCount) els.platformAdminHome30dCount.textContent = String(totals.home_30d || 0);
  if (els.platformAdminDashboard30dCount) els.platformAdminDashboard30dCount.textContent = String(totals.dashboard_30d || 0);

  if (typeof window.renderAccessLineChart === 'function') {
    window.renderAccessLineChart(els.platformAdminHomeChart, data.daily?.home || [], {
      color: '#4f85ff',
      fill: 'rgba(79,133,255,.16)'
    });
    window.renderAccessLineChart(els.platformAdminDashboardChart, data.daily?.dashboard || [], {
      color: '#7be2ab',
      fill: 'rgba(123,226,171,.16)'
    });
  }

  renderPlatformAdminAnalyticsDailyTable(data.daily?.home || [], data.daily?.dashboard || []);
  els.platformAdminAnalyticsStatusText.textContent = `今日 home ${totals.home_today || 0} / dashboard ${totals.dashboard_today || 0}、直近30日を表示中です。`;
}

async function loadPlatformAdminTeams() {
  if (!window.isPlatformAdminUser) return;
  const fn = window.getAllTeamsForAdmin;
  if (typeof fn !== 'function') return;
  if (els.platformAdminStatusText) els.platformAdminStatusText.textContent = '全チームを読み込んでいます...';
  const { data, error } = await fn();
  if (error) {
    if (els.platformAdminStatusText) els.platformAdminStatusText.textContent = error.message || '全チーム一覧の取得に失敗しました。';
    return;
  }
  platformAdminTeamsCache = Array.isArray(data) ? data : [];
  platformAdminTeamsCache = await hydratePlatformAdminTeamMetrics(platformAdminTeamsCache);
  renderPlatformAdminTeams();
  if (els.platformAdminStatusText) els.platformAdminStatusText.textContent = `${platformAdminTeamsCache.length}チームを表示しています。`;
  const forcedTeamId = String(getAdminForcedTeamId() || '').trim();
  const targetTeamId = forcedTeamId && platformAdminTeamsCache.some(row => String(row.id) === forcedTeamId)
    ? forcedTeamId
    : (selectedPlatformAdminTeamId && platformAdminTeamsCache.some(row => String(row.id) === String(selectedPlatformAdminTeamId))
      ? selectedPlatformAdminTeamId
      : String(platformAdminTeamsCache[0]?.id || ''));

  const analyticsPromise = Promise.resolve()
    .then(() => loadPlatformAdminAccessAnalytics())
    .catch(error => console.error('loadPlatformAdminAccessAnalytics error:', error));

  if (targetTeamId) await renderPlatformAdminDetail(targetTeamId);
  else {
    resetPlatformAdminDetailView();
    setPlatformAdminActionStatus('チームを選ぶと、強制切替 / バックアップ / 復元 / チーム削除の詳細が表示されます。', false);
  }

  await analyticsPromise;
}

async function changePlatformAdminTeamStatus(teamId, nextStatus) {
  if (!window.isPlatformAdminUser || !teamId) return;
  const fn = window.setTeamStatusForAdmin;
  if (typeof fn !== 'function') return;
  const status = String(nextStatus || 'active') === 'suspended' ? 'suspended' : 'active';
  let reason = '';
  if (status === 'suspended') {
    reason = window.prompt('停止理由を入力してください。', '運営者による停止') || '';
    if (!reason.trim()) return;
  }
  const { error } = await fn(teamId, status, reason);
  if (error) {
    alert(error.message || 'チーム状態の更新に失敗しました。');
    return;
  }
  await loadPlatformAdminTeams();
}

async function exportSelectedPlatformAdminTeamBackup() {
  const teamId = String(selectedPlatformAdminTeamId || els.exportPlatformTeamBackupBtn?.dataset?.teamId || '').trim();
  if (!teamId) {
    alert('バックアップ対象のチームを選択してください。');
    return;
  }
  const teamName = String(els.exportPlatformTeamBackupBtn?.dataset?.teamName || getPlatformAdminTeamLabel(teamId, 'team')).trim() || 'team';
  setPlatformAdminActionStatus(`${teamName} のバックアップを作成しています...`, false);
  await exportWorkspaceBackupByTeamId(teamId, {
    teamName,
    filePrefix: 'dropoff_team_backup',
    successStatus: `${teamName} のバックアップJSONを書き出しました。`
  });
}

function triggerImportPlatformTeamBackup() {
  const teamId = String(selectedPlatformAdminTeamId || '').trim();
  if (!teamId) {
    alert('復元対象のチームを選択してください。');
    return;
  }
  triggerImportAll(els.importPlatformTeamBackupInput);
}

async function importSelectedPlatformAdminTeamBackup() {
  const teamId = String(selectedPlatformAdminTeamId || els.importPlatformTeamBackupInput?.dataset?.teamId || '').trim();
  if (!teamId) {
    alert('復元対象のチームを選択してください。');
    return;
  }
  const teamName = String(els.importPlatformTeamBackupInput?.dataset?.teamName || getPlatformAdminTeamLabel(teamId, 'このチーム')).trim() || 'このチーム';
  setPlatformAdminActionStatus(`${teamName} の復元を実行しています...`, false);
  await importWorkspaceDataFromInput(els.importPlatformTeamBackupInput, {
    teamId,
    teamName,
    successMessage: `${teamName} の復元が完了しました。`,
    successStatus: `${teamName} の復元が完了しました。`
  });
}


async function deleteRowsByTeamIdSafeForAdmin(tableName, teamId) {
  const safeTable = String(tableName || '').trim();
  const safeTeamId = String(teamId || '').trim();
  if (!safeTable || !safeTeamId) return { error: null };

  try {
    const { error } = await supabaseClient
      .from(safeTable)
      .delete()
      .eq('team_id', safeTeamId);

    if (error) {
      const message = String(error?.message || '');
      const code = String(error?.code || '');
      const missing = (typeof isMissingTableError === 'function' && isMissingTableError(error)) ||
        code === 'PGRST205' || /Could not find the table/i.test(message) || /schema cache/i.test(message);
      if (missing) return { error: null };
      return { error };
    }

    return { error: null };
  } catch (error) {
    return { error };
  }
}

async function deleteTeamForAdminFallback(teamId) {
  const safeTeamId = String(teamId || '').trim();
  if (!safeTeamId) {
    return { error: new Error('削除対象のチームIDがありません。') };
  }

  const tableResolvers = [
    () => getDispatchUnifiedTableName(),
    () => getVehicleDailyRunsTableName(),
    () => getTableName('invitations'),
    () => getTableName('casts'),
    () => getTableName('vehicles'),
    () => getTableName('origins'),
    () => getTableName('team_members')
  ];

  for (const resolveTableName of tableResolvers) {
    let tableName = '';
    try {
      tableName = String(resolveTableName?.() || '').trim();
    } catch (_) {
      tableName = '';
    }
    if (!tableName) continue;

    const { error } = await deleteRowsByTeamIdSafeForAdmin(tableName, safeTeamId);
    if (error) {
      console.error('deleteTeamForAdminFallback delete failed:', tableName, error);
      return { error };
    }
  }

  try {
    const teamsTable = String(getTableName('teams') || '').trim();
    if (!teamsTable) {
      return { error: new Error('teams テーブル名を取得できませんでした。') };
    }

    const { error } = await supabaseClient
      .from(teamsTable)
      .delete()
      .eq('id', safeTeamId);

    if (error) {
      console.error('deleteTeamForAdminFallback team delete failed:', error);
      return { error };
    }
  } catch (error) {
    return { error };
  }

  try {
    platformAdminTeamsCache = (Array.isArray(platformAdminTeamsCache) ? platformAdminTeamsCache : []).filter(row => !isSameTeamId(row?.id, safeTeamId));
  } catch (_) {}

  return { error: null };
}

async function deleteSelectedPlatformAdminTeam() {
  const teamId = String(selectedPlatformAdminTeamId || els.deletePlatformTeamBtn?.dataset?.teamId || '').trim();
  if (!teamId) {
    alert('削除対象のチームを選択してください。');
    return;
  }
  const teamName = String(els.deletePlatformTeamBtn?.dataset?.teamName || getPlatformAdminTeamLabel(teamId, 'このチーム')).trim() || 'このチーム';
  const detail = await (typeof window.getTeamSummaryForAdmin === 'function' ? window.getTeamSummaryForAdmin(teamId) : { data: null, error: null });
  const summary = detail?.data || null;
  const countsText = summary
    ? `

削除対象
ユーザー: ${Number(summary.members_count || 0)}人
車両: ${Number(summary.vehicles_count || 0)}台
送り先: ${Number(summary.casts_count || 0)}人
起点: ${Number(summary.origins_count || 0)}件
Dispatch: ${Number(summary.dispatches_count || 0)}件`
    : '';

  if (!window.confirm(`${teamName} を削除しますか？
この操作は元に戻せません。${countsText}`)) return;

  const typedTeamName = String(window.prompt(`削除を続けるには、チーム名「${teamName}」をそのまま入力してください。`, '') || '').trim();
  if (typedTeamName !== teamName) {
    alert('チーム名が一致しなかったため、削除を中止しました。');
    return;
  }

  if (isSameTeamId(getAdminForcedTeamId(), teamId)) {
    const forcedOk = window.confirm(`このチームは現在「強制切替中」です。
削除すると運営者表示も通常モードへ戻ります。実行しますか？`);
    if (!forcedOk) return;
  }

  if (!window.confirm(`最終確認です。${teamName} の関連データとチーム本体を削除します。実行しますか？`)) return;

  const fn = typeof window.deleteTeamForAdmin === 'function'
    ? window.deleteTeamForAdmin
    : deleteTeamForAdminFallback;

  setPlatformAdminActionStatus(`${teamName} を削除しています...`, false);
  const { error } = await fn(teamId);
  if (error) {
    setPlatformAdminActionStatus(`${teamName} の削除に失敗しました: ${error.message || error}`, true);
    alert(error.message || 'チーム削除に失敗しました。');
    return;
  }

  const forcedTeamId = String(getAdminForcedTeamId() || '').trim();
  const wasForcedTeam = isSameTeamId(forcedTeamId, teamId);
  if (wasForcedTeam) {
    clearWorkspaceScopedRuntimeCaches();
    clearAdminForceTeamStorage();
    window.currentWorkspaceInfo = null;
  }
  if (isSameTeamId(selectedPlatformAdminTeamId, teamId)) {
    selectedPlatformAdminTeamId = null;
  }

  setPlatformAdminActionStatus(`${teamName} を削除しました。`, false);
  await loadPlatformAdminTeams();

  if (wasForcedTeam) {
    window.location.href = 'dashboard.html?platform_admin=1';
  }
}

async function setSelectedPlatformAdminTeamPlan(nextPlanType) {
  if (!window.isPlatformAdminUser) return;
  const teamId = String(selectedPlatformAdminTeamId || els.platformAdminSetFreePlanBtn?.dataset?.teamId || els.platformAdminSetPaidPlanBtn?.dataset?.teamId || '').trim();
  if (!teamId) {
    setPlatformAdminPlanStatus('チームを選択してください。', true);
    return;
  }

  const normalizedPlanType = String(nextPlanType || '').trim() === 'paid' ? 'paid' : 'free';
  const teamLabel = getPlatformAdminTeamLabel(teamId, 'このチーム');
  const confirmText = normalizedPlanType === 'paid'
    ? `${teamLabel} を Paid に切り替えます。実行しますか？`
    : `${teamLabel} を free に切り替えます。実行しますか？`;
  if (!window.confirm(confirmText)) return;

  const fn = window.updateTeamPlanForAdmin;
  if (typeof fn !== 'function') {
    setPlatformAdminPlanStatus('プラン更新機能を読み込めませんでした。', true);
    return;
  }

  const freeBtn = els.platformAdminSetFreePlanBtn;
  const paidBtn = els.platformAdminSetPaidPlanBtn;
  const prevFreeText = freeBtn?.textContent || 'free にする';
  const prevPaidText = paidBtn?.textContent || 'Paid にする';
  if (freeBtn) {
    freeBtn.disabled = true;
    freeBtn.textContent = normalizedPlanType === 'free' ? '切替中...' : prevFreeText;
  }
  if (paidBtn) {
    paidBtn.disabled = true;
    paidBtn.textContent = normalizedPlanType === 'paid' ? '切替中...' : prevPaidText;
  }
  setPlatformAdminPlanStatus(`${teamLabel} を${getPlanTypeLabel(normalizedPlanType)}へ切り替えています...`, false);

  try {
    const { data, error } = await fn(teamId, normalizedPlanType);
    if (error) {
      setPlatformAdminPlanStatus(error.message || 'プラン切替に失敗しました。', true);
      return;
    }

    platformAdminTeamsCache = (Array.isArray(platformAdminTeamsCache) ? platformAdminTeamsCache : []).map(row => (
      String(row?.id || '') === teamId
        ? { ...row, ...(data || {}), id: teamId, team_name: row?.team_name || data?.team_name || row?.name || '-' }
        : row
    ));
    renderPlatformAdminTeams();
    await renderPlatformAdminDetail(teamId);

    const currentWorkspaceTeamId = String(window.currentWorkspaceTeamId || getCurrentWorkspaceTeamIdSync() || window.currentWorkspaceInfo?.id || '').trim();
    if (currentWorkspaceTeamId && currentWorkspaceTeamId === teamId) {
      await loadCurrentTeamPlan();
    }

    setPlatformAdminPlanStatus(`${teamLabel} を${getPlanTypeLabel(normalizedPlanType)}へ切り替えました。`, false);
  } finally {
    if (freeBtn) freeBtn.textContent = prevFreeText;
    if (paidBtn) paidBtn.textContent = prevPaidText;
  }
}

async function saveSelectedPlatformAdminTeamName() {
  if (!window.isPlatformAdminUser) return;
  const teamId = String(els.platformAdminTeamNameInput?.dataset?.teamId || selectedPlatformAdminTeamId || '').trim();
  if (!teamId) {
    setPlatformAdminTeamNameStatus('チームを選択してください。', true);
    return;
  }
  const nextName = String(els.platformAdminTeamNameInput?.value || '').trim().replace(/\s+/g, ' ');
  if (!nextName) {
    setPlatformAdminTeamNameStatus('チーム名を入力してください。', true);
    els.platformAdminTeamNameInput?.focus();
    return;
  }
  if (nextName.length > 80) {
    setPlatformAdminTeamNameStatus('チーム名は80文字以内で入力してください。', true);
    els.platformAdminTeamNameInput?.focus();
    return;
  }

  const currentLabel = getPlatformAdminTeamLabel(teamId, 'このチーム');
  if (nextName === currentLabel) {
    setPlatformAdminTeamNameStatus('変更はありません。', false);
    return;
  }

  const fn = window.updateTeamNameForAdmin;
  if (typeof fn !== 'function') {
    setPlatformAdminTeamNameStatus('チーム名更新機能を読み込めませんでした。', true);
    return;
  }

  const saveBtn = els.savePlatformAdminTeamNameBtn;
  const previousText = saveBtn?.textContent || 'チーム名を保存';
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = '保存中...';
  }
  if (els.platformAdminTeamNameInput) els.platformAdminTeamNameInput.disabled = true;
  setPlatformAdminTeamNameStatus('チーム名を保存しています...', false);

  try {
    const { data, error } = await fn(teamId, nextName);
    if (error) {
      setPlatformAdminTeamNameStatus(error.message || 'チーム名の保存に失敗しました。', true);
      return;
    }
    const resolvedName = String(data?.team_name || data?.name || nextName).trim() || nextName;
    updatePlatformAdminTeamNameCaches(teamId, resolvedName);
    renderPlatformAdminTeams();
    await renderPlatformAdminDetail(teamId);
    setPlatformAdminTeamNameStatus(`チーム名を「${resolvedName}」に更新しました。`, false);
  } finally {
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.textContent = previousText;
    }
    if (els.platformAdminTeamNameInput) els.platformAdminTeamNameInput.disabled = false;
  }
}

function bindPlatformAdminTableEvents() {
  if (!els.platformAdminTeamsTableBody || els.platformAdminTeamsTableBody.dataset.bound === '1') return;
  els.platformAdminTeamsTableBody.dataset.bound = '1';
  els.platformAdminTeamsTableBody.addEventListener('click', async event => {
    const row = event.target.closest('tr[data-admin-team-id]');
    const teamId = row?.dataset?.adminTeamId || '';
    if (!teamId) return;
    if (event.target.closest('.admin-team-detail-btn')) {
      await renderPlatformAdminDetail(teamId);
      return;
    }
    if (event.target.closest('.admin-team-switch-btn')) {
      const rowData = platformAdminTeamsCache.find(item => String(item.id || '') === String(teamId || '')) || {};
      forceSwitchToTeam(teamId, rowData.team_name || rowData.name || '');
      return;
    }
    if (event.target.closest('.admin-team-suspend-btn')) {
      await changePlatformAdminTeamStatus(teamId, 'suspended');
      return;
    }
    if (event.target.closest('.admin-team-resume-btn')) {
      await changePlatformAdminTeamStatus(teamId, 'active');
    }
  });
}

async function refreshAdminHeaderVisibility() {
  const btn = els.adminHeaderBtn;
  if (!btn) return;
  const isOwner = isOwnerUser();
  btn.classList.toggle("hidden", !isOwner);
  btn.textContent = "ユーザー管理";
}

function goToAdminPage() {
  activateTab('settingsTab');
  const target = els.userManagementSection;
  ensureSettingsSectionExpanded(target, true);
  if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function setupEvents() {
  els.logoutBtn?.addEventListener("click", logout);
  els.adminHeaderBtn?.addEventListener("click", goToAdminPage);
  els.refreshProfilesBtn?.addEventListener("click", loadProfilesForSettings);
  els.startPaidCheckoutBtn?.addEventListener('click', startPaidCheckoutFromSettings);
  els.openBillingPortalBtn?.addEventListener('click', openBillingPortalFromSettings);
  els.refreshPlanInfoBtn?.addEventListener('click', refreshPlanInfoFromSettings);
  els.refreshInvitationsBtn?.addEventListener("click", loadInvitationSettingsData);
  els.sendInvitationBtn?.addEventListener("click", submitInvitationFromSettings);
  els.inviteRoleSelect?.addEventListener("change", renderInvitationRoleOptions);
  bindProfileTableEvents();
  bindInvitationTableEvents();
  bindPlatformAdminTableEvents();
  els.refreshPlatformAdminTeamsBtn?.addEventListener('click', loadPlatformAdminTeams);
  els.refreshPlatformAdminAnalyticsBtn?.addEventListener('click', loadPlatformAdminAccessAnalytics);
  els.platformAdminBackBtn?.addEventListener('click', clearForceTeamMode);
  els.switchPlatformTeamBtn?.addEventListener('click', () => forceSwitchToTeam(els.switchPlatformTeamBtn?.dataset?.teamId || '', els.switchPlatformTeamBtn?.dataset?.teamName || ''));
  els.suspendPlatformTeamBtn?.addEventListener('click', async () => {
    if (selectedPlatformAdminTeamId) await changePlatformAdminTeamStatus(selectedPlatformAdminTeamId, 'suspended');
  });
  els.resumePlatformTeamBtn?.addEventListener('click', async () => {
    if (selectedPlatformAdminTeamId) await changePlatformAdminTeamStatus(selectedPlatformAdminTeamId, 'active');
  });
  els.exportPlatformTeamBackupBtn?.addEventListener('click', exportSelectedPlatformAdminTeamBackup);
  els.importPlatformTeamBackupBtn?.addEventListener('click', triggerImportPlatformTeamBackup);
  els.importPlatformTeamBackupInput?.addEventListener('change', importSelectedPlatformAdminTeamBackup);
  els.deletePlatformTeamBtn?.addEventListener('click', deleteSelectedPlatformAdminTeam);
  els.savePlatformAdminTeamNameBtn?.addEventListener('click', saveSelectedPlatformAdminTeamName);
  els.platformAdminSetFreePlanBtn?.addEventListener('click', () => setSelectedPlatformAdminTeamPlan('free'));
  els.platformAdminSetPaidPlanBtn?.addEventListener('click', () => setSelectedPlatformAdminTeamPlan('paid'));
  els.platformAdminTeamNameInput?.addEventListener('keydown', event => {
    if (event.key === 'Enter') {
      event.preventDefault();
      saveSelectedPlatformAdminTeamName();
    }
  });
  els.exportAllBtn?.addEventListener("click", exportAllData);
  els.importAllBtn?.addEventListener("click", triggerImportAll);
  els.importAllFileInput?.addEventListener("change", importAllDataFromFile);
  els.openManualBtn?.addEventListener("click", openManual);
  els.dangerResetBtn?.addEventListener("click", resetAllDataDanger);
  els.resetCastsBtn?.addEventListener("click", resetAllCastsDanger);
  els.resetVehiclesBtn?.addEventListener("click", resetAllVehiclesDanger);
  els.fetchOriginLatLngBtn?.addEventListener("click", fetchOriginLatLngFromAddress);
  els.openOriginGoogleMapBtn?.addEventListener("click", openOriginGoogleMapFromSettings);
  els.saveOriginBtn?.addEventListener("click", saveOriginFromSettings);
  els.useOriginDraftBtn?.addEventListener("click", useOriginDraftFromSettings);
  els.cancelOriginEditBtn?.addEventListener("click", resetOriginForm);
  els.originSlotSelect?.addEventListener("change", () => {
    const slotNo = normalizeOriginSlotNo(els.originSlotSelect?.value || "");
    const row = getOriginRowBySlot(slotNo);
    if (row) fillOriginForm(row, { slotNo });
    else fillOriginForm({}, { slotNo });
  });

  ensureCastTravelMinutesUi();
  els.saveCastBtn?.addEventListener("click", saveCast);
  els.guessAreaBtn?.addEventListener("click", guessCastArea);
  els.fetchCastTravelMinutesBtn?.addEventListener("click", async () => {
    const access = canUseGoogleApi({ purpose: 'coordinate_lookup' });
    if (!access.allowed) {
      alert(access.reason || 'このプランではAPI座標取得を利用できません。');
      return;
    }
    await triggerCastAddressGeocodeNow();
  });
  els.castAddress?.addEventListener("input", () => {
    const nextAddress = String(els.castAddress?.value || "").trim();
    const nextKey = normalizeGeocodeAddressKey(nextAddress || "");
    if (!nextAddress) {
      if (els.castLat) els.castLat.value = "";
      if (els.castLng) els.castLng.value = "";
      if (els.castLatLngText) els.castLatLngText.value = "";
      if (els.castDistanceKm) els.castDistanceKm.value = "";
      if (els.castTravelMinutes) els.castTravelMinutes.value = "";
      lastCastGeocodeKey = "";
      syncCastBlankMetricsUi();
      return;
    }
    if (nextKey && nextKey !== lastCastGeocodeKey) {
      if (els.castLat) els.castLat.value = "";
      if (els.castLng) els.castLng.value = "";
      if (els.castLatLngText) els.castLatLngText.value = "";
      if (els.castDistanceKm) els.castDistanceKm.value = "";
      if (els.castTravelMinutes) els.castTravelMinutes.value = "";
      updateCastDistanceHint();
    }
    scheduleCastAutoGeocode();
  });
  els.castAddress?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    triggerCastAddressGeocodeNow();
  });
  els.castLatLngText?.addEventListener("change", () => {
    const hasText = String(els.castLatLngText?.value || "").trim();
    if (!hasText) {
      if (!String(els.castAddress?.value || '').trim()) syncCastBlankMetricsUi();
      return;
    }
    applyCastLatLng();
  });
  els.openGoogleMapBtn?.addEventListener("click", () => openGoogleMap(els.castAddress?.value || "", els.castLat?.value, els.castLng?.value));
  els.cancelEditBtn?.addEventListener("click", resetCastForm);
  els.importCsvBtn?.addEventListener("click", () => els.csvFileInput?.click());
  els.exportCsvBtn?.addEventListener("click", exportCastsCsv);
  els.csvFileInput?.addEventListener("change", importCastCsvFile);
  els.castSearchRunBtn?.addEventListener("click", renderCastSearchResults);
  els.castSearchResetBtn?.addEventListener("click", resetCastSearchFilters);
  els.castSearchName?.addEventListener("input", renderCastSearchResults);
  els.castSearchArea?.addEventListener("input", renderCastSearchResults);
  els.castSearchAddress?.addEventListener("input", renderCastSearchResults);
  els.castSearchPhone?.addEventListener("input", renderCastSearchResults);

  els.saveVehicleBtn?.addEventListener("click", saveVehicle);
  els.vehicleHomeLatLngText?.addEventListener("change", () => {
    const hasText = String(els.vehicleHomeLatLngText?.value || "").trim();
    if (!hasText) {
      if (els.vehicleHomeLat) els.vehicleHomeLat.value = "";
      if (els.vehicleHomeLng) els.vehicleHomeLng.value = "";
      setVehicleGeoStatus("idle", "座標を「緯度, 経度」の形式で貼り付けると自動反映します");
      return;
    }
    applyVehicleLatLng();
  });
  els.cancelVehicleEditBtn?.addEventListener("click", resetVehicleForm);
  els.importVehicleCsvBtn?.addEventListener("click", () => els.vehicleCsvFileInput?.click());
  els.exportVehicleCsvBtn?.addEventListener("click", exportVehiclesCsv);
  els.vehicleCsvFileInput?.addEventListener("change", importVehicleCsvFile);
  els.exportPlansCsvBtn?.addEventListener("click", exportPlansCsv);
  els.importPlansCsvBtn?.addEventListener("click", triggerImportPlansCsv);
  els.plansCsvFileInput?.addEventListener("change", importPlansCsvFile);
  els.previewMileageReportBtn?.addEventListener("click", async () => {
    await previewDriverMileageReport();
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });
  els.exportMileageReportBtn?.addEventListener("click", exportDriverMileageReportXlsx);
  els.mileageReportStartDate?.addEventListener("change", async () => {
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });
  els.mileageReportEndDate?.addEventListener("change", async () => {
    await refreshHomeMonthlyVehicleList();
    renderVehiclesTable();
  });

  els.savePlanBtn?.addEventListener("click", savePlan);
  els.guessPlanAreaBtn?.addEventListener("click", guessPlanArea);
  els.clearPlansBtn?.addEventListener("click", clearAllPlans);

  els.saveActualBtn?.addEventListener("click", saveActual);
  els.guessActualAreaBtn?.addEventListener("click", guessActualArea);

  bindPlanAndActualFormEvents();
  setupSearchableCastInputs();
  bindDispatchEvents();
  bindPostDispatchEvents();

  els.checkAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(true));
  els.uncheckAllVehiclesBtn?.addEventListener("click", () => toggleAllVehicles(false));
  els.clearManualLastVehicleBtn?.addEventListener("click", clearManualLastVehicle);
  els.resetMonthlySummaryBtn?.addEventListener("click", resetMonthlySummary);

  els.dispatchDate?.addEventListener("change", syncDateAndReloadFromDispatchDate);
  els.planDate?.addEventListener("change", syncDateAndReloadFromPlanDate);
  els.actualDate?.addEventListener("change", syncDateAndReloadFromActualDate);

  els.sendLineBtn?.addEventListener("click", sendDispatchResultToLine);
  els.saveDailyMileageBtn?.addEventListener("click", saveDailyMileageReports);

  els.copyActualTableBtn?.addEventListener("click", copyActualTableFormatted);
}

function formatActualCopyDate(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [y, m, d] = raw.split("-");
    return `${y}/${m}/${d}`;
  }
  return raw;
}

function formatActualCopyDistance(distanceValue) {
  if (distanceValue === null || distanceValue === undefined || distanceValue === "") return "-";
  const num = Number(distanceValue);
  if (Number.isFinite(num)) {
    const body = Number.isInteger(num) ? num.toFixed(0) : num.toFixed(1).replace(/\.0$/, "");
    return `${body}km`;
  }
  const text = String(distanceValue).trim();
  if (!text) return "-";
  return /km$/i.test(text) ? text : `${text}km`;
}

function buildActualTableCopyText() {
  const items = Array.isArray(currentActualsCache) ? currentActualsCache.filter(Boolean) : [];
  if (!items.length) return "";

  const hourLabelFn = typeof getHourLabel === "function"
    ? getHourLabel
    : (hour => `${Number(hour || 0)}時`);
  const areaLabelFn = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel
    : (value => String(value || "無し"));
  const statusLabelFn = typeof getStatusText === "function"
    ? getStatusText
    : (value => {
        const raw = String(value || "pending").toLowerCase();
        if (raw === "done") return "完了";
        if (raw === "cancel") return "キャンセル";
        return "未完了";
      });

  const dateLabel = formatActualCopyDate(els.actualDate?.value || els.dispatchDate?.value || "");
  const lines = [`【実際の送り${dateLabel ? ` ${dateLabel}` : ""}】`, ""];

  const sortedItems = [...items].sort((a, b) => {
    const hourDiff = Number(a?.actual_hour ?? 0) - Number(b?.actual_hour ?? 0);
    if (hourDiff !== 0) return hourDiff;
    const areaDiff = String(areaLabelFn(a?.destination_area || "無し")).localeCompare(String(areaLabelFn(b?.destination_area || "無し")), 'ja');
    if (areaDiff !== 0) return areaDiff;
    const nameA = String(a?.person_name || a?.casts?.name || "");
    const nameB = String(b?.person_name || b?.casts?.name || "");
    return nameA.localeCompare(nameB, 'ja');
  });

  const hours = [...new Set(sortedItems.map(item => Number(item?.actual_hour ?? 0)))].sort((a, b) => a - b);

  hours.forEach((hour, hourIndex) => {
    if (hourIndex > 0) lines.push("");
    lines.push(`■ ${hourLabelFn(hour)}`);

    const hourItems = sortedItems.filter(item => Number(item?.actual_hour ?? 0) === hour);
    const areaKeys = [...new Set(hourItems.map(item => String(areaLabelFn(item?.destination_area || "無し"))))];

    areaKeys.forEach(area => {
      lines.push(`【${area || "無し"}】`);
      hourItems
        .filter(item => String(areaLabelFn(item?.destination_area || "無し")) === area)
        .forEach(item => {
          const name = String(item?.person_name || item?.casts?.name || "-").trim() || "-";
          const distance = formatActualCopyDistance(item?.distance_km);
          const status = String(statusLabelFn(item?.status) || "未完了").trim() || "未完了";
          const note = String(item?.memo || item?.note || "").trim();
          const body = [`・${name}`, distance !== "-" ? distance : null, status, note || null]
            .filter(Boolean)
            .join(" / ");
          lines.push(body);
        });
      lines.push("");
    });

    while (lines.length && lines[lines.length - 1] === "") lines.pop();
  });

  return lines.join("\n").trim();
}

async function copyTextWithFallback(text) {
  const value = String(text || "");
  if (!value.trim()) return false;

  try {
    if (navigator?.clipboard?.writeText && window.isSecureContext) {
      await navigator.clipboard.writeText(value);
      return true;
    }
  } catch (error) {
    console.warn("clipboard api failed", error);
  }

  try {
    const ta = document.createElement("textarea");
    ta.value = value;
    ta.setAttribute("readonly", "readonly");
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    ta.style.left = "-1000px";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    ta.setSelectionRange(0, ta.value.length);
    const copied = document.execCommand("copy");
    document.body.removeChild(ta);
    return !!copied;
  } catch (error) {
    console.warn("textarea fallback failed", error);
    return false;
  }
}

async function copyActualTableFormatted() {
  const text = buildActualTableCopyText();
  if (!text) {
    alert("コピーする実際の送りがありません");
    return;
  }

  const copied = await copyTextWithFallback(text);
  if (copied) {
    alert("表をコピーしました");
    return;
  }

  console.error("actual table copy failed");
  alert("コピーに失敗しました");
}

document.addEventListener("DOMContentLoaded", async () => {
  if (dashboardInitialized) {
    return;
  }
  dashboardInitialized = true;

  try {
    console.log("SUPABASE_URL:", SUPABASE_URL);

    const ok = await ensureAuth();
    if (!ok) return;
    if (window.currentWorkspaceSuspended && !window.isPlatformAdminUser) {
      renderSuspendedWorkspaceMode();
      return;
    }

    await stabilizePlatformAdminForcedContext();
    await loadCurrentTeamPlan();
    applyRoleUi();
    await refreshAdminHeaderVisibility();
    renderAdminForceModeBanner();
    await loadProfilesForSettings();
    await loadInvitationSettingsData();
    if (window.isPlatformAdminUser) await loadPlatformAdminTeams();
    ensureCastTravelMinutesUi();
    initializeMileageReportDefaultDates();
    setupTabs();
    setupEvents();

    resetCastForm();
    syncCastBlankMetricsUi();
    resetVehicleForm();
    resetPlanForm();
    resetActualForm();

    await initializeOriginManagement();

    const today = todayStr();
    if (els.dispatchDate) els.dispatchDate.value = today;
    if (els.planDate) els.planDate.value = today;
    if (els.actualDate) els.actualDate.value = today;
    forceResetMileageReportInputs(today);

    syncScheduleRendererDeps();
    await loadHomeAndAll();
    try {
      if (typeof window.recordDashboardAccess === 'function') {
        await window.recordDashboardAccess({
          userId: currentUser?.id || window.currentUser?.id || null,
          teamId: await resolveCurrentWorkspaceTeamId()
        });
      }
    } catch (_) {}
    applyBillingReturnStateMessage();
    try {
      const query = new URLSearchParams(String(window.location.search || ''));
      if (window.isPlatformAdminUser && !getAdminForcedTeamId() && query.get('platform_admin') === '1') {
        activateTab('platformAdminTab');
      }
    } catch (e) {}
    syncMileageReportRange(today, true);
    window.requestAnimationFrame(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true));
    window.setTimeout(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true), 0);
    window.setTimeout(() => syncMileageReportRange(els.dispatchDate?.value || todayStr(), true), 120);
    bindMileageReportSyncListeners();
    renderManualLastVehicleInfo();
  } catch (err) {
    console.error("dashboard init error:", err);
    alert("初期化中にエラーが発生しました。Console を確認してください。");
  }
});


/* ===== THEMIS v3.7 配車AI強化版 patch start ===== */
const THEMIS_V37_LEARN_KEY = "themis_v37_dispatch_learning_v1";

function getThemisV37LearningStore() {
  try {
    return JSON.parse(window.localStorage.getItem(THEMIS_V37_LEARN_KEY) || "{}") || {};
  } catch (e) {
    console.error(e);
    return {};
  }
}

function saveThemisV37LearningStore(store) {
  try {
    window.localStorage.setItem(THEMIS_V37_LEARN_KEY, JSON.stringify(store || {}));
  } catch (e) {
    console.error(e);
  }
}

function normalizeMunicipalityLabel(value) {
  return String(value || "").trim().replace(/[　\s]+/g, "");
}

function extractMunicipalityFromAddress(address) {
  const normalized = normalizeAddressText(address || "");
  if (!normalized) return "";

  const patterns = [
    /(東京都[^0-9\-]{1,12}?区)/,
    /(東京都[^0-9\-]{1,16}?市)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?市)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?郡[^0-9\-]{1,16}町)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?郡[^0-9\-]{1,16}村)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?町)/,
    /((?:北海道|大阪府|京都府|[^都道府県]{2,6}県)[^0-9\-]{1,16}?村)/
  ];

  for (const pattern of patterns) {
    const matched = normalized.match(pattern);
    if (matched && matched[1]) return normalizeMunicipalityLabel(matched[1]);
  }
  return "";
}

const THEMIS_V37_MUNICIPALITY_AREA_HINTS = [
  { area: "葛飾方面", keys: ["東京都葛飾区"] },
  { area: "足立方面", keys: ["東京都足立区"] },
  { area: "江戸川方面", keys: ["東京都江戸川区"] },
  { area: "墨田方面", keys: ["東京都墨田区"] },
  { area: "江東方面", keys: ["東京都江東区"] },
  { area: "荒川方面", keys: ["東京都荒川区"] },
  { area: "台東方面", keys: ["東京都台東区"] },
  { area: "市川方面", keys: ["千葉県市川市"] },
  { area: "船橋方面", keys: ["千葉県船橋市", "千葉県習志野市"] },
  { area: "鎌ヶ谷方面", keys: ["千葉県鎌ケ谷市", "千葉県鎌ヶ谷市"] },
  { area: "我孫子方面", keys: ["千葉県我孫子市"] },
  { area: "柏方面", keys: ["千葉県柏市"] },
  { area: "流山方面", keys: ["千葉県流山市"] },
  { area: "野田方面", keys: ["千葉県野田市"] },
  { area: "松戸近郊", keys: ["千葉県松戸市"] },
  { area: "三郷方面", keys: ["埼玉県三郷市"] },
  { area: "吉川方面", keys: ["埼玉県吉川市"] },
  { area: "八潮方面", keys: ["埼玉県八潮市"] },
  { area: "草加方面", keys: ["埼玉県草加市"] },
  { area: "越谷方面", keys: ["埼玉県越谷市"] },
  { area: "取手方面", keys: ["茨城県取手市"] },
  { area: "藤代方面", keys: ["茨城県取手市藤代"] },
  { area: "守谷方面", keys: ["茨城県守谷市"] },
  { area: "つくば方面", keys: ["茨城県つくば市"] },
  { area: "牛久方面", keys: ["茨城県牛久市"] }
];const _THEMIS_V36_guessArea = guessArea;
// v3.7 municipality extraction is intentionally disabled.
// Keep the cleaner pre-v3.7 area labeling/display while preserving other v3.7 logic.
guessArea = function(lat, lng, address = "") {
  return _THEMIS_V36_guessArea(lat, lng, address);
};

function getThemisV37LearnedAreaScore(homeArea, destArea) {
  const store = getThemisV37LearningStore();
  const key = `${getCanonicalArea(homeArea) || normalizeAreaLabel(homeArea)}__${getCanonicalArea(destArea) || normalizeAreaLabel(destArea)}`;
  return Number(store.areaPair?.[key] || 0);
}

function learnThemisV37FromDoneRows(rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const store = getThemisV37LearningStore();
  store.areaPair = store.areaPair || {};
  store.routePair = store.routePair || {};

  const byVehicle = new Map();
  rows.forEach(row => {
    const vehicleId = Number(row.vehicle_id || 0);
    if (!vehicleId) return;
    if (!byVehicle.has(vehicleId)) byVehicle.set(vehicleId, []);
    byVehicle.get(vehicleId).push(row);

    const homeArea = normalizeAreaLabel(
      allVehiclesCache.find(v => Number(v.id) === vehicleId)?.home_area || row.driver_home_area || ""
    );
    const destArea = normalizeAreaLabel(row.destination_area || row.cluster_area || "無し");
    if (homeArea && destArea && homeArea !== "無し" && destArea !== "無し") {
      const key = `${getCanonicalArea(homeArea) || homeArea}__${getCanonicalArea(destArea) || destArea}`;
      store.areaPair[key] = Math.min(120, Number(store.areaPair[key] || 0) + 3);
    }
  });

  for (const rowsByVehicle of byVehicle.values()) {
    const ordered = [...rowsByVehicle].sort((a, b) => {
      const ah = Number(a.actual_hour ?? 0);
      const bh = Number(b.actual_hour ?? 0);
      if (ah !== bh) return ah - bh;
      return Number(a.stop_order || 0) - Number(b.stop_order || 0);
    });
    for (let i = 0; i < ordered.length - 1; i += 1) {
      const a = normalizeAreaLabel(ordered[i].destination_area || "");
      const b = normalizeAreaLabel(ordered[i + 1].destination_area || "");
      if (!a || !b || a === "無し" || b === "無し") continue;
      const key = `${getCanonicalArea(a) || a}__${getCanonicalArea(b) || b}`;
      store.routePair[key] = Math.min(80, Number(store.routePair[key] || 0) + 2);
    }
  }

  saveThemisV37LearningStore(store);
}

const _THEMIS_V36_confirmDailyToMonthly = confirmDailyToMonthly;
confirmDailyToMonthly = async function() {
  const doneRowsBefore = Array.isArray(currentActualsCache)
    ? currentActualsCache.filter(x => normalizeStatus(x.status) === "done")
    : [];
  const result = await _THEMIS_V36_confirmDailyToMonthly.apply(this, arguments);
  try {
    learnThemisV37FromDoneRows(doneRowsBefore);
  } catch (e) {
    console.error(e);
  }
  return result;
};

const _THEMIS_V36_getLastTripHomePriorityWeight = getLastTripHomePriorityWeight;
getLastTripHomePriorityWeight = function(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster) {
  let weight = _THEMIS_V36_getLastTripHomePriorityWeight(clusterArea, homeArea, isLastRun, isDefaultLastHourCluster);
  const learned = getThemisV37LearnedAreaScore(homeArea, clusterArea);
  const strict = getStrictHomeCompatibilityScore(clusterArea, homeArea);
  const direction = getDirectionAffinityScore(clusterArea, homeArea);

  let returnTimeScore = 0;
  if (strict >= 78) returnTimeScore += 42;
  else if (strict >= 52) returnTimeScore += 22;
  if (direction >= 72) returnTimeScore += 18;
  else if (direction >= 28) returnTimeScore += 8;
  if (isHardReverseForHome(clusterArea, homeArea)) returnTimeScore -= (isLastRun ? 120 : 50);

  weight += learned * (isLastRun ? 1.6 : 0.8);
  weight += returnTimeScore * (isLastRun ? 1.8 : (isDefaultLastHourCluster ? 1.2 : 0.35));
  return weight;
};

function getThemisV37LearnedRoutePairScore(areaA, areaB) {
  const store = getThemisV37LearningStore();
  const key1 = `${getCanonicalArea(areaA) || normalizeAreaLabel(areaA)}__${getCanonicalArea(areaB) || normalizeAreaLabel(areaB)}`;
  const key2 = `${getCanonicalArea(areaB) || normalizeAreaLabel(areaB)}__${getCanonicalArea(areaA) || normalizeAreaLabel(areaA)}`;
  return Math.max(Number(store.routePair?.[key1] || 0), Number(store.routePair?.[key2] || 0));
}

function getThemisV37RouteSequenceScore(fromItem, toItem) {
  const pointA = getItemLatLng(fromItem);
  const pointB = getItemLatLng(toItem);
  const areaA = normalizeAreaLabel(fromItem?.destination_area || fromItem?.cluster_area || fromItem?.planned_area || "無し");
  const areaB = normalizeAreaLabel(toItem?.destination_area || toItem?.cluster_area || toItem?.planned_area || "無し");
  const routeFlow = getRouteFlowCompatibilityBetweenAreas(areaA, areaB);
  const continuityPenalty = getPairRouteContinuityPenalty(areaA, areaB);
  const learned = getThemisV37LearnedRoutePairScore(areaA, areaB);
  let score = routeFlow * 2.4 + learned * 4.2 - continuityPenalty * 1.35;

  if (pointA && pointB) {
    const leg = estimateRoadKmBetweenPoints(pointA.lat, pointA.lng, pointB.lat, pointB.lng);
    score -= leg * 3.2;
  } else {
    score -= Math.abs(Number(toItem?.distance_km || 0) - Number(fromItem?.distance_km || 0)) * 0.4;
  }

  const dirA = getAreaDirectionCluster(areaA);
  const dirB = getAreaDirectionCluster(areaB);
  if (dirA && dirB && dirA === dirB) score += 18;
  return score;
}

sortItemsByNearestRoute = function(items) {
  const remaining = [...items];
  const sorted = [];
  let currentLat = ORIGIN_LAT;
  let currentLng = ORIGIN_LNG;
  let currentItem = null;

  while (remaining.length) {
    let bestIndex = 0;
    let bestScore = -Infinity;

    remaining.forEach((item, index) => {
      const point = getItemLatLng(item);
      let score = 0;
      if (point) {
        score -= estimateRoadKmBetweenPoints(currentLat, currentLng, point.lat, point.lng) * 3.0;
      } else {
        score -= Number(item.distance_km || 999999) * 1.1;
      }

      if (currentItem) {
        score += getThemisV37RouteSequenceScore(currentItem, item);
      } else {
        const area = normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || "無し");
        score += getRouteFlowSortWeight(area) * 4.8;
        score += getAreaAffinityScore(area, "松戸近郊") * 0.12;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    const picked = remaining.splice(bestIndex, 1)[0];
    sorted.push(picked);
    currentItem = picked;

    const pickedPoint = getItemLatLng(picked);
    if (pickedPoint) {
      currentLat = pickedPoint.lat;
      currentLng = pickedPoint.lng;
    }
  }

  return sorted;
};

const _THEMIS_V36_runAutoDispatch = runAutoDispatch;
runAutoDispatch = async function() {
  const result = await _THEMIS_V36_runAutoDispatch.apply(this, arguments);
  try {
    await loadActualsByDate(els.actualDate?.value || todayStr());
    renderDailyDispatchResult();
  } catch (e) {
    console.error(e);
  }
  return result;
};
/* ===== THEMIS v3.7 配車AI強化版 patch end ===== */


/* ===== THEMIS v5.4 配車AI強化版 patch start ===== */

function getThemisV54RowPoint(row) {
  const lat = toNullableNumber(row?.casts?.latitude ?? row?.latitude);
  const lng = toNullableNumber(row?.casts?.longitude ?? row?.longitude);
  if (isValidLatLng(lat, lng)) return { lat, lng };
  return null;
}

function getThemisV54RowArea(row) {
  return normalizeAreaLabel(
    row?.destination_area ||
    row?.planned_area ||
    row?.cluster_area ||
    row?.casts?.area ||
    ""
  );
}

function getThemisV54RowDistance(row) {
  return Number(
    row?.distance_km ??
    row?.casts?.distance_km ??
    0
  ) || 0;
}

function getThemisV54StoredTravelMinutes(row) {
  const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
  return stored > 0 ? stored : 0;
}

function getThemisV54SegmentDistanceKm(fromPoint, toRow) {
  const toPoint = getThemisV54RowPoint(toRow);
  if (fromPoint && toPoint) {
    const rad = (deg) => (deg * Math.PI) / 180;
    const R = 6371;
    const dLat = rad(toPoint.lat - fromPoint.lat);
    const dLng = rad(toPoint.lng - fromPoint.lng);
    const lat1 = rad(fromPoint.lat);
    const lat2 = rad(toPoint.lat);
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Number((R * c).toFixed(1));
  }
  return getThemisV54RowDistance(toRow);
}

function getThemisV54SegmentMinutes(fromPoint, toRow) {
  const segKm = Math.max(0, Number(getThemisV54SegmentDistanceKm(fromPoint, toRow) || 0));
  const area = getThemisV54RowArea(toRow);
  const storedMinutes = getThemisV54StoredTravelMinutes(toRow);
  const baseDistance = Math.max(0.1, Number(getThemisV54RowDistance(toRow) || segKm || 0.1));

  if (storedMinutes > 0) {
    const derivedSpeed = Math.max(16, Math.min(60, (baseDistance / storedMinutes) * 60));
    return Math.max(1, Math.round((segKm / derivedSpeed) * 60));
  }

  return Math.max(1, Math.round(estimateFallbackTravelMinutes(segKm, area)));
}

function getThemisV54VehicleIdFromRows(rows) {
  const first = Array.isArray(rows) ? rows.find(Boolean) : null;
  return Number(first?.vehicle_id || 0);
}

function getThemisV54VehicleIsLastTrip(rows) {
  const vehicleId = getThemisV54VehicleIdFromRows(rows);
  return vehicleId > 0 ? isDriverLastTripChecked(vehicleId) : false;
}

function getThemisV54TravelSummary(rows) {
  const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!ordered.length) {
    return {
      outboundKm: 0,
      returnKm: 0,
      totalKm: 0,
      outboundMinutes: 0,
      returnMinutes: 0,
      totalMinutes: 0,
      sendOnlyMinutes: 0,
      stopCount: 0,
      segmentKm: [],
      segmentMinutes: [],
      model: "legacy_segment_speed"
    };
  }

  const stopCount = ordered.length;
  const segmentKm = [];
  const segmentMinutes = [];
  let currentPoint = { lat: ORIGIN_LAT, lng: ORIGIN_LNG };
  let outboundKm = 0;
  let outboundMinutes = 0;

  ordered.forEach((row, index) => {
    let legKm = 0;
    let legMinutes = 0;
    if (index === 0) {
      legKm = Math.max(0, Number(getThemisV54RowDistance(row) || 0));
      legMinutes = getThemisV54StoredTravelMinutes(row);
      if (!(legMinutes > 0)) {
        legMinutes = Math.max(1, Math.round(
          estimateFallbackTravelMinutes(legKm, getThemisV54RowArea(row))
        ));
      }
    } else {
      legKm = getThemisV54SegmentDistanceKm(currentPoint, row);
      legMinutes = getThemisV54SegmentMinutes(currentPoint, row);
    }

    legKm = Math.max(0, Number(legKm || 0));
    legMinutes = Math.max(1, Math.round(Number(legMinutes || 0)));
    segmentKm.push(Number(legKm.toFixed(1)));
    segmentMinutes.push(legMinutes);
    outboundKm += legKm;
    outboundMinutes += legMinutes;

    const point = getThemisV54RowPoint(row);
    if (point) currentPoint = point;
  });

  const sendOnlyMinutes = Math.max(0, Math.round(outboundMinutes));
  const isLastTrip = getThemisV54VehicleIsLastTrip(ordered);
  const lastRow = ordered[ordered.length - 1] || null;
  let returnKm = 0;
  let returnMinutes = 0;
  if (!isLastTrip && lastRow) {
    returnKm = Math.max(0, Number(getThemisV54RowDistance(lastRow) || 0));
    returnMinutes = Math.max(0, Math.round(getThemisV54StoredTravelMinutes(lastRow) || 0));
    if (!(returnMinutes > 0)) {
      returnMinutes = Math.max(0, Math.round(
        estimateFallbackTravelMinutes(returnKm, getThemisV54RowArea(lastRow))
      ));
    }
  }
  const totalKm = Math.max(0, Number((outboundKm + returnKm).toFixed(1)));
  const totalMinutes = Math.max(0, Math.round(sendOnlyMinutes + returnMinutes));

  return {
    outboundKm: Number(outboundKm.toFixed(1)),
    returnKm: Number(returnKm.toFixed(1)),
    totalKm,
    outboundMinutes: Math.round(outboundMinutes),
    returnMinutes,
    totalMinutes,
    sendOnlyMinutes,
    stopCount,
    segmentKm,
    segmentMinutes,
    model: "legacy_segment_speed"
  };
}

getRowsOutboundMinutes = function(rows) {
  return getThemisV54TravelSummary(rows).outboundMinutes;
};

getRowsReturnMinutes = function(rows) {
  return getThemisV54TravelSummary(rows).returnMinutes;
};

getRowsTravelTimeSummary = function(rows) {
  return getThemisV54TravelSummary(rows);
};

calculateRouteDistanceGlobal = function(rows) {
  return Number(getThemisV54TravelSummary(rows).outboundKm || 0);
};

calcVehicleRotationForecastGlobal = function(vehicle, orderedRows) {
  const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
  if (!rows.length) {
    return {
      routeDistanceKm: 0,
      returnDistanceKm: 0,
      zoneLabel: "-",
      predictedDepartureTime: "-",
      predictedReturnTime: "-",
      predictedReadyTime: "-",
      predictedReturnMinutes: 0,
      extraSharedDelayMinutes: 0,
      stopCount: 0,
      returnAfterLabel: "-"
    };
  }

  const firstHour = rows.reduce((min, row) => {
    const val = Number(row.actual_hour ?? row.plan_hour ?? 0);
    return Number.isFinite(val) ? Math.min(min, val) : min;
  }, 99);

  const baseHour = firstHour === 99 ? 0 : firstHour;
  const summary = getThemisV54TravelSummary(rows);
  const routeDistanceKm = Number(summary.outboundKm || 0);
  const returnDistanceKm = Number(summary.returnKm || 0);
  const representativeArea = getRepresentativeAreaFromRows(rows);
  const primaryZone = getDistanceZoneInfoGlobal(Math.max(routeDistanceKm, returnDistanceKm), representativeArea);

  const departDelayMinutes = (typeof getExpectedDepartureDelayMinutes === "function" ? getExpectedDepartureDelayMinutes(baseHour) : 0);
  const predictedDepartureAbs = baseHour * 60 + departDelayMinutes;
  const predictedReturnAbs = predictedDepartureAbs + Number(summary.totalMinutes || 0);
  const predictedReadyAbs = predictedReturnAbs + 1;

  let extraSharedDelayMinutes = 0;
  if (rows.length >= 2) {
    const firstOnlySummary = getThemisV54TravelSummary([rows[0]]);
    extraSharedDelayMinutes = Math.max(0, Number(summary.totalMinutes || 0) - Number(firstOnlySummary.totalMinutes || 0));
  }

  return {
    routeDistanceKm: Number(routeDistanceKm.toFixed(1)),
    returnDistanceKm: Number(returnDistanceKm.toFixed(1)),
    zoneLabel: primaryZone.label,
    predictedDepartureTime: formatClockTimeFromMinutesGlobal(predictedDepartureAbs),
    predictedReturnTime: formatClockTimeFromMinutesGlobal(predictedReturnAbs),
    predictedReadyTime: formatClockTimeFromMinutesGlobal(predictedReadyAbs),
    predictedReturnMinutes: Math.round(Number(summary.totalMinutes || 0)),
    extraSharedDelayMinutes: Math.round(extraSharedDelayMinutes),
    stopCount: rows.length,
    returnAfterLabel: `${Math.round(Number(summary.totalMinutes || 0))}分後`
  };
};

sortItemsByNearestRoute = function(items) {
  const remaining = [...(items || [])].filter(Boolean);
  const sorted = [];
  let currentPoint = { lat: ORIGIN_LAT, lng: ORIGIN_LNG };
  let currentItem = null;

  while (remaining.length) {
    let bestIndex = 0;
    let bestScore = -Infinity;

    remaining.forEach((item, index) => {
      const point = getThemisV54RowPoint(item);
      const area = getThemisV54RowArea(item);
      const storedTravel = getThemisV54StoredTravelMinutes(item);
      const baseDistance = getThemisV54RowDistance(item);

      let score = 0;

      const legKm = point
        ? estimateRoadKmBetweenPoints(currentPoint.lat, currentPoint.lng, point.lat, point.lng)
        : Math.max(0, baseDistance);

      score -= legKm * 3.2;

      if (currentItem) {
        const prevArea = getThemisV54RowArea(currentItem);
        score += getRouteFlowCompatibilityBetweenAreas(prevArea, area) * 2.1;
        score -= getPairRouteContinuityPenalty(prevArea, area) * 1.35;
        if (isHardReverseMixForRoute(prevArea, area)) score -= 460;
        score += getDirectionAffinityScore(prevArea, area) * 0.42;
      } else {
        score += getRouteFlowSortWeight(area) * 3.4;
        score += Math.min(40, storedTravel) * 0.7;
      }

      // 1件目は travel_minutes の長い送り先を先頭にしすぎないよう軽く抑える
      if (!currentItem && storedTravel > 0) {
        score -= Math.max(0, storedTravel - 25) * 0.18;
      }

      // 近いのに大きく逆流する候補は避ける
      if (currentItem) {
        const prevArea = getThemisV54RowArea(currentItem);
        if (getDirectionAffinityScore(prevArea, area) <= -38) {
          score -= 120;
        }
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    const picked = remaining.splice(bestIndex, 1)[0];
    sorted.push(picked);
    currentItem = picked;

    const pickedPoint = getThemisV54RowPoint(picked);
    if (pickedPoint) currentPoint = pickedPoint;
  }

  return sorted;
};

function getThemisV54VehicleProjectedAvg(vehicleId, monthlyMap, extraAssignedDistanceMap, extraAssignedCountMap) {
  const stats = monthlyMap?.get(Number(vehicleId)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
  const reportDate = els.dispatchDate?.value || els.actualDate?.value || todayStr();

  const hasReportToday = Array.isArray(currentDailyReportsCache)
    ? currentDailyReportsCache.some(
        row => Number(row.vehicle_id || 0) === Number(vehicleId) &&
               String(row.report_date || "") === String(reportDate)
      )
    : false;

  const extraDistance = Number(extraAssignedDistanceMap.get(Number(vehicleId)) || 0);
  const extraCount = Number(extraAssignedCountMap.get(Number(vehicleId)) || 0);
  const projectedDays = Math.max(1, Number(stats.workedDays || 0) + (hasReportToday ? 0 : (extraCount > 0 ? 1 : 0)));
  return (Number(stats.totalDistance || 0) + extraDistance) / projectedDays;
}

optimizeAssignmentsByDistanceBalance = function(assignments, items, vehicles, monthlyMap) {
  return Array.isArray(assignments) ? assignments : [];
};

applyLastTripDistanceCorrectionToAssignments = function(assignments, items, vehicles, monthlyMap) {
  const working = (assignments || []).map(a => ({ ...a }));
  if (!working.length || !Array.isArray(vehicles) || !vehicles.length) return working;

  const itemMap = new Map((items || []).map(item => [Number(item.id), item]));
  const dateStr = els.actualDate?.value || todayStr();
  const defaultLastHour = getDefaultLastHour(dateStr);
  const targetHour = working.some(a => Number(a.actual_hour ?? 0) === Number(defaultLastHour))
    ? Number(defaultLastHour)
    : Math.max(...working.map(a => Number(a.actual_hour ?? 0)));

  const targetRows = working.filter(a => Number(a.actual_hour ?? 0) === Number(targetHour));
  if (!targetRows.length) return working;

  const manualVehicleId = getManualLastVehicleId();

  const projectedDistanceByVehicle = new Map();
  const projectedCountByVehicle = new Map();
  working.forEach(a => {
    const item = itemMap.get(Number(a.item_id));
    projectedDistanceByVehicle.set(
      Number(a.vehicle_id),
      Number(projectedDistanceByVehicle.get(Number(a.vehicle_id)) || 0) +
        Number(item?.distance_km ?? a.distance_km ?? 0)
    );
    projectedCountByVehicle.set(
      Number(a.vehicle_id),
      Number(projectedCountByVehicle.get(Number(a.vehicle_id)) || 0) + 1
    );
  });

  const evaluate = (vehicle, item) => {
    const area = getThemisV54RowArea(item);
    const strict = getStrictHomeCompatibilityScore(area, vehicle?.home_area || "");
    const direction = Math.max(0, getDirectionAffinityScore(area, vehicle?.home_area || ""));
    const affinity = getAreaAffinityScore(area, vehicle?.home_area || "");
    const vehicleMatch = getVehicleAreaMatchScore(vehicle, area);
    const hardReverse = isHardReverseForHome(area, vehicle?.home_area || "");
    const projectedAvg = getThemisV54VehicleProjectedAvg(vehicle?.id, monthlyMap, projectedDistanceByVehicle, projectedCountByVehicle);

    let score = strict * 9 + direction * 5 + affinity * 3.8 + vehicleMatch * 0.8;
    score -= projectedAvg * 0.95;
    if (hardReverse) score -= 9999;
    if (Number(vehicle?.id) === Number(manualVehicleId) && strict >= 52 && !hardReverse) score += 140;
    if (isDriverLastTripChecked(vehicle?.id)) score += 120;
    return score;
  };

  for (const target of targetRows) {
    const item = itemMap.get(Number(target.item_id));
    if (!item) continue;

    const currentVehicle = vehicles.find(v => Number(v.id) === Number(target.vehicle_id));
    let best = { vehicle: currentVehicle, score: evaluate(currentVehicle, item) };

    for (const vehicle of vehicles) {
      const seatCapacity = Number(vehicle.seat_capacity || 4);
      const load = working.filter(
        a =>
          Number(a.vehicle_id) === Number(vehicle.id) &&
          Number(a.actual_hour ?? 0) === Number(targetHour) &&
          Number(a.item_id) !== Number(target.item_id)
      ).length;
      if (load >= seatCapacity) continue;

      const candidateScore = evaluate(vehicle, item);
      if (candidateScore > best.score) best = { vehicle, score: candidateScore };
    }

    if (best.vehicle && Number(best.vehicle.id) !== Number(target.vehicle_id) && best.score >= evaluate(currentVehicle, item) + 12) {
      target.vehicle_id = best.vehicle.id;
      target.driver_name = best.vehicle.driver_name || "";
      target.manual_last_vehicle = Number(best.vehicle.id) === Number(manualVehicleId);
    }
  }

  return working;
};

const _THEMIS_V54_BASE_runAutoDispatch = runAutoDispatch;
runAutoDispatch = async function() {
  const result = await _THEMIS_V54_BASE_runAutoDispatch.apply(this, arguments);
  try {
    await loadActualsByDate(els.actualDate?.value || todayStr());
    renderDailyDispatchResult();
  } catch (error) {
    console.error("THEMIS v5.4 rerender error:", error);
  }
  return result;
};

/* ===== THEMIS v5.4 配車AI強化版 patch end ===== */


/* ===== THEMIS v5.5.1 patch start ===== */
(function(){
  const _THEMIS_V55_BASE_optimizeAssignments = optimizeAssignments;

  function v551RoundTripMinutesForRow(row) {
    const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
    if (stored > 0) return stored * 2;
    return Math.max(0, Math.round(estimateFallbackTravelMinutes(Number(row?.distance_km || row?.casts?.distance_km || 0), normalizeAreaLabel(row?.destination_area || row?.casts?.area || '')) * 2));
  }

  function v551AreaSignature(area) {
    const normalized = normalizeAreaLabel(area || '');
    return {
      normalized,
      canonical: getCanonicalArea(normalized) || '',
      group: getAreaDisplayGroup(normalized) || ''
    };
  }

  function v551IsFriendlyDirection(areaA, areaB) {
    const a = v551AreaSignature(areaA);
    const b = v551AreaSignature(areaB);
    if (!a.normalized || !b.normalized) return false;
    if (a.canonical && b.canonical && a.canonical === b.canonical) return true;
    if (a.group && b.group && a.group === b.group) return true;
    const affinity = getAreaAffinityScore(a.normalized, b.normalized);
    const direction = getDirectionAffinityScore(a.normalized, b.normalized);
    return affinity >= 62 || direction >= 26;
  }

  function v551IsStrongReverse(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    return getDirectionAffinityScore(a, b) <= -38;
  }

  function v551VehicleStateSummary(assignments, vehicles, hour) {
    const map = new Map();
    (vehicles || []).forEach(v => map.set(Number(v.id), { count:0, areas:[], rows:[], vehicle:v }));
    (assignments || []).filter(r => Number(r?.actual_hour ?? 0) === Number(hour)).forEach(r => {
      const id = Number(r.vehicle_id || 0);
      if (!map.has(id)) return;
      const state = map.get(id);
      state.count += 1;
      state.rows.push(r);
      const area = normalizeAreaLabel(r.destination_area || r.cluster_area || r.planned_area || r.casts?.area || '');
      if (area) state.areas.push(area);
    });
    return map;
  }

  function v551CanMoveIntoVehicle(row, vehicleState, vehicle) {
    const cap = Number(vehicle?.seat_capacity || 4);
    if (Number(vehicleState?.count || 0) >= cap) return false;
    const rowArea = normalizeAreaLabel(row.destination_area || row.cluster_area || row.planned_area || row.casts?.area || '');
    const areas = Array.isArray(vehicleState?.areas) ? vehicleState.areas : [];
    if (!areas.length) return true;
    if (areas.some(a => v551IsStrongReverse(rowArea, a))) return false;
    return true;
  }

  function v551PickBundleVehicle(bundleRows, vehicleStates, vehicles, monthlyMap) {
    const farthestBundleRow = [...(bundleRows || [])].sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || bundleRows[0];
    const rowArea = normalizeAreaLabel(farthestBundleRow?.destination_area || farthestBundleRow?.cluster_area || farthestBundleRow?.planned_area || farthestBundleRow?.casts?.area || '');
    const rowDistance = Number(farthestBundleRow?.distance_km || 0);
    let best = null;
    for (const [vehicleId, state] of vehicleStates.entries()) {
      const vehicle = state.vehicle || (vehicles || []).find(v => Number(v.id) === Number(vehicleId));
      if (!vehicle) continue;
      const capacity = Number(vehicle.seat_capacity || 4);
      if (state.count + bundleRows.length > capacity) continue;
      if (!bundleRows.every(r => v551CanMoveIntoVehicle(r, state, vehicle))) continue;

      let score = 0;
      const areas = state.areas || [];
      const existingRows = Array.isArray(state.rows) ? state.rows : [];
      const existingAnchorRow = [...existingRows].sort((a, b) => Number(b?.distance_km || 0) - Number(a?.distance_km || 0))[0] || null;
      const existingAnchorArea = normalizeAreaLabel(existingAnchorRow?.destination_area || existingAnchorRow?.cluster_area || existingAnchorRow?.planned_area || existingAnchorRow?.casts?.area || '');
      const existingAnchorDistance = Number(existingAnchorRow?.distance_km || 0);

      if (areas.length) {
        for (const area of areas) {
          const direction = Number(getDirectionAffinityScore(rowArea, area) || 0);
          const affinity = Number(getAreaAffinityScore(rowArea, area) || 0);
          if (v551IsStrongReverse(rowArea, area)) {
            score -= 1200;
          } else {
            score += direction * 1.8;
            score += affinity * 0.7;
            if (v551IsFriendlyDirection(rowArea, area)) score += 140;
          }
        }
      } else {
        score += 20;
      }

      if (existingAnchorArea) {
        const anchorDir = Number(getDirectionAffinityScore(rowArea, existingAnchorArea) || 0);
        const anchorAffinity = Number(getAreaAffinityScore(rowArea, existingAnchorArea) || 0);
        if (v551IsStrongReverse(rowArea, existingAnchorArea)) {
          score -= 1800;
        } else {
          score += anchorDir * 4.2;
          score += anchorAffinity * 1.2;
        }
        score -= Math.abs(existingAnchorDistance - rowDistance) * 0.9;
      }

      const monthly = monthlyMap.get(Number(vehicleId)) || {};
      const avg = Number(monthly.averageDistance || monthly.avgDistance || 0);
      score -= avg * 0.08;
      score -= Number(state.count || 0) * 5;

      if (!best || score > best.score) best = { vehicleId:Number(vehicleId), score };
    }
    return best?.vehicleId || null;
  }

  function v551RebundleAssignments(assignments, items, vehicles, monthlyMap) {
    let rows = Array.isArray(assignments) ? assignments.map(r => ({ ...r })) : [];
    if (!rows.length) return rows;

    const grouped = new Map();
    rows.forEach(row => {
      const hour = Number(row?.actual_hour ?? 0);
      const area = normalizeAreaLabel(row.destination_area || row.cluster_area || row.planned_area || row.casts?.area || '');
      const sig = v551AreaSignature(area);
      const key = `${hour}__${sig.canonical || sig.group || area}`;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(row);
    });

    for (const bundleRows of grouped.values()) {
      if (bundleRows.length < 2) continue;
      const first = bundleRows[0];
      const roundTrip = v551RoundTripMinutesForRow(first);
      if (roundTrip < 60) continue;
      const hour = Number(first.actual_hour ?? 0);
      const vehicleIds = [...new Set(bundleRows.map(r => Number(r.vehicle_id || 0)).filter(Boolean))];
      if (vehicleIds.length <= 1) continue;
      const vehicleStates = v551VehicleStateSummary(rows, vehicles, hour);
      const targetVehicleId = v551PickBundleVehicle(bundleRows, vehicleStates, vehicles, monthlyMap);
      if (!targetVehicleId) continue;
      rows = rows.map(r => {
        if (bundleRows.some(br => Number(br.id) === Number(r.id))) {
          return { ...r, vehicle_id: targetVehicleId };
        }
        return r;
      });
    }

    return rows;
  }

  optimizeAssignments = function(items, vehicles, monthlyMap) {
    let rows = _THEMIS_V55_BASE_optimizeAssignments.apply(this, arguments);
    if (!Array.isArray(rows)) rows = [];
    rows = v551RebundleAssignments(rows, items, vehicles, monthlyMap || new Map());
    return rows;
  };

  const _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe = typeof buildRotationTimelineHtmlSafe === 'function' ? buildRotationTimelineHtmlSafe : null;
  buildRotationTimelineHtmlSafe = function(vehicles, activeItems) {
    try {
      const timeline = (Array.isArray(vehicles) ? vehicles : [])
        .map(vehicle => {
          const rows = (Array.isArray(activeItems) ? activeItems : []).filter(item => sameVehicleAssignmentId(item?.vehicle_id, vehicle?.id));
          if (!rows.length) return null;
          const orderedRows = (typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function')
            ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
            : rows;
          const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
          const summary = getVehicleDailySummary(vehicle, orderedRows);
          return {
            name: vehicle?.driver_name || vehicle?.plate_number || '-',
            returnAfterLabel: forecast?.returnAfterLabel || '-',
            nextRunTime: forecast?.predictedReadyTime || '-',
            rotationMinutes: Number(forecast?.predictedReturnMinutes || 0),
            totalKm: Number(summary?.totalKm || 0),
            totalJobs: Number(summary?.jobCount || 0)
          };
        })
        .filter(Boolean);
      if (!timeline.length) return '';
      return `
        <div class="panel-card" style="margin-bottom:16px;">
          <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
          <div style="display:flex; flex-wrap:wrap; gap:8px;">
            ${timeline.map(item => `
              <div class="chip" style="padding:8px 12px;">
                <strong>${escapeHtml(item.name)}</strong>
                / 戻り ${escapeHtml(item.returnAfterLabel)}
                / 次便可能 ${escapeHtml(item.nextRunTime)}
                / 回転時間 ${Number(item.rotationMinutes || 0)}分
                / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
                / ${Number(item.totalJobs || 0)}件
              </div>
            `).join('')}
          </div>
        </div>
      `;
    } catch (e) {
      console.error('buildRotationTimelineHtmlSafe v5.5.1 error:', e);
      return _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe ? _THEMIS_V55_BASE_buildRotationTimelineHtmlSafe.apply(this, arguments) : '';
    }
  };

  const _THEMIS_V55_BASE_getVehicleRotationForecastSafe = getVehicleRotationForecastSafe;
  getVehicleRotationForecastSafe = function(vehicle, orderedRows) {
    const forecast = _THEMIS_V55_BASE_getVehicleRotationForecastSafe.apply(this, arguments) || {};
    if (!forecast.returnAfterLabel && Number.isFinite(forecast.predictedReturnMinutes)) {
      forecast.returnAfterLabel = `${Math.round(Number(forecast.predictedReturnMinutes || 0))}分後`;
    }
    if (!forecast.predictedReadyTime) forecast.predictedReadyTime = '-';
    return forecast;
  };
})();
/* ===== THEMIS v5.5.1 patch end ===== */


/* ===== THEMIS v5.5.3 patch start ===== */
(function(){
  function v553RoundTripMinutesForRow(row) {
    const stored = getStoredTravelMinutes(row?.casts?.travel_minutes ?? row?.travel_minutes);
    if (stored > 0) return Math.max(0, Math.round(stored * 2));
    const km = Number(row?.distance_km || row?.casts?.distance_km || 0);
    const area = normalizeAreaLabel(row?.destination_area || row?.cluster_area || row?.planned_area || row?.casts?.area || '');
    return Math.max(0, Math.round(estimateFallbackTravelMinutes(km, area) * 2));
  }

  function v553AreaNorm(rowOrArea) {
    if (typeof rowOrArea === 'string') return normalizeAreaLabel(rowOrArea || '');
    return normalizeAreaLabel(rowOrArea?.destination_area || rowOrArea?.cluster_area || rowOrArea?.planned_area || rowOrArea?.casts?.area || '');
  }

  function v553FriendlyArea(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    const ca = getCanonicalArea(a) || '';
    const cb = getCanonicalArea(b) || '';
    const ga = getAreaDisplayGroup(a) || '';
    const gb = getAreaDisplayGroup(b) || '';
    if (ca && cb && ca === cb) return true;
    if (ga && gb && ga === gb && getDirectionAffinityScore(a, b) > -20) return true;
    return getAreaAffinityScore(a, b) >= 60 || getDirectionAffinityScore(a, b) >= 22;
  }

  function v553HardReverseArea(areaA, areaB) {
    const a = normalizeAreaLabel(areaA || '');
    const b = normalizeAreaLabel(areaB || '');
    if (!a || !b) return false;
    return getAreaAffinityScore(a, b) <= 26 || getDirectionAffinityScore(a, b) <= -34;
  }

  function v553BuildVehicleStates(assignments, vehicles, hour) {
    const map = new Map();
    (Array.isArray(vehicles) ? vehicles : []).forEach(v => {
      map.set(Number(v.id), { vehicle: v, count: 0, rows: [], areas: [] });
    });
    (Array.isArray(assignments) ? assignments : [])
      .filter(r => Number(r?.actual_hour ?? 0) === Number(hour))
      .forEach(r => {
        const vid = Number(r?.vehicle_id || 0);
        if (!map.has(vid)) return;
        const state = map.get(vid);
        state.count += 1;
        state.rows.push(r);
        const area = v553AreaNorm(r);
        if (area) state.areas.push(area);
      });
    return map;
  }

  function v553FindLongFriendlyComponents(hourRows) {
    const rows = (Array.isArray(hourRows) ? hourRows : []).filter(r => v553RoundTripMinutesForRow(r) >= 55);
    const visited = new Set();
    const components = [];
    for (const row of rows) {
      const id = Number(row?.id || 0);
      if (!id || visited.has(id)) continue;
      const stack = [row];
      const component = [];
      visited.add(id);
      while (stack.length) {
        const cur = stack.pop();
        component.push(cur);
        const curArea = v553AreaNorm(cur);
        for (const other of rows) {
          const oid = Number(other?.id || 0);
          if (!oid || visited.has(oid)) continue;
          const otherArea = v553AreaNorm(other);
          if (v553FriendlyArea(curArea, otherArea) && !v553HardReverseArea(curArea, otherArea)) {
            visited.add(oid);
            stack.push(other);
          }
        }
      }
      if (component.length >= 2) components.push(component);
    }
    return components;
  }

  function v553VehicleCanAcceptComponent(state, vehicle, componentRows) {
    const seat = Math.max(1, Number(vehicle?.seat_capacity || 4));
    const componentIds = new Set((componentRows || []).map(r => Number(r?.id || 0)));
    const currentRows = Array.isArray(state?.rows) ? state.rows : [];
    const nonComponentRows = currentRows.filter(r => !componentIds.has(Number(r?.id || 0)));
    if (nonComponentRows.length + componentRows.length > seat) return false;
    const existingAreas = nonComponentRows.map(v553AreaNorm).filter(Boolean);
    for (const row of componentRows) {
      const area = v553AreaNorm(row);
      if (existingAreas.some(a => v553HardReverseArea(area, a))) return false;
    }
    return true;
  }

  function v553PickTargetVehicle(componentRows, states, monthlyMap) {
    const idsInComponent = new Set(componentRows.map(r => Number(r?.id || 0)));
    const componentAreas = componentRows.map(v553AreaNorm).filter(Boolean);
    let best = null;
    for (const [vehicleId, state] of states.entries()) {
      const vehicle = state.vehicle;
      if (!vehicle) continue;
      if (!v553VehicleCanAcceptComponent(state, vehicle, componentRows)) continue;

      const currentRows = Array.isArray(state.rows) ? state.rows : [];
      const componentAlreadyHere = currentRows.filter(r => idsInComponent.has(Number(r?.id || 0))).length;
      const nonComponentAreas = currentRows
        .filter(r => !idsInComponent.has(Number(r?.id || 0)))
        .map(v553AreaNorm)
        .filter(Boolean);

      let score = 0;
      score += componentAlreadyHere * 900;
      if (componentAlreadyHere > 0) score += 250;

      for (const compArea of componentAreas) {
        for (const area of nonComponentAreas) {
          if (v553FriendlyArea(compArea, area)) score += 120;
          if (v553HardReverseArea(compArea, area)) score -= 600;
        }
      }

      const monthly = monthlyMap?.get(Number(vehicleId)) || {};
      score -= Number(monthly.averageDistance || 0) * 0.04;
      score -= Number(state.count || 0) * 6;

      if (!best || score > best.score) best = { vehicleId: Number(vehicleId), score };
    }
    return best?.vehicleId || null;
  }

  function v553RebundleAssignments(assignments, vehicles, monthlyMap) {
    let rows = Array.isArray(assignments) ? assignments.map(r => ({ ...r })) : [];
    if (!rows.length) return rows;

    const hours = [...new Set(rows.map(r => Number(r?.actual_hour ?? 0)))];
    for (const hour of hours) {
      const hourRows = rows.filter(r => Number(r?.actual_hour ?? 0) === Number(hour));
      const components = v553FindLongFriendlyComponents(hourRows);
      if (!components.length) continue;

      for (const componentRows of components) {
        const currentVehicles = [...new Set(componentRows.map(r => Number(r?.vehicle_id || 0)).filter(Boolean))];
        if (currentVehicles.length <= 1) continue;
        const states = v553BuildVehicleStates(rows, vehicles, hour);
        const targetVehicleId = v553PickTargetVehicle(componentRows, states, monthlyMap);
        if (!targetVehicleId) continue;
        const componentIds = new Set(componentRows.map(r => Number(r?.id || 0)));
        rows = rows.map(r => componentIds.has(Number(r?.id || 0)) ? { ...r, vehicle_id: targetVehicleId } : r);
      }
    }

    return rows;
  }

  function v553FormatMinutes(totalMinutes) {
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    return `${safe}分`;
  }

  function v553FormatClock(totalMinutes) {
    if (typeof formatClockTimeFromMinutesGlobal === 'function') return formatClockTimeFromMinutesGlobal(totalMinutes);
    if (typeof formatClockTimeFromMinutes === 'function') return formatClockTimeFromMinutes(totalMinutes);
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    const h = Math.floor(safe / 60) % 24;
    const m = safe % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }

  const _THEMIS_V553_BASE_getVehicleRotationForecastSafe = getVehicleRotationForecastSafe;
  getVehicleRotationForecastSafe = function(vehicle, orderedRows) {
    const forecast = _THEMIS_V553_BASE_getVehicleRotationForecastSafe.apply(this, arguments) || {};
    const rows = Array.isArray(orderedRows) ? orderedRows.filter(Boolean) : [];
    if (!rows.length) {
      if (!forecast.returnAfterLabel) forecast.returnAfterLabel = '-';
      if (!forecast.predictedReadyTime) forecast.predictedReadyTime = '-';
      return forecast;
    }

    const timeSummary = getRowsTravelTimeSummary(rows);
    const isLastTrip = isDriverLastTripChecked(Number(vehicle?.id || 0));
    const fallbackMinutes = Math.max(0, Math.round(isLastTrip ? timeSummary.sendOnlyMinutes : timeSummary.totalMinutes));
    const runAt = Number(lastAutoDispatchRunAtMinutes || getCurrentClockMinutes() || 0);
    const fallbackReady = v553FormatClock(runAt + fallbackMinutes);

    const missingMinutes = !(Number(forecast.predictedReturnMinutes || 0) > 0);
    const missingReturnLabel = !forecast.returnAfterLabel || forecast.returnAfterLabel === '-' || forecast.returnAfterLabel === '0分後';
    const missingReady = !forecast.predictedReadyTime || forecast.predictedReadyTime === '-';
    const missingRotation = !(Number(forecast.rotationMinutes || 0) > 0);

    if (missingMinutes) forecast.predictedReturnMinutes = fallbackMinutes;
    if (missingReturnLabel) forecast.returnAfterLabel = `${fallbackMinutes}分後`;
    if (missingReady) forecast.predictedReadyTime = fallbackReady;
    if (missingRotation) forecast.rotationMinutes = fallbackMinutes;

    return forecast;
  };

  const _THEMIS_V553_BASE_buildRotationTimelineHtmlSafe = buildRotationTimelineHtmlSafe;
  buildRotationTimelineHtmlSafe = function(vehicles, activeItems) {
    try {
      const timeline = (Array.isArray(vehicles) ? vehicles : [])
        .map(vehicle => {
          const rows = (Array.isArray(activeItems) ? activeItems : []).filter(item => sameVehicleAssignmentId(item?.vehicle_id, vehicle?.id));
          if (!rows.length) return null;
          const orderedRows = (typeof moveManualLastItemsToEnd === 'function' && typeof sortItemsByNearestRoute === 'function')
            ? moveManualLastItemsToEnd(sortItemsByNearestRoute(rows))
            : rows;
          const forecast = getVehicleRotationForecastSafe(vehicle, orderedRows);
          const summary = getVehicleDailySummary(vehicle, orderedRows);
          const rotationMinutes = Number(forecast?.rotationMinutes || forecast?.predictedReturnMinutes || 0);
          return {
            name: vehicle?.driver_name || vehicle?.plate_number || '-',
            returnAfterLabel: forecast?.returnAfterLabel || `${rotationMinutes}分後`,
            nextRunTime: forecast?.predictedReadyTime || '-',
            rotationMinutes,
            totalKm: Number(summary?.totalKm || 0),
            totalJobs: Number(summary?.jobCount || 0)
          };
        })
        .filter(Boolean);
      if (!timeline.length) return '';
      return `
        <div class="panel-card" style="margin-bottom:16px;">
          <h3 style="margin-bottom:10px;">車両稼働タイムライン</h3>
          <div style="display:flex; flex-wrap:wrap; gap:8px;">
            ${timeline.map(item => `
              <div class="chip" style="padding:8px 12px;">
                <strong>${escapeHtml(item.name)}</strong>
                / 戻り ${escapeHtml(item.returnAfterLabel)}
                / 次便可能 ${escapeHtml(item.nextRunTime)}
                / 回転時間 ${Math.round(Number(item.rotationMinutes || 0))}分
                / 累計 ${Number(item.totalKm || 0).toFixed(1)}km
                / ${Number(item.totalJobs || 0)}件
              </div>
            `).join('')}
          </div>
        </div>
      `;
    } catch (e) {
      console.error('buildRotationTimelineHtmlSafe v5.5.3 error:', e);
      return _THEMIS_V553_BASE_buildRotationTimelineHtmlSafe.apply(this, arguments);
    }
  };
})();

function renderDispatchSummaryFromDOM(){
  const resultWrap = document.getElementById("dailyDispatchResult");
  const bar = document.getElementById("dispatchSummaryBar");
  if(!resultWrap || !bar) return;

  const cards = resultWrap.querySelectorAll(".dispatch-card, .vehicle-card, .result-card");

  let vehicleCount = 0;
  let totalJobs = 0;
  let times = [];

  cards.forEach(card=>{
    vehicleCount++;

    const jobs = card.querySelectorAll(".job, .dispatch-job, li");
    totalJobs += jobs.length;

    // 時間っぽいテキスト探す（例: 45分）
    const text = card.innerText;
    const match = text.match(/(\d+)\s*分/);
    if(match){
      times.push(parseInt(match[1]));
    }
  });

  if(vehicleCount === 0){
    bar.innerHTML = "";
    return;
  }

  let min = "-", max = "-", avg = "-";

  if(times.length){
    min = Math.min(...times);
    max = Math.max(...times);
    avg = Math.round(times.reduce((a,b)=>a+b,0)/times.length);
  }

  bar.innerHTML = `
    🚗 ${vehicleCount}台｜
    📦 ${totalJobs}件｜
    ⏱ 最短 ${min}｜
    ⏳ 最長 ${max}｜
    📉 平均 ${avg}
  `;
}

// フック（既存処理に干渉しない）
const _origRender = window.renderDailyDispatchResult;
window.renderDailyDispatchResult = function(){
  if(_origRender) _origRender();
  setTimeout(renderDispatchSummaryFromDOM, 50);
};

/* ===== THEMIS v6.9.13 drop-order unify hotfix start ===== */
(function(){
  function __themisOrderDistance(row) {
    const v = Number(row?.distance_km ?? row?.casts?.distance_km ?? 0);
    if (Number.isFinite(v) && v > 0) return v;
    const pt = (typeof getThemisV54RowPoint === 'function') ? getThemisV54RowPoint(row) : null;
    if (pt && typeof estimateRoadKmBetweenPoints === 'function') {
      return Number(estimateRoadKmBetweenPoints(ORIGIN_LAT, ORIGIN_LNG, pt.lat, pt.lng) || 999999);
    }
    return 999999;
  }

  function __themisOrderTravel(row) {
    if (typeof getThemisV54StoredTravelMinutes === 'function') {
      const v = Number(getThemisV54StoredTravelMinutes(row) || 0);
      if (Number.isFinite(v) && v > 0) return v;
    }
    const v = Number(row?.travel_minutes ?? row?.casts?.travel_minutes ?? 0);
    return Number.isFinite(v) && v > 0 ? v : 999999;
  }

  function __themisOrderId(row) {
    const v = normalizeDispatchEntityId(row?.id || row?.cast_id || row?.casts?.id || null);
    return Number.isFinite(v) ? v : 0;
  }

  function __themisCompareByOrigin(a, b) {
    const da = __themisOrderDistance(a);
    const db = __themisOrderDistance(b);
    if (da !== db) return da - db;

    const ta = __themisOrderTravel(a);
    const tb = __themisOrderTravel(b);
    if (ta !== tb) return ta - tb;

    return __themisOrderId(a) - __themisOrderId(b);
  }

  function __themisCompareFromPrevious(prev, a, b) {
    const prevD = __themisOrderDistance(prev);
    const da = Math.abs(__themisOrderDistance(a) - prevD);
    const db = Math.abs(__themisOrderDistance(b) - prevD);
    if (da !== db) return da - db;

    const oa = __themisOrderDistance(a);
    const ob = __themisOrderDistance(b);
    if (oa !== ob) return oa - ob;

    const ta = __themisOrderTravel(a);
    const tb = __themisOrderTravel(b);
    if (ta !== tb) return ta - tb;

    return __themisOrderId(a) - __themisOrderId(b);
  }

  sortItemsByNearestRoute = function(items) {
    const remaining = [...(items || [])].filter(Boolean);
    if (!remaining.length) return [];

    remaining.sort(__themisCompareByOrigin);
    const ordered = [remaining.shift()];

    while (remaining.length) {
      const prev = ordered[ordered.length - 1];
      remaining.sort((a, b) => __themisCompareFromPrevious(prev, a, b));
      ordered.push(remaining.shift());
    }

    return ordered;
  };
})();
/* ===== THEMIS v6.9.13 drop-order unify hotfix end ===== */


// ===== THEMIS dispatchCore isolated bridge =====

function optimizeAssignments(items, vehicles, monthlyMap, options = {}) {
  return callDispatchCoreSafe(items, vehicles, monthlyMap, options);
}

function callDispatchCoreSafe(items, vehicles, monthlyMap, options = {}) {
  if (window.DispatchCore && typeof window.DispatchCore.optimizeAssignments === "function") {
    return window.DispatchCore.optimizeAssignments(items, vehicles, monthlyMap, options);
  }
  return [];
}

function runSimulationDispatchPreview() {
  if (typeof window.__THEMIS_SIMULATION_DISPATCH_PREVIEW_OVERRIDE__ === "function") {
    return window.__THEMIS_SIMULATION_DISPATCH_PREVIEW_OVERRIDE__();
  }
  if (typeof window.renderSimulationDispatchPreview === 'function') {
    return window.renderSimulationDispatchPreview();
  }
  alert('試算プレビュー関数が見つかりません');
  return [];
}

window.__THEMIS_SIMULATION_DISPATCH_PREVIEW_OVERRIDE__ = function __THEMIS_SIMULATION_DISPATCH_PREVIEW_OVERRIDE__() {
  const hour = Number(els.simulationSlotSelect?.value ?? (typeof getSimulationSlotHourSafe === 'function' ? getSimulationSlotHourSafe() : null) ?? getOperationBaseHour());
  if (typeof setSimulationSlotHourSafe === 'function') setSimulationSlotHourSafe(hour);

  const includePlanInflow = Boolean(els.simulationIncludePlanInflow?.checked);
  const built = typeof window.buildSimulationRowsForHour === 'function'
    ? window.buildSimulationRowsForHour(hour, { includePlanInflow })
    : (typeof __simulationBuildSimulationRowsForHourFallback === 'function'
        ? __simulationBuildSimulationRowsForHourFallback(hour, { includePlanInflow })
        : { rows: [], summary: { slotPlanCount: 0, inflowPlanCount: 0 } });

  const sourceRows = Array.isArray(built?.rows) ? built.rows.filter(Boolean) : [];
  const vehicles = Array.isArray(getSelectedVehiclesForToday?.()) ? getSelectedVehiclesForToday().filter(Boolean) : [];

  if (!vehicles.length) {
    alert('可能車両を選択してください');
    return [];
  }
  if (!sourceRows.length) {
    alert('この便の試算対象がありません');
    return [];
  }

  const tempIdBase = hour * 100000 + 90000000;
  const rowsForDispatch = sourceRows.map((row, index) => ({
    ...row,
    id: tempIdBase + index + 1,
    actual_hour: Number(row?.actual_hour ?? row?.plan_hour ?? hour),
    plan_hour: Number(row?.plan_hour ?? row?.actual_hour ?? hour)
  }));
  const sourceRowByTempId = new Map(rowsForDispatch.map((row, index) => [Number(row.id), sourceRows[index]]));

  const monthlyMap = typeof buildMonthlyDistanceMapForCurrentMonth === 'function'
    ? buildMonthlyDistanceMapForCurrentMonth()
    : new Map();

  let assignments = [];
  try {
    if (window.DispatchCore?.optimizeAssignments) {
      assignments = window.DispatchCore.optimizeAssignments(rowsForDispatch, vehicles, monthlyMap, { mode: 'simulation_preview' });
    } else if (typeof optimizeAssignments === 'function') {
      assignments = optimizeAssignments(rowsForDispatch, vehicles, monthlyMap, { mode: 'simulation_preview' });
    }
  } catch (error) {
    console.error('simulation dispatch preview error:', error);
    assignments = [];
  }
  if (!Array.isArray(assignments)) assignments = [];

  const mappedAssignments = assignments.map(assignment => ({
    ...assignment,
    item_id: Number(assignment?.item_id || 0),
    source_row: sourceRowByTempId.get(Number(assignment?.item_id || 0)) || null
  }));
  const assignedVehicleByTempId = new Map(mappedAssignments.map(a => [Number(a.item_id), Number(a?.vehicle_id || 0)]));
  const previewRows = rowsForDispatch.map((row, index) => ({
    ...sourceRows[index],
    vehicle_id: assignedVehicleByTempId.get(Number(row.id)) || 0
  }));

  const assignedCount = previewRows.filter(row => Number(row?.vehicle_id || 0) > 0).length;
  const unassignedRows = previewRows.filter(row => Number(row?.vehicle_id || 0) <= 0);

  if (typeof setLastSimulationResultSafe === 'function') {
    setLastSimulationResultSafe({
      type: 'dispatch_preview',
      hour,
      assignments: mappedAssignments,
      rows: previewRows,
      summary: built?.summary || { slotPlanCount: 0, inflowPlanCount: 0 }
    });
  }

  if (els.simulationDiagnosis) {
    const diagStatus = unassignedRows.length ? 'warn' : 'ok';
    const buildPill = typeof buildStatusPill === 'function'
      ? buildStatusPill
      : ((label, value) => `<span class="status-pill ok"><span class="status-pill-label">${escapeHtml(label)}</span><span class="status-pill-value">${escapeHtml(value)}</span></span>`);
    els.simulationDiagnosis.className = `hybrid-diagnosis ${diagStatus}`;
    els.simulationDiagnosis.innerHTML = `
      <div class="hybrid-state-row">
        ${buildPill('便', `${getHourLabel(hour)}`, diagStatus)}
        ${buildPill('対象', `${sourceRows.length}名`, diagStatus)}
        ${buildPill('車両', `${vehicles.length}台`, vehicles.length ? 'ok' : 'warn')}
        ${buildPill('定員', `${vehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0)}`, 'ok')}
      </div>
      <div class="hybrid-legend-grid">
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">判定</div>
          <div class="hybrid-legend-value">${unassignedRows.length ? `未割当 ${unassignedRows.length}名` : 'OK'}</div>
        </div>
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">方面系統</div>
          <div class="hybrid-legend-value">${new Set(previewRows.map(row => getAreaDisplayGroup(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '無し')))).size}</div>
        </div>
        <div class="hybrid-legend-card">
          <div class="hybrid-legend-title">次便NG車両</div>
          <div class="hybrid-legend-value">${(typeof getVehicleDeadNamesForHourSafe === 'function' ? getVehicleDeadNamesForHourSafe(hour) : []).join(' / ') || 'なし'}</div>
        </div>
      </div>
    `;
  }

  if (els.simulationPreview) {
    const assignedVehicleMap = new Map(vehicles.map(vehicle => [Number(vehicle?.id || 0), { vehicle, rows: [] }]));
    previewRows.forEach(row => {
      const vehicleId = Number(row?.vehicle_id || 0);
      if (vehicleId > 0 && assignedVehicleMap.has(vehicleId)) {
        assignedVehicleMap.get(vehicleId).rows.push(row);
      }
    });

    const assignedBlocks = [...assignedVehicleMap.values()]
      .filter(entry => entry.rows.length > 0)
      .map(entry => {
        const name = escapeHtml(entry.vehicle?.driver_name || entry.vehicle?.plate_number || '-');
        const rowsHtml = entry.rows
          .sort((a, b) => Number(a?.distance_km || 0) - Number(b?.distance_km || 0))
          .map(row => `<div class="sim-mini-chip"><span class="sim-mini-name">${escapeHtml(row?.casts?.name || '-')}</span><span class="sim-mini-area">${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span></div>`)
          .join('');
        return `
          <div class="sim-vehicle-card">
            <div class="sim-preview-head">
              <h4 class="sim-preview-title">${name}</h4>
              <span class="chip">${entry.rows.length}名</span>
            </div>
            <div class="sim-mini-chip-grid">${rowsHtml}</div>
          </div>
        `;
      })
      .join('');

    const unassignedHtml = unassignedRows.length
      ? `
        <div class="sim-vehicle-card">
          <div class="sim-preview-head">
            <h4 class="sim-preview-title">未割当</h4>
            <span class="chip">${unassignedRows.length}名</span>
          </div>
          <div class="sim-mini-chip-grid">
            ${unassignedRows.map(row => `<div class="sim-mini-chip"><span class="sim-mini-name">${escapeHtml(row?.casts?.name || '-')}</span><span class="sim-mini-area">${escapeHtml(normalizeAreaLabel(row?.destination_area || row?.planned_area || row?.casts?.area || '-'))}</span></div>`).join('')}
          </div>
        </div>
      `
      : '';

    els.simulationPreview.className = 'simulation-preview';
    els.simulationPreview.innerHTML = `
      <div class="sim-preview-head">
        <h4 class="sim-preview-title">試算対象一覧</h4>
        <span class="chip">${getHourLabel(hour)}</span>
      </div>
      <div class="sim-preview-meta">対象便予定 ${Number(built?.summary?.slotPlanCount || 0)}名 / 前便未処理予定 ${Number(built?.summary?.inflowPlanCount || 0)}名 / 割当 ${assignedCount}名 / 未割当 ${unassignedRows.length}名</div>
      ${assignedBlocks || '<div class="muted">割当車両はありません</div>'}
      ${unassignedHtml}
    `;
  }

  return mappedAssignments;
};


/* ===== DROP OFF monthly daily-runs ui bridge v1 start ===== */
(function(){
  let __dropoffMonthlyUiStatsMap = new Map();

  function __dropoffMonthRange(dateStr) {
    const base = String(dateStr || (els?.dispatchDate?.value || todayStr?.() || new Date().toISOString().slice(0,10)));
    const monthKey = typeof getMonthKey === 'function' ? getMonthKey(base) : base.slice(0, 7);
    const startDate = `${monthKey}-01`;
    const start = new Date(`${startDate}T00:00:00`);
    const next = new Date(start.getFullYear(), start.getMonth() + 1, 1);
    const nextStartDate = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}-01`;
    return { startDate, nextStartDate, baseDate: base };
  }

  function __dropoffResolveMonthlyLocalVehicleId(row) {
    const resolved = typeof resolveVehicleLocalNumericId === 'function'
      ? resolveVehicleLocalNumericId(row?.vehicle_id)
      : Number(row?.vehicle_id || 0);
    return Number(resolved || 0);
  }

  async function __dropoffFetchMonthlyDailyRunRows(baseDate) {
    const range = __dropoffMonthRange(baseDate);
    if (!supabaseClient?.from) return [];

    const teamId = typeof resolveWorkspaceTeamIdForDailyRuns === 'function'
      ? await resolveWorkspaceTeamIdForDailyRuns()
      : null;

    let query = supabaseClient
      .from(getVehicleDailyRunsTableName())
      .select('team_id, vehicle_id, run_date, reference_distance_km, trip_count, drive_minutes, is_workday')
      .gte('run_date', range.startDate)
      .lt('run_date', range.nextStartDate)
      .order('run_date', { ascending: true });

    if (teamId) query = query.eq('team_id', teamId);

    const { data, error } = await query;
    if (error) {
      if (typeof isMissingTableError === 'function' && isMissingTableError(error)) {
        console.warn('dropoff_vehicle_daily_runs is not ready:', error);
        return [];
      }
      console.error('monthly daily-runs fetch error:', error);
      return [];
    }

    return Array.isArray(data) ? data : [];
  }

  function __dropoffBuildMonthlyStatsMap(rows) {
    const map = new Map();
    (Array.isArray(rows) ? rows : []).forEach(row => {
      const vehicleId = __dropoffResolveMonthlyLocalVehicleId(row);
      if (!(vehicleId > 0)) return;

      const prev = map.get(vehicleId) || {
        totalDistance: 0,
        workedDays: 0,
        avgDistance: 0
      };

      prev.totalDistance += Number(row?.reference_distance_km || 0);
      if (row?.is_workday !== false) prev.workedDays += 1;
      map.set(vehicleId, prev);
    });

    map.forEach(stats => {
      stats.totalDistance = Number(Number(stats.totalDistance || 0).toFixed(1));
      stats.workedDays = Number(stats.workedDays || 0);
      stats.avgDistance = stats.workedDays > 0
        ? Number((stats.totalDistance / stats.workedDays).toFixed(1))
        : 0;
    });

    return map;
  }

  async function __dropoffRefreshMonthlyUi(baseDate) {
    const rows = await __dropoffFetchMonthlyDailyRunRows(baseDate);
    __dropoffMonthlyUiStatsMap = __dropoffBuildMonthlyStatsMap(rows);
    if (typeof renderHomeMonthlyVehicleList === 'function') renderHomeMonthlyVehicleList();
    if (typeof renderVehiclesTable === 'function') renderVehiclesTable();
    if (typeof renderDailyVehicleChecklist === 'function') renderDailyVehicleChecklist();
  }

  function __dropoffGetMonthlyStatsMap(baseDate) {
    if (__dropoffMonthlyUiStatsMap instanceof Map && __dropoffMonthlyUiStatsMap.size > 0) {
      return __dropoffMonthlyUiStatsMap;
    }
    return new Map();
  }

  window.__dropoffGetMonthlyStatsMap = __dropoffGetMonthlyStatsMap;

  renderHomeMonthlyVehicleList = function() {
    if (!els?.homeMonthlyVehicleList) return;
    const statsMap = __dropoffGetMonthlyStatsMap();
    els.homeMonthlyVehicleList.innerHTML = '';
    if (!allVehiclesCache.length) {
      els.homeMonthlyVehicleList.innerHTML = `<div class="chip">車両なし</div>`;
      return;
    }
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = statsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const row = document.createElement('div');
      row.className = 'home-monthly-item';
      row.innerHTML = `
        <span class="chip">${escapeHtml(vehicle.driver_name || vehicle.plate_number || '-')}</span>
        <span class="chip">${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))}</span>
        <span class="chip">帰宅:${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))}</span>
        <span class="chip">月間:${Number(stats.totalDistance || 0).toFixed(1)}km</span>
        <span class="chip">出勤:${Number(stats.workedDays || 0)}日</span>
        <span class="chip">平均:${Number(stats.avgDistance || 0).toFixed(1)}km</span>`;
      els.homeMonthlyVehicleList.appendChild(row);
    });
  };

  renderVehiclesTable = function() {
    if (!els?.vehiclesTableBody) return;
    const isReadonlyUser = isReadonlyUserRole();
    const statsMap = __dropoffGetMonthlyStatsMap();
    els.vehiclesTableBody.innerHTML = '';
    const vehiclesTable = els.vehiclesTableBody.closest('table');
    const vehiclesHeaderRow = vehiclesTable?.querySelector('thead tr');
    if (vehiclesHeaderRow) {
      vehiclesHeaderRow.innerHTML = `
        <th>ドライバー</th>
        <th>車両</th>
        <th>担当方面</th>
        <th>帰宅方面</th>
        <th>定員</th>
        <th>月間距離(km)</th>
        <th>出勤日数</th>
        <th>1日平均(km)</th>
        <th>操作</th>
      `;
    }
    if (!allVehiclesCache.length) {
      els.vehiclesTableBody.innerHTML = `<tr><td colspan="9" class="muted">車両がありません</td></tr>`;
      return;
    }
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = statsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const tr = document.createElement('tr');
      const actionsHtml = isReadonlyUser
        ? '<span class="muted">閲覧専用</span>'
        : `
        <button class="btn ghost vehicle-edit-btn" data-id="${vehicle.id}">編集</button>
        <button class="btn danger vehicle-delete-btn" data-id="${vehicle.id}">削除</button>`;
      tr.innerHTML = `
      <tr>
      <td>${escapeHtml(vehicle.driver_name || '-')}</td>
      <td>${escapeHtml(vehicle.plate_number || '-')}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))}</td>
      <td>${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))}</td>
      <td>${vehicle.seat_capacity ?? '-'}</td>
      <td>${Number(stats.totalDistance || 0).toFixed(1)}</td>
      <td>${Number(stats.workedDays || 0)}</td>
      <td>${Number(stats.avgDistance || 0).toFixed(1)}</td>
      <td class="actions-cell">${actionsHtml}</td>`;
      els.vehiclesTableBody.appendChild(tr);
    });
    if (!isReadonlyUser) {
      els.vehiclesTableBody.querySelectorAll('.vehicle-edit-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const vehicle = allVehiclesCache.find(x => Number(x.id) === Number(btn.dataset.id));
          if (vehicle) fillVehicleForm(vehicle);
        });
      });
      els.vehiclesTableBody.querySelectorAll('.vehicle-delete-btn').forEach(btn => {
        btn.addEventListener('click', async () => deleteVehicle(Number(btn.dataset.id)));
      });
    }
  };

  renderDailyVehicleChecklist = function() {
    if (!els?.dailyVehicleChecklist) return;
    els.dailyVehicleChecklist.innerHTML = '';
    if (!allVehiclesCache.length) {
      els.dailyVehicleChecklist.innerHTML = `<div class="muted">車両がありません</div>`;
      return;
    }
    const monthlyStatsMap = __dropoffGetMonthlyStatsMap();
    const header = document.createElement('div');
    header.className = 'vehicle-check-header';
    header.innerHTML = `<div class="vehicle-check-header-info"></div><div class="vehicle-check-header-col">可能車両</div><div class="vehicle-check-header-col">ラスト便</div>`;
    els.dailyVehicleChecklist.appendChild(header);
    getSortedVehiclesForDisplay().forEach(vehicle => {
      const stats = monthlyStatsMap.get(Number(vehicle.id)) || { totalDistance: 0, workedDays: 0, avgDistance: 0 };
      const avgDistanceText = `${Number(stats.avgDistance || 0).toFixed(1)}km`;
      const row = document.createElement('div');
      row.className = 'vehicle-check-item';
      row.innerHTML = `
        <div class="vehicle-check-info">
          <div class="vehicle-check-name">${escapeHtml(vehicle.driver_name || '-')}</div>
          <div class="vehicle-check-car">車両 ${escapeHtml(vehicle.plate_number || '-')}</div>
          <div class="vehicle-check-meta">担当 ${escapeHtml(normalizeAreaLabel(vehicle.vehicle_area || '-'))} / 帰宅 ${escapeHtml(normalizeAreaLabel(vehicle.home_area || '-'))} / 定員 ${vehicle.seat_capacity ?? '-'} / 1日平均距離 ${avgDistanceText}</div>
        </div>
        <label class="vehicle-check-toggle vehicle-check-toggle-work">
          <input class="vehicle-check-input" type="checkbox" data-id="${vehicle.id}" ${activeVehicleIdsForToday.has(Number(vehicle.id)) ? 'checked' : ''} />
          <span>可能車両</span>
        </label>
        <label class="vehicle-check-toggle vehicle-check-toggle-last">
          <input class="driver-last-trip-input" type="checkbox" data-id="${vehicle.id}" ${isDriverLastTripChecked(vehicle.id) ? 'checked' : ''} />
          <span>ラスト便</span>
        </label>`;
      els.dailyVehicleChecklist.appendChild(row);
    });
    renderOperationAndSimulationUI();
    els.dailyVehicleChecklist.querySelectorAll('.vehicle-check-input').forEach(input => {
      input.addEventListener('change', () => {
        const id = Number(input.dataset.id);
        if (input.checked) activeVehicleIdsForToday.add(id); else activeVehicleIdsForToday.delete(id);
        renderDailyMileageInputs();
        renderDailyDispatchResult();
        renderDailyVehicleChecklist();
      });
    });
    els.dailyVehicleChecklist.querySelectorAll('.driver-last-trip-input').forEach(input => {
      input.addEventListener('change', () => {
        const id = Number(input.dataset.id);
        setDriverLastTripChecked(id, input.checked);
        if (input.checked && !activeVehicleIdsForToday.has(id)) activeVehicleIdsForToday.add(id);
        renderDailyMileageInputs();
        renderDailyDispatchResult();
        renderDailyVehicleChecklist();
      });
    });
  };

  const __dropoffRefreshNames = ['saveDailyMileageReports','confirmDailyToMonthly','resetMonthlySummary','syncDateAndReloadFromDispatchDate','syncDateAndReloadFromPlanDate','syncDateAndReloadFromActualDate','loadHomeAndAll','previewDriverMileageReport'];
  __dropoffRefreshNames.forEach(name => {
    const orig = globalThis[name];
    if (typeof orig !== 'function') return;
    globalThis[name] = async function(...args) {
      const result = await orig.apply(this, args);
      try { await __dropoffRefreshMonthlyUi(els?.dispatchDate?.value || todayStr()); } catch (e) { console.error(`${name} monthly refresh error:`, e); }
      return result;
    };
  });

  window.addEventListener('load', () => {
    setTimeout(() => { __dropoffRefreshMonthlyUi(els?.dispatchDate?.value || todayStr()).catch(err => console.error('monthly refresh on load error:', err)); }, 300);
  });

  window.__dropoffRefreshMonthlyUi = __dropoffRefreshMonthlyUi;
})();
/* ===== DROP OFF monthly daily-runs ui bridge v1 end ===== */


function getVehicleAreaMatchScore(vehicle, area) {
  const targetRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(area || "")
    : String(area || "").trim();
  if (!targetRaw || targetRaw === "無し") return 0;

  const targetCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(targetRaw) || targetRaw)
    : targetRaw;
  const targetGroup = (typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(targetRaw))
    ? targetRaw
    : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(targetRaw) : targetCanonical);

  const vehicleAreaRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(vehicle?.vehicle_area || "")
    : String(vehicle?.vehicle_area || "").trim();
  const homeAreaRaw = typeof normalizeAreaLabel === "function"
    ? normalizeAreaLabel(vehicle?.home_area || "")
    : String(vehicle?.home_area || "").trim();

  const vehicleCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(vehicleAreaRaw) || vehicleAreaRaw)
    : vehicleAreaRaw;
  const homeCanonical = typeof getCanonicalArea === "function"
    ? (getCanonicalArea(homeAreaRaw) || homeAreaRaw)
    : homeAreaRaw;

  const vehicleGroup = vehicleAreaRaw
    ? ((typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(vehicleAreaRaw))
        ? vehicleAreaRaw
        : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(vehicleAreaRaw) : vehicleCanonical))
    : vehicleCanonical;
  const homeGroup = homeAreaRaw
    ? ((typeof THEMIS_DISPLAY_GROUPS !== "undefined" && THEMIS_DISPLAY_GROUPS && THEMIS_DISPLAY_GROUPS.has(homeAreaRaw))
        ? homeAreaRaw
        : (typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup(homeAreaRaw) : homeCanonical))
    : homeCanonical;

  function calcScore(baseCanonical, baseGroup, weight, useHomeReverseGuard) {
    if (!baseCanonical && !baseGroup) return -999;

    let score = 0;
    if (baseCanonical && targetCanonical && baseCanonical === targetCanonical) {
      score = 100;
    } else if (baseGroup && targetGroup && baseGroup === targetGroup) {
      score = 88;
    } else {
      const affinity = typeof getAreaAffinityScore === "function"
        ? Number(getAreaAffinityScore(baseCanonical || baseGroup, targetCanonical) || 0)
        : 0;
      const direction = typeof getDirectionAffinityScore === "function"
        ? Number(getDirectionAffinityScore(baseCanonical || baseGroup, targetCanonical) || 0)
        : 0;

      score = affinity * 0.72 + direction * 0.34;
      if (direction <= -38) score -= 55;
      if (direction <= -95) score -= 95;
    }

    if (typeof isHardReverseMixForRoute === "function" && isHardReverseMixForRoute(baseCanonical || baseGroup, targetCanonical)) {
      score -= 130;
    }
    if (useHomeReverseGuard && typeof isHardReverseForHome === "function" && isHardReverseForHome(targetCanonical, baseCanonical || baseGroup)) {
      score -= 120;
    }

    return score * weight;
  }

  const vehicleScore = calcScore(vehicleCanonical, vehicleGroup, 1.0, false);
  const homeScore = calcScore(homeCanonical, homeGroup, 0.72, true);

  let best = Math.max(vehicleScore, homeScore, 0);

  if (vehicleGroup && homeGroup && vehicleGroup === homeGroup && vehicleGroup === targetGroup) {
    best += 8;
  }

  return Math.max(-160, Math.min(110, Math.round(best)));
}




/* ===== THEMIS pure-dispatch isolation overrides ===== */
(function(){
  function __themisPureOrderRows(items){
    const rows = Array.isArray(items) ? items.slice() : [];
    return rows.sort((a, b) => {
      const ah = Number(a?.actual_hour ?? a?.plan_hour ?? 0);
      const bh = Number(b?.actual_hour ?? b?.plan_hour ?? 0);
      if (ah !== bh) return ah - bh;

      const aVehicle = Number(a?.vehicle_id || 0);
      const bVehicle = Number(b?.vehicle_id || 0);
      if (aVehicle !== bVehicle) return aVehicle - bVehicle;

      const aStop = Number(a?.stop_order || 0);
      const bStop = Number(b?.stop_order || 0);
      if (aStop > 0 || bStop > 0) {
        if (aStop !== bStop) return aStop - bStop;
      }

      const ad = Number(a?.distance_km ?? a?.casts?.distance_km ?? 0);
      const bd = Number(b?.distance_km ?? b?.casts?.distance_km ?? 0);
      if (ad !== bd) return ad - bd;

      return Number(a?.id || 0) - Number(b?.id || 0);
    });
  }

  if (typeof window !== 'undefined') {
    window.__THEMIS_PURE_DISPATCH_MODE__ = true;
  }

  __hasEnoughVehiclesForDisplayGroups = function(){ return false; };
  __buildAssignmentsPreserveDisplayGroups = function(){ return []; };
  __buildEmergencyAssignments = function(){ return []; };
  resolveCapacityOverflowLocally = function(assignments){
    return Array.isArray(assignments) ? assignments : [];
  };
  sortItemsByNearestRoute = function(items){
    return __themisPureOrderRows(items);
  };
})();
/* ===== /THEMIS pure-dispatch isolation overrides ===== */
