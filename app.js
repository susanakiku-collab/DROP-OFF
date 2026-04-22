const supabaseClient = window.supabase.createClient(
  window.APP_CONFIG.SUPABASE_URL,
  window.APP_CONFIG.SUPABASE_ANON_KEY
);

const loginPanel = document.getElementById("loginPanel");
const resetPanel = document.getElementById("resetPanel");
const authTitle = document.getElementById("authTitle");
const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginBtn = document.getElementById("loginBtn");
const openSignupBtn = document.getElementById("openSignupBtn");
const signupPanel = document.getElementById("signupPanel");
const signupEmail = document.getElementById("signupEmail");
const signupPassword = document.getElementById("signupPassword");
const signupPasswordConfirm = document.getElementById("signupPasswordConfirm");
const signupBtn = document.getElementById("signupBtn");
const backToLoginFromSignupBtn = document.getElementById("backToLoginFromSignupBtn");
const newPassword = document.getElementById("newPassword");
const confirmPassword = document.getElementById("confirmPassword");
const resetPasswordBtn = document.getElementById("resetPasswordBtn");
const backToLoginBtn = document.getElementById("backToLoginBtn");
const authMessage = document.getElementById("authMessage");

const RECOVERY_FLAG_KEY = "dropoff_password_recovery_pending";
const PENDING_SIGNUP_KEY = "__DROP_OFF_PENDING_SIGNUP__";

let recoveryMode = false;

function persistWorkspaceTeamId(teamId, userId = "") {
  const value = String(teamId || "").trim();
  const safeUserId = String(userId || "").trim();
  if (!value) return;
  try { window.localStorage.setItem("dropoff_workspace_team_id", value); } catch (e) {}
  try { window.localStorage.setItem("current_dropoff_team_id", value); } catch (e) {}
  try { window.localStorage.setItem("workspaceTeamId", value); } catch (e) {}
  try { if (safeUserId) window.localStorage.setItem(`dropoff_workspace_team_id_v1_${safeUserId}`, value); } catch (e) {}
}

function readPendingSignup() {
  try {
    const raw = window.localStorage.getItem(PENDING_SIGNUP_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function clearPendingSignup() {
  try { window.localStorage.removeItem(PENDING_SIGNUP_KEY); } catch (e) {}
}

function getPendingSignupForUser(user) {
  const email = String(user?.email || "").trim().toLowerCase();
  const pending = readPendingSignup();
  if (pending && String(pending.email || "").trim().toLowerCase() === email) {
    return pending;
  }

  const metaTeam = String(user?.user_metadata?.team_name || "").trim();
  const metaDisplay = String(user?.user_metadata?.display_name || "").trim();
  if (email && metaTeam && metaDisplay) {
    return { email, teamName: metaTeam, displayName: metaDisplay };
  }
  return null;
}

async function userAlreadyHasWorkspace(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return false;
  try {
    const { data, error } = await supabaseClient
      .from("dropoff_team_members")
      .select("team_id")
      .eq("user_id", safeUserId)
      .limit(1)
      .maybeSingle();

    if (error) {
      console.warn("team membership check failed:", error);
      return false;
    }

    const teamId = String(data?.team_id || "").trim();
    if (teamId) {
      persistWorkspaceTeamId(teamId, safeUserId);
      clearPendingSignup();
      return true;
    }
  } catch (error) {
    console.warn("team membership check exception:", error);
  }
  return false;
}

async function createWorkspaceAfterConfirmedLogin(user) {
  const safeUserId = String(user?.id || "").trim();
  if (!safeUserId) return null;

  if (await userAlreadyHasWorkspace(safeUserId)) {
    return { teamId: window.localStorage.getItem("dropoff_workspace_team_id") || "" };
  }

  const pending = getPendingSignupForUser(user);
  const teamName = String(pending?.teamName || "").trim();
  const displayName = String(pending?.displayName || "").trim();
  if (!teamName || !displayName) return null;

  const { data, error } = await supabaseClient.rpc("create_dropoff_workspace_for_signup", {
    p_team_name: teamName,
    p_display_name: displayName || null,
  });

  if (error) throw error;

  const teamId = String(data || "").trim();
  if (!teamId) throw new Error("初回ワークスペースの作成に失敗しました。");

  persistWorkspaceTeamId(teamId, safeUserId);
  clearPendingSignup();
  return { teamId, onboarding: true };
}

function buildDashboardUrl() {
  try {
    const teamId = window.localStorage.getItem('dropoff_workspace_team_id') || window.localStorage.getItem('current_dropoff_team_id') || '';
    const safeTeamId = String(teamId || '').trim();
    return safeTeamId ? `dashboard.html?team_id=${encodeURIComponent(safeTeamId)}` : 'dashboard.html';
  } catch (e) {
    return 'dashboard.html';
  }
}

function setMessage(message, isError = true) {
  if (!authMessage) return;
  authMessage.textContent = message;
  authMessage.style.color = isError ? "#ff6b6b" : "#7be2ab";
}

function markRecoveryPending() {
  try {
    sessionStorage.setItem(RECOVERY_FLAG_KEY, "1");
  } catch (e) {}
}

function clearRecoveryPending() {
  try {
    sessionStorage.removeItem(RECOVERY_FLAG_KEY);
  } catch (e) {}
}

function hasRecoveryPending() {
  try {
    return sessionStorage.getItem(RECOVERY_FLAG_KEY) === "1";
  } catch (e) {
    return false;
  }
}

function showLoginMode() {
  recoveryMode = false;
  if (authTitle) authTitle.textContent = "ログイン";
  loginPanel?.classList.remove("hidden");
  signupPanel?.classList.add("hidden");
  resetPanel?.classList.add("hidden");
  if (loginPassword) loginPassword.value = "";
}

function showRecoveryMode() {
  recoveryMode = true;
  markRecoveryPending();
  if (authTitle) authTitle.textContent = "パスワード再設定";
  loginPanel?.classList.add("hidden");
  signupPanel?.classList.add("hidden");
  resetPanel?.classList.remove("hidden");
  setMessage("新しいパスワードを入力してください。", false);
}

function showSignupMode() {
  recoveryMode = false;
  if (authTitle) authTitle.textContent = "初回参加";
  loginPanel?.classList.add("hidden");
  signupPanel?.classList.remove("hidden");
  resetPanel?.classList.add("hidden");
  if (signupPassword) signupPassword.value = "";
  if (signupPasswordConfirm) signupPasswordConfirm.value = "";
  setMessage("招待されたメールアドレスで初回参加を登録してください。", false);
}

function isRecoveryUrl() {
  const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ""));
  const queryParams = new URLSearchParams(window.location.search);
  const type = hashParams.get("type") || queryParams.get("type") || "";
  const code = hashParams.get("code") || queryParams.get("code") || "";
  const tokenHash = hashParams.get("token_hash") || queryParams.get("token_hash") || "";
  const accessToken = hashParams.get("access_token") || queryParams.get("access_token") || "";
  const refreshToken = hashParams.get("refresh_token") || queryParams.get("refresh_token") || "";
  return type === "recovery" || Boolean(code || tokenHash || accessToken || refreshToken);
}

function isAlreadyRegisteredError(error) {
  const message = String(error?.message || "").toLowerCase();
  const code = String(error?.code || "").toLowerCase();
  return (
    code === "user_already_exists" ||
    message.includes("already registered") ||
    message.includes("already been registered") ||
    message.includes("user already registered")
  );
}

function isEmailNotConfirmedError(error) {
  const message = String(error?.message || "").toLowerCase();
  const code = String(error?.code || "").toLowerCase();
  return code === "email_not_confirmed" || message.includes("email not confirmed");
}

function isMissingInvitationCheckFunctionError(error) {
  const code = String(error?.code || "").toUpperCase();
  const message = String(error?.message || "");
  return code === "PGRST202" || /dropoff_check_invitation/i.test(message);
}

function getInviteEmailRedirectUrl() {
  try {
    return new URL("login.html", window.location.href).toString();
  } catch (e) {
    return window.location.href;
  }
}

async function checkInvitationBeforeSignup(email) {
  const normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) {
    return { allowed: false, reason: "メールアドレスを入力してください。" };
  }

  try {
    const { data, error } = await supabaseClient.rpc("dropoff_check_invitation", {
      p_email: normalizedEmail
    });

    if (error) {
      if (isMissingInvitationCheckFunctionError(error)) {
        throw new Error("招待事前確認SQLが未適用です。付属SQLをSupabaseで実行してください。");
      }
      throw error;
    }

    const result = data && typeof data === "object" ? data : {};
    return {
      allowed: result.allowed === true,
      status: String(result.status || "").trim().toLowerCase(),
      reason: String(result.reason || "").trim(),
      invitation_id: result.invitation_id || null,
      team_id: result.team_id || null,
      invited_role: result.invited_role || "user",
      display_name: result.display_name || ""
    };
  } catch (error) {
    console.error("dropoff_check_invitation failed:", error);
    throw error;
  }
}

async function completeInvitedAccessAndRedirect() {
  const { data: userData, error: userError } = await supabaseClient.auth.getUser();
  if (userError) throw userError;
  const user = userData?.user || null;
  if (!user) {
    throw new Error("登録後のユーザー情報を取得できませんでした。");
  }

  if (typeof window.ensureInvitedAccessForUser === "function") {
    const accessInfo = await window.ensureInvitedAccessForUser(user);
    if (accessInfo?.allowed === false) {
      await supabaseClient.auth.signOut();
      throw new Error(accessInfo.reason || "このアカウントでは参加できません。招待を確認してください。");
    }
  }

  setMessage("参加登録が完了しました。ダッシュボードへ移動します。", false);
  window.location.href = "dashboard.html";
}

async function trySignInAfterSignup(email, password) {
  const { data, error } = await supabaseClient.auth.signInWithPassword({
    email,
    password
  });

  if (error) {
    if (isEmailNotConfirmedError(error)) {
      signupPassword.value = "";
      signupPasswordConfirm.value = "";
      setMessage("登録は受け付けました。認証メールを開いて有効化したあと、ログインしてください。", false);
      return false;
    }
    throw error;
  }

  if (!data?.session) {
    signupPassword.value = "";
    signupPasswordConfirm.value = "";
    setMessage("登録は完了しましたが、ログインセッションを開始できませんでした。ログイン画面から入り直してください。", false);
    showLoginMode();
    if (loginEmail) loginEmail.value = email;
    return false;
  }

  await completeInvitedAccessAndRedirect();
  return true;
}

async function handleLogin() {
  try {
    setMessage("ログイン中...", false);

    const email = loginEmail?.value.trim() || "";
    const password = loginPassword?.value.trim() || "";

    if (!email || !password) {
      setMessage("ログインIDとパスワードを入力してください。");
      return;
    }

    const { data, error } = await supabaseClient.auth.signInWithPassword({
      email,
      password
    });

    if (error) {
      setMessage("ログイン失敗: " + error.message);
      return;
    }

    clearRecoveryPending();

    let nextUrl = buildDashboardUrl();
    try {
      const user = data?.user || null;
      const created = await createWorkspaceAfterConfirmedLogin(user);
      if (created?.teamId) {
        nextUrl = created?.onboarding
          ? `dashboard.html?onboarding=1&team_id=${encodeURIComponent(created.teamId)}`
          : `dashboard.html?team_id=${encodeURIComponent(created.teamId)}`;
      }
    } catch (workspaceError) {
      console.error("workspace create after login failed:", workspaceError);
      setMessage("ログイン後の初期設定に失敗しました。もう一度ログインしてください。", true);
      return;
    }

    setMessage("ログイン成功", false);
    console.log("login success:", data);
    window.location.href = nextUrl;
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

async function handleInviteSignup() {
  try {
    const email = signupEmail?.value.trim().toLowerCase() || "";
    const password = signupPassword?.value || "";
    const confirm = signupPasswordConfirm?.value || "";

    if (!email || !password || !confirm) {
      setMessage("メールアドレスとパスワードを入力してください。");
      return;
    }

    if (password.length < 6) {
      setMessage("パスワードは6文字以上で入力してください。");
      return;
    }

    if (password !== confirm) {
      setMessage("確認用パスワードが一致しません。");
      return;
    }

    setMessage("招待情報を確認中...", false);
    const invitationCheck = await checkInvitationBeforeSignup(email);

    if (!invitationCheck?.allowed) {
      signupPassword.value = "";
      signupPasswordConfirm.value = "";
      setMessage(invitationCheck?.reason || "このメールアドレスでは初回参加できません。招待をご確認ください。");
      return;
    }

    setMessage("初回参加を登録中...", false);

    const { data, error } = await supabaseClient.auth.signUp({
      email,
      password,
      options: {
        emailRedirectTo: getInviteEmailRedirectUrl()
      }
    });

    if (error) {
      if (isAlreadyRegisteredError(error)) {
        signupPassword.value = "";
        signupPasswordConfirm.value = "";
        showLoginMode();
        if (loginEmail) loginEmail.value = email;
        setMessage("このメールアドレスはすでに登録済みです。ログイン画面から入ってください。", false);
        return;
      }
      setMessage("初回参加の登録に失敗しました: " + error.message);
      return;
    }

    if (data?.session) {
      await completeInvitedAccessAndRedirect();
      return;
    }

    const signedIn = await trySignInAfterSignup(email, password);
    if (signedIn) return;

    showLoginMode();
    if (loginEmail) loginEmail.value = email;
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

async function handlePasswordReset() {
  try {
    const password = newPassword?.value || "";
    const confirm = confirmPassword?.value || "";

    if (!password || !confirm) {
      setMessage("新しいパスワードを入力してください。");
      return;
    }

    if (password.length < 6) {
      setMessage("パスワードは6文字以上で入力してください。");
      return;
    }

    if (password !== confirm) {
      setMessage("確認用パスワードが一致しません。");
      return;
    }

    setMessage("パスワードを更新中...", false);

    const { data: sessionData } = await supabaseClient.auth.getSession();
    if (!sessionData?.session) {
      setMessage("再設定セッションが見つかりません。メールの最新リンクからもう一度開いてください。");
      return;
    }

    const { error } = await supabaseClient.auth.updateUser({ password });
    if (error) {
      setMessage("パスワード更新失敗: " + error.message);
      return;
    }

    const { data: userData } = await supabaseClient.auth.getUser();
    const userEmail = userData?.user?.email || "";

    setMessage("パスワードを更新しました。新しいパスワードでログインしてください。", false);

    await supabaseClient.auth.signOut();
    clearRecoveryPending();

    if (loginEmail && userEmail) loginEmail.value = userEmail;
    if (newPassword) newPassword.value = "";
    if (confirmPassword) confirmPassword.value = "";

    window.history.replaceState({}, document.title, window.location.pathname);
    showLoginMode();
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

async function initAuthScreen() {
  try {
    supabaseClient.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        markRecoveryPending();
        showRecoveryMode();
      }
    });

    const { data } = await supabaseClient.auth.getSession();
    const hasSession = Boolean(data?.session);
    const recoveryRequested = isRecoveryUrl() || hasRecoveryPending();

    if (recoveryRequested) {
      if (hasSession) {
        showRecoveryMode();
        return;
      }

      setMessage("再設定リンクを確認しています...", false);
      setTimeout(async () => {
        const { data: delayed } = await supabaseClient.auth.getSession();
        if (delayed?.session) {
          showRecoveryMode();
        } else if (hasRecoveryPending()) {
          showRecoveryMode();
        } else {
          setMessage("再設定リンクの確認に失敗しました。最新のメールから開き直してください。");
        }
      }, 900);
      return;
    }

    if (hasSession) {
      window.location.href = buildDashboardUrl();
      return;
    }

    clearRecoveryPending();
    showLoginMode();
  } catch (err) {
    console.error(err);
    showLoginMode();
  }
}

if (loginBtn) loginBtn.addEventListener("click", handleLogin);
if (openSignupBtn) openSignupBtn.addEventListener("click", showSignupMode);
if (signupBtn) signupBtn.addEventListener("click", handleInviteSignup);
if (resetPasswordBtn) resetPasswordBtn.addEventListener("click", handlePasswordReset);
if (backToLoginBtn) {
  backToLoginBtn.addEventListener("click", async () => {
    try {
      await supabaseClient.auth.signOut();
    } catch (e) {}
    clearRecoveryPending();
    window.history.replaceState({}, document.title, window.location.pathname);
    setMessage("");
    showLoginMode();
  });
}
if (backToLoginFromSignupBtn) {
  backToLoginFromSignupBtn.addEventListener("click", () => {
    setMessage("");
    showLoginMode();
  });
}

loginPassword?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") handleLogin();
});
signupPasswordConfirm?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") handleInviteSignup();
});
signupPassword?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") handleInviteSignup();
});
confirmPassword?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") handlePasswordReset();
});
newPassword?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") handlePasswordReset();
});

initAuthScreen();
