(function () {
  const form = document.getElementById('signupForm');
  const teamNameInput = document.getElementById('signupTeamName');
  const displayNameInput = document.getElementById('signupDisplayName');
  const emailInput = document.getElementById('signupEmail');
  const passwordInput = document.getElementById('signupPassword');
  const submitBtn = document.getElementById('signupSubmitBtn');
  const messageEl = document.getElementById('signupMessage');

  const supabaseUrl = window.APP_CONFIG?.SUPABASE_URL;
  const supabaseAnonKey = window.APP_CONFIG?.SUPABASE_ANON_KEY;
  const supabaseClient = window.supabaseClient
    || (window.supabase?.createClient && supabaseUrl && supabaseAnonKey
      ? window.supabase.createClient(supabaseUrl, supabaseAnonKey)
      : null);

  if (supabaseClient && !window.supabaseClient) {
    window.supabaseClient = supabaseClient;
  }

  const SIGNUP_TEAM_STORAGE_KEY = '__DROP_OFF_SIGNUP_TEAM_ID__';
  const SIGNUP_USER_ID_KEY = '__DROP_OFF_SIGNUP_USER_ID__';
  const SIGNUP_USER_EMAIL_KEY = '__DROP_OFF_SIGNUP_USER_EMAIL__';

  function getWorkspaceCacheKey(userId) {
    return userId ? `dropoff_workspace_team_id_v1_${userId}` : 'dropoff_workspace_team_id_v1';
  }

  function persistPreferredWorkspaceTeam(teamId, user) {
    const normalized = String(teamId || '').trim();
    if (!normalized) return;
    const userId = String(user?.id || '').trim();
    const userEmail = String(user?.email || emailInput?.value || '').trim().toLowerCase();
    try { window.localStorage.setItem(SIGNUP_TEAM_STORAGE_KEY, normalized); } catch (_) {}
    try { if (userId) window.localStorage.setItem(SIGNUP_USER_ID_KEY, userId); } catch (_) {}
    try { if (userEmail) window.localStorage.setItem(SIGNUP_USER_EMAIL_KEY, userEmail); } catch (_) {}
    try { window.localStorage.setItem('dropoff_workspace_team_id', normalized); } catch (_) {}
    try { window.localStorage.setItem('workspaceTeamId', normalized); } catch (_) {}
    try { window.localStorage.setItem('current_dropoff_team_id', normalized); } catch (_) {}
    try { window.localStorage.setItem('__DROP_OFF_LAST_TEAM_ID__', normalized); } catch (_) {}
    try { if (userId) window.localStorage.setItem(getWorkspaceCacheKey(userId), normalized); } catch (_) {}
    try { window.__DROP_OFF_CURRENT_TEAM_ID__ = normalized; } catch (_) {}
  }

  function setMessage(message, type) {
    if (!messageEl) return;
    messageEl.textContent = message || '';
    messageEl.classList.remove('error', 'success');
    if (type) messageEl.classList.add(type);
  }

  function setLoading(loading) {
    if (!submitBtn) return;
    submitBtn.disabled = !!loading;
    submitBtn.textContent = loading ? '登録中...' : '新規登録して開始';
  }

  function normalizeErrorMessage(error) {
    const code = String(error?.code || '').toLowerCase();
    const raw = String(error?.message || '');
    const lower = raw.toLowerCase();

    if (code === 'user_already_exists' || lower.includes('already registered') || lower.includes('user already registered')) {
      return 'このメールアドレスはすでに登録されています。ログイン画面からお試しください。';
    }
    if (lower.includes('password')) {
      return 'パスワード条件を満たしていません。6文字以上で入力してください。';
    }
    if (lower.includes('not authenticated')) {
      return '登録後の認証情報を取得できませんでした。画面を再読込してもう一度お試しください。';
    }
    return raw || '登録に失敗しました。';
  }

  async function signUpDropoffUserDirect(email, password) {
    if (!supabaseClient?.auth?.signUp) {
      throw new Error('Supabase 初期化に失敗しました。config.js と CDN読込を確認してください。');
    }
    return await supabaseClient.auth.signUp({ email, password });
  }

  async function createWorkspaceForSignupDirect(teamName, displayName) {
    if (!supabaseClient?.rpc) {
      throw new Error('Supabase RPC が使えません。');
    }
    const { data, error } = await supabaseClient.rpc('create_dropoff_workspace_for_signup', {
      p_team_name: teamName,
      p_display_name: displayName || null,
    });
    return { teamId: data, error };
  }

  async function ensureSession() {
    if (!supabaseClient?.auth?.getSession) return null;
    const { data, error } = await supabaseClient.auth.getSession();
    if (error) return null;
    return data?.session || null;
  }

  async function handleSubmit(event) {
    event.preventDefault();
    setMessage('', null);

    const teamName = String(teamNameInput?.value || '').trim();
    const displayName = String(displayNameInput?.value || '').trim();
    const email = String(emailInput?.value || '').trim().toLowerCase();
    const password = String(passwordInput?.value || '');

    if (!teamName) {
      setMessage('チーム名 / 会社名を入力してください。', 'error');
      teamNameInput?.focus();
      return;
    }
    if (!displayName) {
      setMessage('表示名を入力してください。', 'error');
      displayNameInput?.focus();
      return;
    }
    if (!email) {
      setMessage('メールアドレスを入力してください。', 'error');
      emailInput?.focus();
      return;
    }
    if (password.length < 6) {
      setMessage('パスワードは6文字以上で入力してください。', 'error');
      passwordInput?.focus();
      return;
    }

    setLoading(true);
    setMessage('アカウントを作成しています...', 'success');

    try {
      const signUpResult = await signUpDropoffUserDirect(email, password);
      if (signUpResult?.error) throw signUpResult.error;

      const signedUpUser = signUpResult?.data?.user || null;
      let session = signUpResult?.data?.session || null;
      if (!session) {
        session = await ensureSession();
      }

      if (!session) {
        setMessage('登録は受け付けました。認証メールが届く設定の場合は、有効化後にログインしてください。', 'success');
        return;
      }

      setMessage('ワークスペースを作成しています...', 'success');
      const workspaceResult = await createWorkspaceForSignupDirect(teamName, displayName);
      if (workspaceResult?.error) throw workspaceResult.error;
      if (!workspaceResult?.teamId) throw new Error('team_id を取得できませんでした。');

      persistPreferredWorkspaceTeam(workspaceResult.teamId, session?.user || signedUpUser || null);

      setMessage('登録完了。初期設定へ移動します。', 'success');
      window.location.href = `dashboard.html?onboarding=1&team_id=${encodeURIComponent(workspaceResult.teamId)}`;
    } catch (error) {
      console.error('signup failed:', error);
      setMessage(normalizeErrorMessage(error), 'error');
    } finally {
      setLoading(false);
    }
  }

  form?.addEventListener('submit', handleSubmit);
})();
