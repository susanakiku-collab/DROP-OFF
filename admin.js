const supabaseClient = window.supabase.createClient(
  window.APP_CONFIG.SUPABASE_URL,
  window.APP_CONFIG.SUPABASE_ANON_KEY
);

const FUNCTION_NAME = "quick-handler";

const accessStatus = document.getElementById("accessStatus");
const adminArea = document.getElementById("adminArea");
const newEmail = document.getElementById("newEmail");
const newPassword = document.getElementById("newPassword");
const createUserBtn = document.getElementById("createUserBtn");
const reloadUsersBtn = document.getElementById("reloadUsersBtn");
const logoutBtn = document.getElementById("logoutBtn");
const result = document.getElementById("result");
const userList = document.getElementById("userList");

let currentUser = null;

function setStatus(message, type = "") {
  accessStatus.textContent = message;
  accessStatus.className = `status ${type}`.trim();
}

function setResult(message, type = "") {
  result.textContent = message;
  result.className = `status ${type}`.trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatDateTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  return `${y}/${m}/${d} ${hh}:${mm}`;
}

async function callAdminFunction(payload) {
  const { data: sessionData } = await supabaseClient.auth.getSession();
  const accessToken = sessionData?.session?.access_token;

  if (!accessToken) {
    throw new Error("ログインセッションが見つかりません。再ログインしてください。");
  }

  const response = await fetch(`${window.APP_CONFIG.SUPABASE_URL}/functions/v1/${FUNCTION_NAME}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${accessToken}`,
      "apikey": window.APP_CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify(payload)
  });

  const json = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(json?.error || `Function error (${response.status})`);
  }

  return json;
}

function renderUserList(users = []) {
  if (!userList) return;

  if (!users.length) {
    userList.innerHTML = `<div class="status">ユーザーがいません。</div>`;
    return;
  }

  userList.innerHTML = users.map(user => {
    const isSelf = currentUser && user.id === currentUser.id;
    const roleLabel = user.is_admin ? "管理者" : "一般";
    const roleClass = user.is_admin ? "admin" : "normal";
    const deleteDisabled = isSelf ? "disabled" : "";
    const deleteTitle = isSelf ? "自分自身は削除できません" : "このユーザーを削除";

    return `
      <div class="user-row">
        <div class="user-main">
          <div class="user-email">${escapeHtml(user.email || "-")}</div>
          <div class="user-meta">作成日: ${escapeHtml(formatDateTime(user.created_at))}</div>
        </div>
        <div class="badge ${roleClass}">${roleLabel}</div>
        <button
          class="danger delete-user-btn"
          data-user-id="${escapeHtml(user.id)}"
          data-email="${escapeHtml(user.email || "")}"
          ${deleteDisabled}
          title="${escapeHtml(deleteTitle)}"
        >削除</button>
      </div>
    `;
  }).join("");

  userList.querySelectorAll(".delete-user-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const userId = btn.dataset.userId || "";
      const email = btn.dataset.email || "";
      await deleteUser(userId, email);
    });
  });
}

async function loadUsers() {
  if (!userList) return;
  userList.innerHTML = `<div class="status">読み込み中...</div>`;

  try {
    const json = await callAdminFunction({ action: "list" });
    renderUserList(Array.isArray(json.users) ? json.users : []);
  } catch (err) {
    userList.innerHTML = `<div class="status err">${escapeHtml(err.message || String(err))}</div>`;
  }
}

async function ensureAdmin() {
  const { data, error } = await supabaseClient.auth.getUser();

  if (error || !data?.user) {
    setStatus("ログイン情報が確認できません。ログイン画面へ戻ります。", "err");
    setTimeout(() => {
      window.location.href = "login.html";
    }, 1200);
    return;
  }

  currentUser = data.user;
  const isAdmin = currentUser.user_metadata?.is_admin === true;

  if (!isAdmin) {
    setStatus("この画面は管理者専用です。ダッシュボードへ戻ります。", "err");
    setTimeout(() => {
      window.location.href = "dashboard.html";
    }, 1200);
    return;
  }

  setStatus(`管理者確認OK\nログイン中: ${currentUser.email}`, "ok");
  adminArea.classList.remove("hidden");
  await loadUsers();
}

async function createUser() {
  const email = newEmail.value.trim();
  const password = newPassword.value.trim();

  if (!email || !password) {
    setResult("メールアドレスとパスワードを入力してください。", "err");
    return;
  }

  if (password.length < 6) {
    setResult("パスワードは6文字以上にしてください。", "err");
    return;
  }

  createUserBtn.disabled = true;
  setResult("作成中...", "warn");

  try {
    const json = await callAdminFunction({
      action: "create",
      email,
      password
    });

    setResult(`作成成功\nログインID: ${json.email || email}`, "ok");
    newEmail.value = "";
    newPassword.value = "";
    await loadUsers();
  } catch (err) {
    setResult("作成失敗: " + err.message, "err");
  } finally {
    createUserBtn.disabled = false;
  }
}

async function deleteUser(userId, email) {
  const confirmed = window.confirm(
    `このログインIDを削除しますか？\n\n${email}\n\n削除後は元に戻せません。`
  );
  if (!confirmed) return;

  setResult("削除中...", "warn");

  try {
    const json = await callAdminFunction({
      action: "delete",
      user_id: userId,
      email
    });

    setResult(`削除成功\nログインID: ${json.email || email}`, "ok");
    await loadUsers();
  } catch (err) {
    setResult("削除失敗: " + err.message, "err");
  }
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.location.href = "index.html";
}

createUserBtn.addEventListener("click", createUser);
reloadUsersBtn.addEventListener("click", loadUsers);
logoutBtn.addEventListener("click", logout);
ensureAdmin();
