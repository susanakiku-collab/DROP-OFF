const PAYPAY_RETURN_STATE_STORAGE_KEY = 'dropoff_paypay_return_state';
const BILLING_RETURN_QUERY_KEY = 'billing';
const PAYPAY_CONFIRM_FUNCTION_NAME = String(window.APP_CONFIG?.PAYPAY_CONFIRM_FUNCTION_NAME || 'paypay-confirm-payment').trim();
const SUPABASE_URL = String(window.APP_CONFIG?.SUPABASE_URL || '').trim();
const SUPABASE_ANON_KEY = String(window.APP_CONFIG?.SUPABASE_ANON_KEY || '').trim();

const statusEl = document.getElementById('paypayReturnStatusText');
const subEl = document.getElementById('paypayReturnSubText');
const backBtn = document.getElementById('backToDashboardBtn');

const supabaseClient = window.supabase?.createClient
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;

function setStatus(message = '', detail = '', isError = false) {
  if (statusEl) {
    statusEl.textContent = String(message || '').trim();
    statusEl.classList.toggle('text-danger', Boolean(isError));
    statusEl.classList.toggle('text-done', !isError && Boolean(String(message || '').trim()));
  }
  if (subEl) subEl.textContent = String(detail || '').trim();
}

function getStoredReturnState() {
  try {
    return JSON.parse(window.sessionStorage.getItem(PAYPAY_RETURN_STATE_STORAGE_KEY) || '{}') || {};
  } catch (_) {
    return {};
  }
}

function clearStoredReturnState() {
  try {
    window.sessionStorage.removeItem(PAYPAY_RETURN_STATE_STORAGE_KEY);
  } catch (_) {}
}

function getDashboardUrl(state = '', teamId = '') {
  const url = new URL('dashboard.html', window.location.href);
  if (state) url.searchParams.set(BILLING_RETURN_QUERY_KEY, state);
  if (teamId) url.searchParams.set('team_id', teamId);
  return url.toString();
}

async function confirmPayPayPayment(merchantPaymentId, teamId) {
  if (!supabaseClient) throw new Error('Supabase client を初期化できませんでした。');
  const sessionResult = await supabaseClient.auth.getSession();
  const accessToken = String(sessionResult?.data?.session?.access_token || '').trim();
  if (!accessToken) throw new Error('ログイン状態を取得できませんでした。再ログインしてください。');

  const response = await fetch(`${SUPABASE_URL}/functions/v1/${PAYPAY_CONFIRM_FUNCTION_NAME}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${accessToken}`,
      apikey: SUPABASE_ANON_KEY
    },
    body: JSON.stringify({
      merchantPaymentId,
      teamId
    })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(String(payload?.error || `PayPay支払い確認に失敗しました (${response.status})`));
  }
  return payload || {};
}

async function runPayPayReturnCheck() {
  const params = new URLSearchParams(window.location.search || '');
  const stored = getStoredReturnState();
  const merchantPaymentId = String(
    params.get('merchant_payment_id') ||
    params.get('merchantPaymentId') ||
    stored?.merchantPaymentId ||
    ''
  ).trim();
  const teamId = String(params.get('team_id') || stored?.teamId || '').trim();

  if (backBtn) {
    backBtn.addEventListener('click', () => {
      window.location.href = getDashboardUrl('', teamId);
    });
  }

  if (!merchantPaymentId) {
    setStatus('支払い情報を取得できませんでした。', 'PayPay決済開始直後に画面を閉じた場合は、設定から再読込してください。', true);
    return;
  }

  try {
    setStatus('支払い確認中です...', '数秒かかる場合があります。');
    let payload = null;
    for (let attempt = 0; attempt < 4; attempt += 1) {
      payload = await confirmPayPayPayment(merchantPaymentId, teamId);
      const status = String(payload?.status || '').trim().toLowerCase();
      if (status === 'completed') {
        clearStoredReturnState();
        setStatus('支払いが完了しました。', '設定画面へ戻ります。', false);
        window.setTimeout(() => {
          window.location.href = getDashboardUrl('paypay_success', teamId);
        }, 900);
        return;
      }
      if (status === 'failed' || status === 'cancelled' || status === 'expired') {
        clearStoredReturnState();
        setStatus('支払いを確認できませんでした。', '支払い状態を確認してから、必要なら再度お試しください。', true);
        window.setTimeout(() => {
          window.location.href = getDashboardUrl('paypay_failed', teamId);
        }, 1200);
        return;
      }
      if (attempt < 3) {
        await new Promise(resolve => window.setTimeout(resolve, 3000));
      }
    }
    setStatus('支払い確認を継続しています。', 'まだ反映前の可能性があります。設定へ戻ってプランを再読込してください。', false);
  } catch (error) {
    console.error('runPayPayReturnCheck error:', error);
    setStatus(error?.message || '支払い確認に失敗しました。', '設定画面へ戻って、必要ならプランを再読込してください。', true);
  }
}

runPayPayReturnCheck();
