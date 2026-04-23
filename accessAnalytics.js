(function(){
  const TABLE_NAME = 'dropoff_access_logs';
  const VISITOR_KEY_STORAGE = 'dropoff_access_visitor_key_v1';
  const SESSION_KEY_STORAGE = 'dropoff_access_session_key_v1';
  const LAST_LOG_PREFIX = 'dropoff_access_last_logged_v1_';
  const THROTTLE_MS = 30 * 60 * 1000;
  const TOKYO_TZ = 'Asia/Tokyo';
  let cachedClient = null;
  let disabledWrites = false;
  let disabledReads = false;

  function safeLocalStorageGet(key) {
    try { return window.localStorage.getItem(key); } catch (_) { return null; }
  }

  function safeLocalStorageSet(key, value) {
    try { window.localStorage.setItem(key, value); } catch (_) {}
  }

  function generateKey(prefix) {
    try {
      if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return prefix + '_' + window.crypto.randomUUID();
      }
    } catch (_) {}
    return prefix + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 12);
  }

  function getVisitorKey() {
    let value = safeLocalStorageGet(VISITOR_KEY_STORAGE);
    if (!value) {
      value = generateKey('visitor');
      safeLocalStorageSet(VISITOR_KEY_STORAGE, value);
    }
    return value;
  }

  function getSessionKey() {
    let value = safeLocalStorageGet(SESSION_KEY_STORAGE);
    if (!value) {
      value = generateKey('session');
      safeLocalStorageSet(SESSION_KEY_STORAGE, value);
    }
    return value;
  }

  function getAnalyticsClient() {
    if (cachedClient) return cachedClient;
    try {
      if (window.supabaseClient) {
        cachedClient = window.supabaseClient;
        return cachedClient;
      }
      if (!window.supabase || !window.APP_CONFIG?.SUPABASE_URL || !window.APP_CONFIG?.SUPABASE_ANON_KEY) return null;
      cachedClient = window.supabase.createClient(window.APP_CONFIG.SUPABASE_URL, window.APP_CONFIG.SUPABASE_ANON_KEY);
      return cachedClient;
    } catch (_) {
      return null;
    }
  }

  function normalizePageType(pageType) {
    return String(pageType || '').trim().toLowerCase() === 'dashboard' ? 'dashboard' : 'home';
  }

  function getThrottleStorageKey(pageType) {
    return LAST_LOG_PREFIX + normalizePageType(pageType);
  }

  function shouldThrottle(pageType) {
    const raw = safeLocalStorageGet(getThrottleStorageKey(pageType));
    const last = Number(raw || 0);
    return Number.isFinite(last) && last > 0 && (Date.now() - last) < THROTTLE_MS;
  }

  function markLogged(pageType) {
    safeLocalStorageSet(getThrottleStorageKey(pageType), String(Date.now()));
  }

  function getTokyoDateString(date) {
    try {
      return new Intl.DateTimeFormat('en-CA', {
        timeZone: TOKYO_TZ,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      }).format(date || new Date());
    } catch (_) {
      const d = date || new Date();
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
  }

  function shiftTokyoDate(dateString, offsetDays) {
    const base = new Date(`${dateString}T00:00:00Z`);
    base.setUTCDate(base.getUTCDate() + Number(offsetDays || 0));
    return getTokyoDateString(base);
  }

  function getDateRange(days) {
    const safeDays = Math.max(1, Math.min(90, Number(days || 30)));
    const end = getTokyoDateString(new Date());
    const start = shiftTokyoDate(end, -(safeDays - 1));
    const labels = [];
    let cursor = start;
    while (cursor <= end) {
      labels.push(cursor);
      cursor = shiftTokyoDate(cursor, 1);
    }
    return { start, end, labels };
  }

  function summarizeAnalyticsRows(rows, days) {
    const { start, end, labels } = getDateRange(days);
    const buckets = { home: {}, dashboard: {} };
    labels.forEach(label => {
      buckets.home[label] = 0;
      buckets.dashboard[label] = 0;
    });

    (Array.isArray(rows) ? rows : []).forEach(row => {
      const pageType = normalizePageType(row?.page_type);
      const label = String(row?.access_date || '').slice(0, 10);
      if (!label || !(label in buckets[pageType])) return;
      buckets[pageType][label] += 1;
    });

    const homeSeries = labels.map(label => ({ date: label, count: buckets.home[label] || 0 }));
    const dashboardSeries = labels.map(label => ({ date: label, count: buckets.dashboard[label] || 0 }));
    const today = end;
    const last7Labels = labels.slice(-7);

    const totalFor = series => series.reduce((sum, row) => sum + Number(row?.count || 0), 0);
    const totalForDates = (series, allow) => series.reduce((sum, row) => sum + (allow.has(row.date) ? Number(row?.count || 0) : 0), 0);

    return {
      start,
      end,
      labels,
      daily: {
        home: homeSeries,
        dashboard: dashboardSeries
      },
      totals: {
        home_today: Number(homeSeries.find(row => row.date === today)?.count || 0),
        dashboard_today: Number(dashboardSeries.find(row => row.date === today)?.count || 0),
        home_7d: totalForDates(homeSeries, new Set(last7Labels)),
        dashboard_7d: totalForDates(dashboardSeries, new Set(last7Labels)),
        home_30d: totalFor(homeSeries),
        dashboard_30d: totalFor(dashboardSeries)
      }
    };
  }

  function extractReferrerHost() {
    try {
      return document.referrer ? new URL(document.referrer).host.slice(0, 120) : null;
    } catch (_) {
      return null;
    }
  }

  function isMissingTableError(error) {
    const message = String(error?.message || '');
    const code = String(error?.code || '').trim();
    return code === 'PGRST205' || code === '42P01' || /relation .+ does not exist/i.test(message) || /Could not find the table/i.test(message) || /schema cache/i.test(message);
  }

  function describeAccessAnalyticsError(error) {
    if (isMissingTableError(error)) return 'アクセス解析テーブルが未作成です。SQLを実行してください。';
    const message = String(error?.message || '').trim();
    if (/row-level security/i.test(message)) return 'アクセス解析の権限設定が未反映です。SQLのPolicyを確認してください。';
    return message || 'アクセス解析の読込に失敗しました。';
  }

  async function insertAccessLog(pageType, context = {}) {
    const safePageType = normalizePageType(pageType);
    if (disabledWrites) return { ok: false, skipped: true };
    if (context?.force !== true && shouldThrottle(safePageType)) return { ok: true, skipped: true };
    const client = getAnalyticsClient();
    if (!client) return { ok: false, skipped: true };

    const payload = {
      page_type: safePageType,
      visitor_key: getVisitorKey(),
      session_key: getSessionKey(),
      page_path: String(context?.pagePath || window.location.pathname || '').slice(0, 190) || null,
      referrer_host: extractReferrerHost(),
      user_agent: String(window.navigator?.userAgent || '').slice(0, 255) || null,
      user_id: context?.userId ? String(context.userId).slice(0, 80) : null,
      team_id: context?.teamId ? String(context.teamId).slice(0, 80) : null
    };

    try {
      const { error } = await client.from(TABLE_NAME).insert(payload);
      if (error) {
        if (isMissingTableError(error)) disabledWrites = true;
        return { ok: false, error };
      }
      markLogged(safePageType);
      return { ok: true, skipped: false };
    } catch (error) {
      return { ok: false, error };
    }
  }

  async function fetchAccessAnalytics(options = {}) {
    if (disabledReads) return { data: null, error: new Error('disabled') };
    const client = getAnalyticsClient();
    if (!client) return { data: null, error: new Error('client_unavailable') };
    const safeDays = Math.max(1, Math.min(90, Number(options?.days || 30)));
    const { start } = getDateRange(safeDays);

    try {
      const { data, error } = await client
        .from(TABLE_NAME)
        .select('page_type, access_date')
        .gte('access_date', start)
        .order('access_date', { ascending: true });
      if (error) {
        if (isMissingTableError(error)) disabledReads = true;
        return { data: null, error };
      }
      return { data: summarizeAnalyticsRows(Array.isArray(data) ? data : [], safeDays), error: null };
    } catch (error) {
      return { data: null, error };
    }
  }

  function buildChartSvg(points, options = {}) {
    const safePoints = Array.isArray(points) ? points : [];
    if (!safePoints.length) return '';

    const width = Number(options.width || 720);
    const height = Number(options.height || 250);
    const padX = 28;
    const padTop = 16;
    const padBottom = 34;
    const graphWidth = Math.max(1, width - padX * 2);
    const graphHeight = Math.max(1, height - padTop - padBottom);
    const maxValue = Math.max(1, ...safePoints.map(row => Number(row?.count || 0)));
    const stepX = safePoints.length > 1 ? graphWidth / (safePoints.length - 1) : 0;
    const color = String(options.color || '#4f85ff');
    const fill = String(options.fill || 'rgba(79,133,255,.16)');

    const coords = safePoints.map((row, index) => {
      const value = Number(row?.count || 0);
      const x = padX + stepX * index;
      const y = padTop + graphHeight - (value / maxValue) * graphHeight;
      return { x, y, value, date: row?.date || '' };
    });

    const linePath = coords.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(' ');
    const areaPath = `${linePath} L ${(padX + graphWidth).toFixed(2)} ${(padTop + graphHeight).toFixed(2)} L ${padX.toFixed(2)} ${(padTop + graphHeight).toFixed(2)} Z`;
    const gridLines = [0, .25, .5, .75, 1].map(ratio => {
      const y = padTop + graphHeight - graphHeight * ratio;
      return `<line x1="${padX}" y1="${y.toFixed(2)}" x2="${(padX + graphWidth).toFixed(2)}" y2="${y.toFixed(2)}" stroke="rgba(255,255,255,.08)" stroke-width="1" />`;
    }).join('');

    const dots = coords.map(point => `<circle cx="${point.x.toFixed(2)}" cy="${point.y.toFixed(2)}" r="3.5" fill="${color}" />`).join('');
    const labels = [coords[0], coords[Math.floor((coords.length - 1) / 2)], coords[coords.length - 1]]
      .filter(Boolean)
      .map((point, index, arr) => {
        const anchor = index === 0 ? 'start' : (index === arr.length - 1 ? 'end' : 'middle');
        const value = String(point.date || '').slice(5).replace('-', '/');
        return `<text x="${point.x.toFixed(2)}" y="${(height - 10).toFixed(2)}" fill="rgba(255,255,255,.62)" font-size="11" text-anchor="${anchor}">${value}</text>`;
      }).join('');

    return `
      <svg viewBox="0 0 ${width} ${height}" class="platform-admin-chart-svg" role="img" aria-label="アクセス推移グラフ">
        ${gridLines}
        <path d="${areaPath}" fill="${fill}" stroke="none"></path>
        <path d="${linePath}" fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
        ${dots}
        ${labels}
        <text x="${padX}" y="12" fill="rgba(255,255,255,.62)" font-size="11">max ${maxValue}</text>
      </svg>
    `;
  }

  function renderAccessLineChart(container, points, options = {}) {
    if (!container) return;
    const safePoints = Array.isArray(points) ? points : [];
    const total = safePoints.reduce((sum, row) => sum + Number(row?.count || 0), 0);
    const today = Number(safePoints[safePoints.length - 1]?.count || 0);

    if (!safePoints.length) {
      container.innerHTML = '<div class="platform-admin-chart-empty">まだアクセスがありません</div>';
      return;
    }

    const hasAny = safePoints.some(row => Number(row?.count || 0) > 0);
    if (!hasAny) {
      container.innerHTML = '<div class="platform-admin-chart-empty">まだアクセスがありません</div>';
      return;
    }

    container.innerHTML = buildChartSvg(safePoints, options) +
      `<div class="platform-admin-chart-caption"><span>30日合計 ${total}</span><span>今日 ${today}</span></div>`;
  }

  window.recordHomeAccess = function() {
    return insertAccessLog('home');
  };
  window.recordDashboardAccess = function(context) {
    return insertAccessLog('dashboard', context || {});
  };
  window.fetchAccessAnalytics = fetchAccessAnalytics;
  window.renderAccessLineChart = renderAccessLineChart;
  window.describeAccessAnalyticsError = describeAccessAnalyticsError;
})();
