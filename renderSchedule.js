// renderSchedule
// Fully restored wrapper/public export version for schedule rendering

(function () {
  function getGlobal(name, fallback) {
    try {
      if (typeof window !== 'undefined' && typeof window[name] !== 'undefined') return window[name];
    } catch (_) {}
    return fallback;
  }

  function setScheduleRendererDeps(deps) {
    try {
      window.__scheduleRendererDeps = deps || {};
    } catch (_) {}
  }

  function resolveScheduleRendererContext() {
    const injected = getGlobal('__scheduleRendererDeps', {}) || {};
    return {
      els: injected.els || getGlobal('els', {}),
      plans: Array.isArray(injected.plans) ? injected.plans : (getGlobal('currentPlansCache', []) || []),
      actuals: Array.isArray(injected.actuals) ? injected.actuals : (getGlobal('currentActualsCache', []) || []),
      helpers: injected.helpers || {},
      actions: injected.actions || {}
    };
  }

  function helper(name, fallback) {
    const ctx = resolveScheduleRendererContext();
    if (ctx.helpers && typeof ctx.helpers[name] === 'function') return ctx.helpers[name];
    const globalFn = getGlobal(name, null);
    if (typeof globalFn === 'function') return globalFn;
    return fallback;
  }

  function getMatrixLegendHtml() {
    return `
      <div class="matrix-legend" aria-label="色の説明">
        <span class="matrix-legend-item"><span class="matrix-dot pending"></span>黄色 = 未完了</span>
        <span class="matrix-legend-item"><span class="matrix-dot done"></span>緑 = 完了</span>
        <span class="matrix-legend-item"><span class="matrix-dot cancel"></span>赤 = キャンセル</span>
      </div>
    `;
  }

  function getPlanLinkedActualStatusCore(planRow, { actuals, els, helpers } = {}) {
    const normalizeStatus = (helpers && helpers.normalizeStatus) || helper('normalizeStatus', s => s || 'pending');
    const todayStr = (helpers && helpers.todayStr) || helper('todayStr', () => new Date().toISOString().slice(0, 10));
    const safeActuals = Array.isArray(actuals) ? actuals : [];
    const linked = safeActuals.find(
      item =>
        Number(item && item.cast_id) === Number(planRow && planRow.cast_id) &&
        String((item && item.actual_date) || (els && els.actualDate && els.actualDate.value) || todayStr()) ===
          String((planRow && planRow.plan_date) || (els && els.planDate && els.planDate.value) || todayStr()) &&
        Number((item && (item.actual_hour ?? item.plan_hour)) ?? -1) === Number((planRow && (planRow.plan_hour ?? planRow.actual_hour)) ?? -1)
    );
    if (linked) return normalizeStatus(linked.status);
    return normalizeStatus(planRow && planRow.status);
  }

  function formatMinutesForScheduleDisplay(totalMinutes) {
    const safe = Math.max(0, Math.round(Number(totalMinutes || 0)));
    const hours = Math.floor(safe / 60);
    const minutes = safe % 60;
    if (!safe) return "";
    if (hours <= 0) return `${minutes}分`;
    if (minutes === 0) return `${hours}時間`;
    return `${hours}時間${minutes}分`;
  }

  function buildMatrixMetaText(distanceKm, travelMinutes) {
    const parts = [];
    const distance = Number(distanceKm || 0);
    const minutes = Number(travelMinutes || 0);
    if (Number.isFinite(distance) && distance > 0) parts.push(`${distance.toFixed(1)}km`);
    if (Number.isFinite(minutes) && minutes > 0) parts.push(`片道${formatMinutesForScheduleDisplay(minutes)}`);
    return parts.length ? ` (${parts.join(' / ')})` : '';
  }

  function buildMatrixNameLine(row, status, addressKey = 'destination_address') {
    const normalizeStatus = helper('normalizeStatus', s => s || 'pending');
    const buildMapLinkHtml = helper('buildMapLinkHtml', ({ name }) => String(name || '-'));
    const escapeHtml = helper('escapeHtml', v => String(v ?? ''));
    const safeStatus = normalizeStatus(status);
    const displayName = (row && row.person_name) || (row && row.casts && row.casts.name) || '-';
    const address = row && (row[addressKey] || row.destination_address || (row.casts && row.casts.address));
    const lat = row && (row.destination_lat ?? (row.casts && (row.casts.latitude ?? row.casts.lat)));
    const lng = row && (row.destination_lng ?? (row.casts && (row.casts.longitude ?? row.casts.lng)));
    const nameHtml = buildMapLinkHtml({
      name: displayName,
      address,
      lat,
      lng,
      className: `map-name-link matrix-name status-${safeStatus}`
    });
    const metaText = buildMatrixMetaText(row && row.distance_km, (row && row.travel_minutes) || (row && row.casts && row.casts.travel_minutes));
    const overflowText = Number(row && row.vehicle_id || 0) <= 0 && String(row && row.driver_name || '').includes('あぶれ')
      ? `<span class="matrix-meta" style="margin-left:6px;"> あぶれ</span>`
      : '';
    return `<span class="matrix-line">${nameHtml}<span class="matrix-meta">${escapeHtml(metaText)}</span>${overflowText}</span>`;
  }

  function renderPlanGroupedTableCore({ els, plans, actuals, actions, helpers }) {
    if (!els || !els.plansGroupedTable) return;

    const safePlans = Array.isArray(plans) ? plans : [];
    const getHourLabel = (helpers && helpers.getHourLabel) || helper('getHourLabel', h => `${Number(h)}時`);
    const getGroupedAreasByDisplay = (helpers && helpers.getGroupedAreasByDisplay) || helper('getGroupedAreasByDisplay', (items, pick) => {
      const seen = new Set();
      return (items || []).map(x => {
        const detailArea = helper('normalizeAreaLabel', v => String(v || '無し'))(pick(x));
        if (seen.has(detailArea)) return null;
        seen.add(detailArea);
        return { detailArea };
      }).filter(Boolean);
    });
    const normalizeAreaLabel = (helpers && helpers.normalizeAreaLabel) || helper('normalizeAreaLabel', v => String(v || '無し'));
    const resolvePreferredAreaLabel = (helpers && helpers.resolvePreferredAreaLabel) || helper('resolvePreferredAreaLabel', (area, address, lat, lng) => normalizeAreaLabel(area || (typeof guessArea === 'function' ? guessArea(lat, lng, address) : '無し') || '無し'));
    const getGroupedAreaHeaderHtml = (helpers && helpers.getGroupedAreaHeaderHtml) || helper('getGroupedAreaHeaderHtml', area => String(area || ''));
    const buildMapLinkHtml = (helpers && helpers.buildMapLinkHtml) || helper('buildMapLinkHtml', ({ name }) => String(name || '-'));
    const escapeHtml = (helpers && helpers.escapeHtml) || helper('escapeHtml', v => String(v ?? ''));
    const getStatusText = (helpers && helpers.getStatusText) || helper('getStatusText', s => String(s || ''));

    if (!safePlans.length) {
      els.plansGroupedTable.innerHTML = `<div class="muted" style="padding:14px;">予定がありません</div>`;
      return;
    }

    const hours = [...new Set(safePlans.map(x => Number(x.plan_hour)))].sort((a, b) => a - b);
    let html = `<div class="grouped-plan-list">`;

    hours.forEach(hour => {
      const hourItems = safePlans.filter(x => Number(x.plan_hour) === hour);
      const groupedAreas = getGroupedAreasByDisplay(hourItems, x => resolvePreferredAreaLabel(
        x?.planned_area || x?.destination_area || x?.casts?.area || '',
        x?.destination_address || x?.casts?.address || '',
        x?.destination_lat ?? x?.casts?.latitude ?? x?.casts?.lat,
        x?.destination_lng ?? x?.casts?.longitude ?? x?.casts?.lng
      ));

      html += `<div class="grouped-section">`;
      html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

      groupedAreas.forEach(({ detailArea }) => {
        const areaItems = hourItems.filter(
          x => normalizeAreaLabel(resolvePreferredAreaLabel(
            x?.planned_area || x?.destination_area || x?.casts?.area || '',
            x?.destination_address || x?.casts?.address || '',
            x?.destination_lat ?? x?.casts?.latitude ?? x?.casts?.lat,
            x?.destination_lng ?? x?.casts?.longitude ?? x?.casts?.lng
          )) === detailArea
        );

        html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

        areaItems.forEach(plan => {
          const linkedStatus = getPlanLinkedActualStatusCore(plan, { actuals, els, helpers });
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${buildMapLinkHtml({
               name: (plan && plan.person_name) || (plan && plan.casts && plan.casts.name),
               address: (plan && plan.destination_address) || (plan && plan.casts && plan.casts.address),
               lat: (plan && (plan.destination_lat ?? (plan.casts && (plan.casts.latitude ?? plan.casts.lat)))),
               lng: (plan && (plan.destination_lng ?? (plan.casts && (plan.casts.longitude ?? plan.casts.lng))))
               })}</strong></div>
               <div>${escapeHtml(resolvePreferredAreaLabel(
                 (plan && (plan.planned_area || plan.destination_area || (plan.casts && plan.casts.area))) || '',
                 (plan && (plan.destination_address || (plan.casts && plan.casts.address))) || '',
                 (plan && (plan.destination_lat ?? (plan.casts && (plan.casts.latitude ?? plan.casts.lat)))) ?? null,
                 (plan && (plan.destination_lng ?? (plan.casts && (plan.casts.longitude ?? plan.casts.lng)))) ?? null
               ))}</div>
               <div>${plan && plan.distance_km != null ? plan.distance_km : ''}</div>
               <div class="op-cell">
                <span class="badge-status ${linkedStatus}">${escapeHtml(getStatusText(linkedStatus))}</span>
                <button class="btn ghost plan-edit-btn" data-id="${plan && plan.id}">編集</button>
                <button class="btn ghost plan-route-btn" data-address="${escapeHtml((plan && plan.destination_address) || (plan && plan.casts && plan.casts.address) || '')}">ルート</button>
                <button class="btn danger plan-delete-btn" data-id="${plan && plan.id}">削除</button>
              </div>
            </div>
          `;
        });
      });

      html += `</div>`;
    });

    html += `</div>`;
    els.plansGroupedTable.innerHTML = html;

    const fillPlanForm = (actions && actions.fillPlanForm) || helper('fillPlanForm', null);
    const openGoogleMap = (actions && actions.openGoogleMap) || helper('openGoogleMap', null);
    const deletePlan = (actions && actions.deletePlan) || helper('deletePlan', null);

    els.plansGroupedTable.querySelectorAll('.plan-edit-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const plan = safePlans.find(x => String(x && x.id || '') === String(btn.dataset.id || ''));
        if (plan && typeof fillPlanForm === 'function') fillPlanForm(plan);
      });
    });

    els.plansGroupedTable.querySelectorAll('.plan-route-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        if (typeof openGoogleMap === 'function') openGoogleMap(btn.dataset.address || '');
      });
    });

    els.plansGroupedTable.querySelectorAll('.plan-delete-btn').forEach(btn => {
      btn.addEventListener('click', async () => {
        if (typeof deletePlan === 'function') await deletePlan(String(btn.dataset.id || ''));
      });
    });
  }

  function renderPlansTimeAreaMatrixCore({ els, plans, actuals, helpers }) {
    if (!els || !els.plansTimeAreaMatrix) return;

    const safePlans = Array.isArray(plans) ? plans : [];
    const normalizeAreaLabel = (helpers && helpers.normalizeAreaLabel) || helper('normalizeAreaLabel', v => String(v || '無し'));
    const resolvePreferredAreaLabel = (helpers && helpers.resolvePreferredAreaLabel) || helper('resolvePreferredAreaLabel', (area, address, lat, lng) => normalizeAreaLabel(area || (typeof guessArea === 'function' ? guessArea(lat, lng, address) : '無し') || '無し'));
    const getHourLabel = (helpers && helpers.getHourLabel) || helper('getHourLabel', h => `${Number(h)}時`);
    const escapeHtml = (helpers && helpers.escapeHtml) || helper('escapeHtml', v => String(v ?? ''));
    const getCurrentOriginRuntime = (helpers && helpers.getCurrentOriginRuntime) || helper('getCurrentOriginRuntime', () => null);
    const buildTimeDirectionMatrix = (helpers && helpers.buildTimeDirectionMatrix) || helper('buildTimeDirectionMatrix', null);

    if (typeof buildTimeDirectionMatrix !== 'function') {
      els.plansTimeAreaMatrix.innerHTML = `<div class="direction-ui-empty">一覧がありません</div>`;
      return;
    }

    const origin = typeof getCurrentOriginRuntime === 'function' ? getCurrentOriginRuntime() : null;
    const matrix = buildTimeDirectionMatrix(safePlans, origin, {
      maxDirections: 6,
      splitThresholdDeg: 35
    }) || { origin: origin || null, hours: [], directions: [], cells: {} };

    const hours = Array.isArray(matrix.hours) ? matrix.hours : [];
    const directions = Array.isArray(matrix.directions) ? matrix.directions : [];
    const cells = matrix.cells || {};
    const safeOriginName = String((matrix.origin && matrix.origin.name) || (origin && origin.name) || '起点').trim() || '起点';

    if (!hours.length || !directions.length) {
      els.plansTimeAreaMatrix.innerHTML = `<div class="direction-ui-empty">一覧がありません</div>`;
      return;
    }

    let html = `
      <div class="direction-ui-shell">
        <div class="direction-ui-header">
          ${getMatrixLegendHtml()}
          <div class="direction-ui-summary">
            <span class="direction-ui-origin">${escapeHtml(safeOriginName)}基準</span>
            <span class="direction-ui-pill">時間帯 ${hours.length}</span>
            <span class="direction-ui-pill">方面 ${directions.length}</span>
          </div>
        </div>
        <div class="direction-ui-table-wrap">
      <table class="matrix-table direction-ui-matrix-table">
        <thead>
          <tr>
            <th>時間</th>
            ${directions.map(direction => `<th>${escapeHtml(direction.label || '方面')}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
    `;

    hours.forEach(hour => {
      html += `<tr><td>${getHourLabel(hour)}</td>`;

      directions.forEach(direction => {
        const rows = Array.isArray(cells[`${hour}__${direction.key}`]) ? cells[`${hour}__${direction.key}`] : [];

        if (!rows.length) {
          html += `<td class="matrix-cell-empty"><span class="matrix-empty-mark">-</span></td>`;
        } else {
          html += `
            <td>
              <div class="matrix-card">
                ${rows.map(row => {
                  const source = row && row.source ? row.source : row;
                  const status = getPlanLinkedActualStatusCore(source, { actuals, els, helpers });
                  const subarea = resolvePreferredAreaLabel(
                    (source && (source.planned_area || source.destination_area || (source.casts && source.casts.area))) || row?.plannedArea || '',
                    (source && (source.destination_address || (source.casts && source.casts.address))) || '',
                    (source && (source.destination_lat ?? (source.casts && (source.casts.latitude ?? source.casts.lat)))) ?? row?.lat ?? null,
                    (source && (source.destination_lng ?? (source.casts && (source.casts.longitude ?? source.casts.lng)))) ?? row?.lng ?? null
                  );
                  return `
                    <div class="matrix-item">
                      ${buildMatrixNameLine(source, status, 'destination_address')}
                      <div class="matrix-subarea">${escapeHtml(subarea)}</div>
                    </div>
                  `;
                }).join('')}
              </div>
            </td>
          `;
        }
      });

      html += `</tr>`;
    });

    html += `</tbody></table></div></div>`;
    els.plansTimeAreaMatrix.innerHTML = html;
  }

  function renderPlanGroupedTable() {
    const ctx = resolveScheduleRendererContext();
    return renderPlanGroupedTableCore(ctx);
  }

  function renderPlansTimeAreaMatrix() {
    const ctx = resolveScheduleRendererContext();
    return renderPlansTimeAreaMatrixCore(ctx);
  }

  window.setScheduleRendererDeps = setScheduleRendererDeps;
  window.getMatrixLegendHtml = getMatrixLegendHtml;
  window.buildMatrixNameLine = buildMatrixNameLine;
  window.getPlanLinkedActualStatusCore = getPlanLinkedActualStatusCore;
  window.renderPlanGroupedTableCore = renderPlanGroupedTableCore;
  window.renderPlansTimeAreaMatrixCore = renderPlansTimeAreaMatrixCore;
  window.renderPlanGroupedTable = renderPlanGroupedTable;
  window.renderPlansTimeAreaMatrix = renderPlansTimeAreaMatrix;
})();
