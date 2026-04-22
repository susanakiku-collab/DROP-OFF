// renderActual
// dashboard.js から安全に分離した実績表描画系
// core は依存注入で動作し、window へラッパー関数を公開する

(function(){
  const actualRendererDeps = {
    getEls: () => (typeof els !== "undefined" ? els : null),
    getActuals: () => (typeof currentActualsCache !== "undefined" ? currentActualsCache : []),
    getActions: () => ({
      fillActualForm: typeof fillActualForm === "function" ? fillActualForm : null,
      openGoogleMap: typeof openGoogleMap === "function" ? openGoogleMap : null,
      deleteActual: typeof deleteActual === "function" ? deleteActual : null,
      updateActualStatus: typeof updateActualStatus === "function" ? updateActualStatus : null
    }),
    getHelpers: () => ({
      normalizeStatus: typeof normalizeStatus === "function" ? normalizeStatus : (v => v),
      getHourLabel: typeof getHourLabel === "function" ? getHourLabel : (h => `${h}時`),
      getGroupedAreasByDisplay: typeof getGroupedAreasByDisplay === "function" ? getGroupedAreasByDisplay : (() => []),
      getGroupedAreaHeaderHtml: typeof getGroupedAreaHeaderHtml === "function" ? getGroupedAreaHeaderHtml : (v => String(v || "")),
      buildMapLinkHtml: typeof buildMapLinkHtml === "function" ? buildMapLinkHtml : ({ name }) => String(name || "-"),
      escapeHtml: typeof escapeHtml === "function" ? escapeHtml : (v => String(v ?? "")),
      normalizeAreaLabel: typeof normalizeAreaLabel === "function" ? normalizeAreaLabel : (v => String(v || "")),
      getStatusText: typeof getStatusText === "function" ? getStatusText : (v => String(v || "")),
      getAreaDisplayGroup: typeof getAreaDisplayGroup === "function" ? getAreaDisplayGroup : (v => String(v || "")),
      getMatrixLegendHtml: typeof getMatrixLegendHtml === "function" ? getMatrixLegendHtml : (() => ""),
      buildMatrixNameLine: typeof buildMatrixNameLine === "function" ? buildMatrixNameLine : (() => ""),
      getCurrentOriginRuntime: typeof getCurrentOriginRuntime === "function" ? getCurrentOriginRuntime : (() => null),
      buildTimeDirectionMatrix: typeof buildTimeDirectionMatrix === "function" ? buildTimeDirectionMatrix : null
    })
  };

  function setActualRendererDeps(nextDeps = {}) {
    Object.assign(actualRendererDeps, nextDeps || {});
  }

  function resolveActualRendererContext() {
    const elsObj = typeof actualRendererDeps.getEls === "function" ? actualRendererDeps.getEls() : null;
    const actuals = typeof actualRendererDeps.getActuals === "function" ? actualRendererDeps.getActuals() : [];
    const actions = typeof actualRendererDeps.getActions === "function" ? actualRendererDeps.getActions() : {};
    const helpers = typeof actualRendererDeps.getHelpers === "function" ? actualRendererDeps.getHelpers() : {};
    return {
      els: elsObj || {},
      actuals: Array.isArray(actuals) ? actuals : [],
      actions: actions || {},
      helpers: helpers || {}
    };
  }

  function renderActualTableCore({ els, actuals, actions, helpers }) {
    if (!els.actualTableWrap) return;

    const {
      normalizeStatus,
      getHourLabel,
      getGroupedAreasByDisplay,
      getGroupedAreaHeaderHtml,
      buildMapLinkHtml,
      escapeHtml,
      normalizeAreaLabel,
      getStatusText
    } = helpers;
    const resolvePreferredAreaLabel = helpers?.resolvePreferredAreaLabel || (typeof window !== 'undefined' && typeof window.resolvePreferredAreaLabel === 'function'
      ? window.resolvePreferredAreaLabel
      : ((area, address, lat, lng) => normalizeAreaLabel(area || (typeof guessArea === 'function' ? guessArea(lat, lng, address) : '無し') || '無し')));

    const tableItems = Array.isArray(actuals) ? actuals : [];

    if (!tableItems.length) {
      els.actualTableWrap.innerHTML = `<div class="muted" style="padding:14px;">Actual一覧はありません</div>`;
      return;
    }

    const hours = [...new Set(tableItems.map(x => Number(x?.actual_hour ?? 0)))].sort((a, b) => a - b);
    let html = `<div class="grouped-actual-list">`;

    hours.forEach(hour => {
      const hourItems = tableItems.filter(x => Number(x?.actual_hour ?? 0) === hour);
      const groupedAreas = getGroupedAreasByDisplay(hourItems, x => resolvePreferredAreaLabel(
        x?.destination_area || x?.planned_area || x?.casts?.area || '',
        x?.destination_address || x?.casts?.address || '',
        x?.destination_lat ?? x?.casts?.latitude ?? x?.casts?.lat,
        x?.destination_lng ?? x?.casts?.longitude ?? x?.casts?.lng
      ));

      html += `<div class="grouped-section">`;
      html += `<div class="grouped-hour-title">${getHourLabel(hour)}</div>`;

      groupedAreas.forEach(({ detailArea }) => {
        const areaItems = hourItems.filter(
          item => normalizeAreaLabel(resolvePreferredAreaLabel(
            item?.destination_area || item?.planned_area || item?.casts?.area || '',
            item?.destination_address || item?.casts?.address || '',
            item?.destination_lat ?? item?.casts?.latitude ?? item?.casts?.lat,
            item?.destination_lng ?? item?.casts?.longitude ?? item?.casts?.lng
          )) === detailArea
        );

        if (!areaItems.length) return;

        html += `<div class="grouped-area-title">${getGroupedAreaHeaderHtml(detailArea)}</div>`;

        areaItems.forEach(item => {
          const itemId = escapeHtml(item?.id || "");
          const normalizedStatus = normalizeStatus(item?.status);
          const overflowText = Number(item?.vehicle_id || 0) <= 0 && String(item?.driver_name || '').includes('あぶれ')
            ? `<span class="badge-status pending" style="margin-left:6px;">あぶれ</span>`
            : '';
          html += `
            <div class="grouped-row">
              <div>${getHourLabel(hour)}</div>
              <div><strong>${buildMapLinkHtml({
                name: item?.casts?.name,
                address: item?.destination_address || item?.casts?.address,
                lat: item?.casts?.latitude,
                lng: item?.casts?.longitude
              })}</strong>${overflowText}</div>
              <div>${escapeHtml(resolvePreferredAreaLabel(
                item?.destination_area || item?.planned_area || item?.casts?.area || '',
                item?.destination_address || item?.casts?.address || '',
                item?.destination_lat ?? item?.casts?.latitude ?? item?.casts?.lat,
                item?.destination_lng ?? item?.casts?.longitude ?? item?.casts?.lng
              ))}</div>
              <div>${item?.distance_km ?? ""}</div>
              <div class="op-cell">
                <div class="state-stack">
                  <button class="btn ghost actual-pending-btn" data-id="${itemId}">未完了</button>
                  <button class="btn primary actual-done-btn" data-id="${itemId}">完了</button>
                  <button class="btn danger actual-cancel-btn" data-id="${itemId}">キャンセル</button>
                  <span class="badge-status ${normalizedStatus}">${escapeHtml(getStatusText(item?.status))}</span>
                </div>
                <button class="btn ghost actual-edit-btn" data-id="${itemId}">編集</button>
                <button class="btn ghost actual-route-btn" data-address="${escapeHtml(item?.destination_address || item?.casts?.address || "")}">ルート</button>
                <button class="btn danger actual-delete-btn" data-id="${itemId}">削除</button>
              </div>
            </div>
          `;
        });
      });

      html += `</div>`;
    });

    html += `</div>`;
    els.actualTableWrap.innerHTML = html;

    els.actualTableWrap.querySelectorAll(".actual-edit-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const item = actuals.find(x => String(x?.id || "") === String(btn.dataset.id || ""));
        if (item && typeof actions.fillActualForm === "function") actions.fillActualForm(item);
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-route-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        if (typeof actions.openGoogleMap === "function") actions.openGoogleMap(btn.dataset.address || "");
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-delete-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.deleteActual === "function") await actions.deleteActual(String(btn.dataset.id || ""));
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-pending-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.updateActualStatus === "function") await actions.updateActualStatus(String(btn.dataset.id || ""), "pending");
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-done-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.updateActualStatus === "function") await actions.updateActualStatus(String(btn.dataset.id || ""), "done");
      });
    });

    els.actualTableWrap.querySelectorAll(".actual-cancel-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        if (typeof actions.updateActualStatus === "function") await actions.updateActualStatus(String(btn.dataset.id || ""), "cancel");
      });
    });
  }

  function renderActualTimeAreaMatrixCore({ els, actuals, helpers }) {
    if (!els.actualTimeAreaMatrix) return;

    const normalizeAreaLabel = helpers?.normalizeAreaLabel || (v => String(v || '無し'));
    const resolvePreferredAreaLabel = helpers?.resolvePreferredAreaLabel || (typeof window !== 'undefined' && typeof window.resolvePreferredAreaLabel === 'function'
      ? window.resolvePreferredAreaLabel
      : ((area, address, lat, lng) => normalizeAreaLabel(area || (typeof guessArea === 'function' ? guessArea(lat, lng, address) : '無し') || '無し')));
    const escapeHtml = helpers?.escapeHtml || (v => String(v ?? ''));
    const getHourLabel = helpers?.getHourLabel || (h => `${Number(h)}時`);
    const getMatrixLegendHtml = helpers?.getMatrixLegendHtml || (() => '');
    const buildMatrixNameLine = helpers?.buildMatrixNameLine || (() => '');
    const normalizeStatus = helpers?.normalizeStatus || (v => v || 'pending');
    const getCurrentOriginRuntime = helpers?.getCurrentOriginRuntime || (() => null);
    const buildTimeDirectionMatrix = helpers?.buildTimeDirectionMatrix || (typeof window !== 'undefined' ? window.buildTimeDirectionMatrix : null);

    if (typeof buildTimeDirectionMatrix !== 'function') {
      els.actualTimeAreaMatrix.innerHTML = `<div class="direction-ui-empty">一覧がありません</div>`;
      return;
    }

    const safeActuals = Array.isArray(actuals) ? actuals : [];
    const origin = typeof getCurrentOriginRuntime === 'function' ? getCurrentOriginRuntime() : null;
    const matrix = buildTimeDirectionMatrix(safeActuals, origin, {
      maxDirections: 6,
      splitThresholdDeg: 35
    }) || { origin: origin || null, hours: [], directions: [], cells: {} };

    const hours = Array.isArray(matrix.hours) ? matrix.hours : [];
    const directions = Array.isArray(matrix.directions) ? matrix.directions : [];
    const cells = matrix.cells || {};
    const safeOriginName = String((matrix.origin && matrix.origin.name) || (origin && origin.name) || '起点').trim() || '起点';

    if (!hours.length || !directions.length) {
      els.actualTimeAreaMatrix.innerHTML = `<div class="direction-ui-empty">一覧がありません</div>`;
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
                  const status = normalizeStatus(source?.status || row?.status);
                  const subarea = resolvePreferredAreaLabel(
                    (source && (source.destination_area || source.planned_area || (source.casts && source.casts.area))) || row?.plannedArea || '',
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
    els.actualTimeAreaMatrix.innerHTML = html;
  }

  function renderActualTable() {
    return renderActualTableCore(resolveActualRendererContext());
  }

  function renderActualTimeAreaMatrix() {
    return renderActualTimeAreaMatrixCore(resolveActualRendererContext());
  }

  window.setActualRendererDeps = setActualRendererDeps;
  window.renderActualTableCore = renderActualTableCore;
  window.renderActualTimeAreaMatrixCore = renderActualTimeAreaMatrixCore;
  window.renderActualTable = renderActualTable;
  window.renderActualTimeAreaMatrix = renderActualTimeAreaMatrix;
})();
