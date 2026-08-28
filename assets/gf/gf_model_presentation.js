(() => {
  'use strict';

  const PRESENTATION_VERSION = 'gf-presentation-v1-2026-08-29';
  const MAX_DISPLAY_AGE = 99;
  const originalRenderFinanceResult = window.renderFinanceResult;
  const originalFinanceCsvText = window.financeCsvText;

  if (
    typeof originalRenderFinanceResult !== 'function'
    || typeof originalFinanceCsvText !== 'function'
  ) {
    throw new Error('GF 結果介面補丁無法連接原有顯示及 CSV 匯出。');
  }

  function number(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function visibleRows(result) {
    if (!Array.isArray(result?.rows)) return [];
    return result.rows.filter((row) => Number(row?.age) <= MAX_DISPLAY_AGE);
  }

  function visibleResult(result) {
    if (!result || !Array.isArray(result.rows)) return result;
    const rows = visibleRows(result);
    const lastRow = rows.at(-1);
    const surrenderRow = rows.find((row) => row.status === '當年退保');
    const visiblePolicyYears = new Set(rows.map((row) => number(row.policy_year)));
    const schedule = Object.fromEntries(
      Object.entries(result.schedule || {}).filter(([policyYear]) => (
        visiblePolicyYears.has(number(policyYear))
      )),
    );
    return {
      ...result,
      rows,
      schedule,
      requestedTotal: rows.reduce((total, row) => total + number(row.requested_withdrawal), 0),
      actualTotal: lastRow ? number(lastRow.cumulative_withdrawal) : 0,
      surrendered: Boolean(surrenderRow),
      surrenderYear: surrenderRow ? number(surrenderRow.policy_year) : null,
      presentationVersion: PRESENTATION_VERSION,
      maximumDisplayedAge: MAX_DISPLAY_AGE,
    };
  }

  function formatMoney(value) {
    return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(number(value));
  }

  function updateMetric(summary, label, value, replacementLabel = label) {
    const card = Array.from(summary?.children || []).find((item) => (
      item.querySelector('span')?.textContent === label
    ));
    if (!card) return;
    const title = card.querySelector('span');
    const amount = card.querySelector('b');
    if (title) title.textContent = replacementLabel;
    if (amount) amount.textContent = value;
  }

  function enhanceSummary(result) {
    const summary = document.getElementById('financeSummary');
    if (!summary) return;
    const display = visibleResult(result);
    updateMetric(summary, '醫療保費總額', `USD ${formatMoney(display.requestedTotal)}`);
    updateMetric(summary, '實際累計提款', `USD ${formatMoney(display.actualTotal)}`);
    updateMetric(
      summary,
      '最終狀態',
      display.surrendered ? `第 ${display.surrenderYear} 年退保` : '保單持續',
      '截至 99 歲狀態',
    );
    summary.dataset.presentationVersion = PRESENTATION_VERSION;
    summary.dataset.maximumAge = String(MAX_DISPLAY_AGE);
  }

  function enhanceTable() {
    const table = document.getElementById('financeTable');
    if (!table) return;
    const headers = Array.from(table.querySelectorAll('thead th')).map((cell) => cell.textContent.trim());
    table.querySelectorAll('tbody tr').forEach((row) => {
      const age = Number(row.cells[0]?.textContent.replace(/,/g, '').trim());
      if (Number.isFinite(age) && age > MAX_DISPLAY_AGE) {
        row.remove();
        return;
      }
      Array.from(row.cells).forEach((cell, index) => {
        cell.dataset.label = headers[index] || '';
      });
    });
    table.dataset.presentationVersion = PRESENTATION_VERSION;
    table.dataset.maximumAge = String(MAX_DISPLAY_AGE);
    table.closest('.tableWrap')?.classList.add('gf-finance-table-wrap');

    const notice = document.getElementById('financeAgeLimitNote') || document.createElement('p');
    notice.id = 'financeAgeLimitNote';
    notice.className = 'gf-age-limit-note';
    notice.textContent = '逐年結果及下載 CSV 顯示至被保人 99 歲。';
    table.closest('.tableWrap')?.after(notice);
  }

  function patchedRenderFinanceResult(result, ...args) {
    const value = originalRenderFinanceResult.call(this, result, ...args);
    enhanceSummary(result);
    enhanceTable();
    return value;
  }

  function patchedFinanceCsvText(result, ...args) {
    const csv = originalFinanceCsvText.call(this, visibleResult(result), ...args);
    if (typeof csv !== 'string') return csv;
    const newline = csv.includes('\r\n') ? '\r\n' : '\n';
    const hasBom = csv.startsWith('\uFEFF');
    const lines = (hasBom ? csv.slice(1) : csv).split(/\r?\n/);
    const headerIndex = lines.findIndex((line) => line.startsWith('"age","policy_year",'));
    if (headerIndex < 0) throw new Error('GF CSV 欠缺逐年表頭，99 歲顯示補丁停止匯出。');
    const insertAt = headerIndex > 0 && lines[headerIndex - 1] === '' ? headerIndex - 1 : headerIndex;
    lines.splice(
      insertAt,
      0,
      `"presentation_version","${PRESENTATION_VERSION}"`,
      `"maximum_display_age","${MAX_DISPLAY_AGE}"`,
    );
    return `${hasBom ? '\uFEFF' : ''}${lines.join(newline)}`;
  }

  window.renderFinanceResult = patchedRenderFinanceResult;
  window.financeCsvText = patchedFinanceCsvText;
  if (window.GFFinanceCore) {
    window.GFFinanceCore.financeCsvText = patchedFinanceCsvText;
  }

  window.GFPresentationPatch = Object.freeze({
    PRESENTATION_VERSION,
    MAX_DISPLAY_AGE,
    visibleRows,
    visibleResult,
  });
})();
