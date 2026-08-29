(() => {
  'use strict';

  const PRESENTATION_VERSION = 'gf-presentation-v1-2026-08-29';
  const MAX_DISPLAY_AGE = 99;
  const originalRenderFinanceResult = window.renderFinanceResult;
  const originalFinanceCsvText = window.financeCsvText;
  const originalInvalidateFinanceResult = window.invalidateFinanceResult;

  if (
    typeof originalRenderFinanceResult !== 'function'
    || typeof originalFinanceCsvText !== 'function'
    || typeof originalInvalidateFinanceResult !== 'function'
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
    const title = document.createElement('div');
    title.className = 'gf-summary-title';
    title.textContent = '建議摘要';
    summary.prepend(title);
    summary.setAttribute('role', 'table');
    summary.querySelectorAll('.metric').forEach((metric) => metric.setAttribute('role', 'row'));
    summary.dataset.presentationVersion = PRESENTATION_VERSION;
    summary.dataset.maximumAge = String(MAX_DISPLAY_AGE);
  }

  function enhanceNotice(result) {
    const notice = document.getElementById('financeNotice');
    if (!notice) return;
    const annual = document.createElement('div');
    const total = document.createElement('div');
    annual.textContent = `年繳保費：USD ${formatMoney(result?.input?.annual)}`;
    total.textContent = `五年總保費：USD ${formatMoney(result?.input?.total)}`;
    notice.replaceChildren(annual, total);
    notice.className = 'notice gf-premium-notice';
  }

  function normalizeProposalColumns(table) {
    const originalHeaders = Array.from(table.querySelectorAll('thead th')).map((cell) => cell.textContent.trim());
    if (originalHeaders.length < 7) return null;
    const valueLabels = originalHeaders.slice(5, -1);
    table.querySelectorAll('tbody tr').forEach((row) => {
      const requestedCell = row.cells[2];
      const actualCell = row.cells[3];
      const cumulativeCell = row.cells[4];
      if (!requestedCell || !actualCell || !cumulativeCell) return;
      actualCell.title = `原定醫療保費：${requestedCell.textContent.trim()}`;
      requestedCell.remove();
      cumulativeCell.classList.add('gf-cumulative');
      row.append(cumulativeCell);
    });
    return valueLabels;
  }

  function buildGroupedHeader(table, suppliedValueLabels = null) {
    const thead = table.querySelector('thead');
    const originalHeaders = Array.from(thead?.querySelectorAll('th') || []).map((cell) => cell.textContent.trim());
    if (!thead || originalHeaders.length < 7) return;

    const fixedLabels = ['年齡', '保單年度', '實際提款醫療保費'];
    const valueLabels = suppliedValueLabels || originalHeaders.slice(5, -1);
    const statusLabel = originalHeaders.at(-1);
    const groupRow = document.createElement('tr');
    const labelRow = document.createElement('tr');

    fixedLabels.forEach((label, index) => {
      const header = document.createElement('th');
      header.textContent = label;
      header.rowSpan = 2;
      header.scope = 'col';
      if (index === 0) header.className = 'gf-sticky-age';
      if (index === 1) header.className = 'gf-sticky-year';
      groupRow.append(header);
    });

    const valueGroup = document.createElement('th');
    valueGroup.textContent = '現金提取後之保單價值';
    valueGroup.colSpan = valueLabels.length;
    valueGroup.scope = 'colgroup';
    valueGroup.className = 'gf-value-group';
    groupRow.append(valueGroup);

    const status = document.createElement('th');
    status.textContent = statusLabel;
    status.rowSpan = 2;
    status.scope = 'col';
    groupRow.append(status);

    const cumulative = document.createElement('th');
    cumulative.textContent = '累計提款';
    cumulative.rowSpan = 2;
    cumulative.scope = 'col';
    cumulative.className = 'gf-cumulative';
    groupRow.append(cumulative);

    valueLabels.forEach((label) => {
      const header = document.createElement('th');
      header.textContent = label === '提款後總額' ? '總額' : label;
      header.scope = 'col';
      labelRow.append(header);
    });

    thead.replaceChildren(groupRow, labelRow);
  }

  function installColumnWidths(table) {
    table.querySelector('colgroup')?.remove();
    const count = table.querySelector('tbody tr')?.cells.length || 0;
    if (!count) return;
    const group = document.createElement('colgroup');
    for (let index = 0; index < count; index += 1) {
      const column = document.createElement('col');
      if (index === 0) column.className = 'gf-age-column';
      if (index === 1) column.className = 'gf-year-column';
      if (index === count - 1) column.className = 'gf-cumulative-column';
      group.append(column);
    }
    table.prepend(group);
  }

  function enhanceTable() {
    const table = document.getElementById('financeTable');
    if (!table) return;
    const valueLabels = normalizeProposalColumns(table);
    buildGroupedHeader(table, valueLabels);
    installColumnWidths(table);
    table.querySelectorAll('tbody tr').forEach((row) => {
      const age = Number(row.cells[0]?.textContent.replace(/,/g, '').trim());
      if (Number.isFinite(age) && age > MAX_DISPLAY_AGE) {
        row.remove();
        return;
      }
      row.cells[0]?.classList.add('gf-sticky-age');
      row.cells[1]?.classList.add('gf-sticky-year');
      const policyYear = Number(row.cells[1]?.textContent.replace(/,/g, '').trim());
      if (Number.isFinite(policyYear) && policyYear % 5 === 0) {
        row.classList.add('gf-five-year-divider');
      }
    });
    table.dataset.presentationVersion = PRESENTATION_VERSION;
    table.dataset.maximumAge = String(MAX_DISPLAY_AGE);
    table.setAttribute('aria-label', 'GF 醫療融資逐年建議表');
    table.closest('.tableWrap')?.classList.add('gf-finance-table-wrap');

    const notice = document.getElementById('financeAgeLimitNote') || document.createElement('p');
    notice.id = 'financeAgeLimitNote';
    notice.className = 'gf-age-limit-note';
    notice.textContent = '逐年結果及下載 CSV 顯示至被保人 99 歲。';
    table.closest('.tableWrap')?.after(notice);
  }

  function setResultsVisible(visible) {
    const summary = document.getElementById('financeSummary');
    const table = document.getElementById('financeTable');
    const wrap = table?.closest('.tableWrap');
    const note = document.getElementById('financeAgeLimitNote');
    if (summary) summary.hidden = !visible;
    if (wrap) wrap.hidden = !visible;
    if (note) note.hidden = !visible;
  }

  function setPdfDownloadEnabled(enabled) {
    const button = document.getElementById('financePdfDownload');
    if (button) button.disabled = !enabled;
  }

  async function downloadMedicalFinancingPdf() {
    if (!currentFinanceResult) return;
    const button = document.getElementById('financePdfDownload');
    const originalLabel = button?.textContent || '下載 PDF';
    if (button) {
      button.disabled = true;
      button.textContent = '製作 PDF…';
    }
    try {
      const response = await fetch('/api/gf/medical-financing-pdf', {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json', Accept: 'application/pdf' },
        body: JSON.stringify({
          result: visibleResult(currentFinanceResult),
          medicalContext: currentMedicalContext || { source: 'manual' },
        }),
      });
      if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        throw new Error(error.error || `PDF 製作失敗（${response.status}）`);
      }
      const blob = await response.blob();
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'gf_medical_financing_proposal.pdf';
      link.click();
      URL.revokeObjectURL(link.href);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error));
    } finally {
      if (button) {
        button.textContent = originalLabel;
        button.disabled = !currentFinanceResult;
      }
    }
  }

  function installPdfDownloadButton() {
    if (document.getElementById('financePdfDownload')) return;
    const csvButton = document.getElementById('financeDownload');
    if (!csvButton) return;
    const button = document.createElement('button');
    button.id = 'financePdfDownload';
    button.type = 'button';
    button.className = 'secondary';
    button.textContent = '下載 PDF';
    button.disabled = true;
    button.addEventListener('click', downloadMedicalFinancingPdf);
    csvButton.after(button);
  }

  function patchedInvalidateFinanceResult(...args) {
    const value = originalInvalidateFinanceResult.apply(this, args);
    setResultsVisible(false);
    setPdfDownloadEnabled(false);
    return value;
  }

  function patchedRenderFinanceResult(result, ...args) {
    const value = originalRenderFinanceResult.call(this, result, ...args);
    enhanceSummary(result);
    enhanceNotice(result);
    enhanceTable();
    setResultsVisible(true);
    setPdfDownloadEnabled(true);
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
  window.invalidateFinanceResult = patchedInvalidateFinanceResult;
  if (window.GFFinanceCore) {
    window.GFFinanceCore.financeCsvText = patchedFinanceCsvText;
  }

  window.GFPresentationPatch = Object.freeze({
    PRESENTATION_VERSION,
    MAX_DISPLAY_AGE,
    visibleRows,
    visibleResult,
    buildGroupedHeader,
    enhanceNotice,
    installPdfDownloadButton,
    installColumnWidths,
    normalizeProposalColumns,
  });

  installPdfDownloadButton();
  setResultsVisible(false);
})();
