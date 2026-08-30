(() => {
  'use strict';

  const PLAN_VERSION = 'gf-payment-plan-v1-2026-08-29';
  const FIVE_YEAR = 'five_year';
  const SINGLE = 'single';
  const singleApi = window.GFSingleExperiment;
  const fiveCore = window.GFFinanceCore;
  const singleRender = window.renderFinanceResult;
  const singleCsv = window.financeCsvText;
  const sharedInvalidate = window.invalidateFinanceResult;

  if (!singleApi || !fiveCore || typeof singleRender !== 'function' || typeof singleCsv !== 'function') {
    throw new Error('GF 雙計劃切換無法連接五年及一次性繳費模型。');
  }

  function number(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function activePlan() {
    return document.getElementById('gfPaymentPlan')?.value || FIVE_YEAR;
  }

  function tagResult(result, paymentMode) {
    if (!result?.input) return result;
    result.paymentMode = paymentMode;
    result.paymentPlanVersion = PLAN_VERSION;
    result.input.paymentMode = paymentMode;
    result.input.paymentPlanVersion = PLAN_VERSION;
    return result;
  }

  function readMedicalInputs(paymentMode) {
    const issueAge = Number(document.getElementById('issueAge').value);
    const startAge = Number(document.getElementById('medicalStartAge').value);
    const endAge = Number(document.getElementById('medicalEndAge').value);
    if (!Number.isInteger(endAge) || endAge > 99) throw new Error('醫療保費終止年齡最多為 99 歲。');
    const automatic = document.getElementById('medicalPremiumSource').value === 'table';
    const context = automatic ? selectedMedicalTableContext() : { source: 'manual' };
    const medicalRows = automatic
      ? medicalRowsFromTable(context.planId, context.formId, context.deductibleId, startAge, endAge)
      : parseMedicalPremiumSchedule(document.getElementById('medicalPremiumTable').value, startAge, endAge);
    if (paymentMode === SINGLE) singleApi.validateMedicalInputs(issueAge, medicalRows);
    else fiveCore.validateMedicalInputs(issueAge, medicalRows);
    return { issueAge, medicalRows, context };
  }

  function renderFiveYearResult(result, proof = null) {
    currentFinanceResult = result;
    currentMinimumProof = proof;
    const summary = document.getElementById('financeSummary');
    clearElement(summary);
    appendMetric(summary, 'GF 年繳保費', `USD ${fmt(result.input.annual)}`);
    appendMetric(summary, '5 年總供款', `USD ${fmt(result.input.total)}`);
    appendMetric(summary, '投保時基本金額', `USD ${fmt(result.input.basic)}`);
    appendMetric(summary, '醫療保費總額', `USD ${fmt(result.requestedTotal)}`);
    appendMetric(summary, '實際累計提款', `USD ${fmt(result.actualTotal)}`);
    appendMetric(summary, '最終狀態', result.surrendered ? `第 ${result.surrenderYear} 年退保` : '保單持續');
    if (proof) appendMetric(summary, '最低值驗證', `USD ${fmt(proof.lowerPremium)} 不足`);
    if (currentMedicalContext?.source === 'table') {
      appendMetric(summary, '醫療保費表', `${currentMedicalContext.formName} · ${currentMedicalContext.deductibleLabel}`);
    }
    const notice = document.getElementById('financeNotice');
    notice.className = result.surrendered ? 'notice bad' : 'notice';
    notice.textContent = result.surrendered
      ? `第 ${result.surrenderYear} 保單年度可用價值不足，該年已提取全部剩餘價值並視作退保。`
      : (proof
        ? `USD ${proof.annualPremium} 可完成至 99 歲的全部提款；低 1 元的 USD ${proof.lowerPremium} 未能完成。`
        : '指定五年供款情境已完成。');
    const table = document.getElementById('financeTable');
    clearElement(table);
    const components = financeComponentMode() === 'components';
    const headers = ['年齡', '保單年度', '醫療保費', '實際提款', '累計提款'];
    if (components) headers.push('保證現金價值', '復歸紅利', '終期紅利');
    headers.push('提款後總額', '狀態');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.forEach((header) => appendCell(headerRow, header, 'th'));
    thead.appendChild(headerRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    result.rows.forEach((item) => {
      const row = document.createElement('tr');
      if (item.status === '當年退保' || item.status === '已退保') row.className = 'force-row';
      else if (number(item.requested_withdrawal) > 0) row.className = 'active-row';
      appendCell(row, item.age);
      appendCell(row, item.policy_year);
      appendCell(row, fmt(item.requested_withdrawal));
      appendCell(row, fmt(item.withdrawal_total));
      appendCell(row, fmt(item.cumulative_withdrawal));
      if (components) {
        appendCell(row, fmt(item.post_guaranteed_cash));
        appendCell(row, fmt(item.post_reversionary_bonus));
        appendCell(row, fmt(item.post_terminal_bonus));
      }
      appendCell(row, fmt(item.post_surrender_total));
      appendCell(row, item.status, 'td', item.status.includes('退保') ? 'status-surrendered' : 'status-active');
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    document.getElementById('financeDownload').disabled = false;
  }

  function invalidateCombinedResult(...args) {
    const value = sharedInvalidate.apply(this, args);
    document.getElementById('experimentMethod')?.remove();
    return value;
  }

  function renderCombinedResult(result, proof = null) {
    const paymentMode = result?.input?.paymentMode || result?.paymentMode || activePlan();
    tagResult(result, paymentMode);
    if (paymentMode === SINGLE) return singleRender(result, proof);
    return renderFiveYearResult(result, proof);
  }

  function combinedCsv(result, medicalContext) {
    const paymentMode = result?.input?.paymentMode || result?.paymentMode || activePlan();
    if (paymentMode === SINGLE) return singleCsv(result, medicalContext);
    return fiveCore.financeCsvText(result, medicalContext);
  }

  function dispatchInvalidate() {
    const installed = window.invalidateFinanceResult;
    return installed && installed !== invalidateCombinedResult
      ? installed()
      : invalidateCombinedResult();
  }

  function dispatchRender(result, proof = null) {
    const installed = window.renderFinanceResult;
    return installed && installed !== renderCombinedResult
      ? installed(result, proof)
      : renderCombinedResult(result, proof);
  }

  function runCombinedFinance() {
    dispatchInvalidate();
    const error = document.getElementById('financeError');
    error.hidden = true;
    error.textContent = '';
    try {
      const paymentMode = activePlan();
      const { issueAge, medicalRows, context } = readMedicalInputs(paymentMode);
      currentMedicalContext = context;
      const mode = document.getElementById('financeMode').value;
      if (paymentMode === SINGLE) {
        if (mode === 'minimum') {
          const proof = singleApi.findMinimumSinglePremium(issueAge, medicalRows);
          proof.result.calculationMode = 'minimum_single';
          dispatchRender(tagResult(proof.result, SINGLE), proof);
        } else {
          const premium = Number(document.getElementById('fixedAnnualPremium').value);
          if (!Number.isInteger(premium) || premium <= 0) throw new Error('一次性繳費請輸入大於 0 的整數美元。');
          const result = singleApi.simulateSinglePremium(premium, issueAge, medicalRows);
          result.calculationMode = 'fixed_single';
          dispatchRender(tagResult(result, SINGLE));
        }
      } else if (mode === 'minimum') {
        const proof = fiveCore.findMinimumAnnualPremium(issueAge, medicalRows);
        proof.result.calculationMode = 'minimum_five_year';
        dispatchRender(tagResult(proof.result, FIVE_YEAR), proof);
      } else {
        const annual = Number(document.getElementById('fixedAnnualPremium').value);
        if (!Number.isFinite(annual) || annual <= 0) throw new Error('請輸入大於 0 的 GF 年繳保費。');
        const result = fiveCore.simulateMedicalFinancing(annual, issueAge, medicalRows);
        result.calculationMode = 'fixed_five_year';
        dispatchRender(tagResult(result, FIVE_YEAR));
      }
    } catch (value) {
      dispatchInvalidate();
      error.textContent = value instanceof Error ? value.message : String(value);
      error.hidden = false;
    }
  }

  function updatePlanUi({ initial = false } = {}) {
    const paymentMode = activePlan();
    const mode = document.getElementById('financeMode');
    const premium = document.getElementById('fixedAnnualPremium');
    const fixedLabel = document.querySelector('label[for="fixedAnnualPremium"]');
    const resultTitle = document.querySelector('main.panel h2');
    const issue = Math.max(0, Math.round(number(document.getElementById('issueAge').value)));
    const start = document.getElementById('medicalStartAge');
    const end = document.getElementById('medicalEndAge');
    end.max = '99';
    if (number(end.value) > 99) end.value = '99';
    if (paymentMode === SINGLE) {
      mode.innerHTML = '<option value="minimum">反推最低 GF 一次性繳費</option><option value="fixed">指定 GF 一次性繳費</option>';
      if (fixedLabel) fixedLabel.textContent = 'GF 一次性繳費（USD）';
      premium.value = '30000';
      premium.min = '7500';
      if (initial || number(start.value) < issue + 2) start.value = String(Math.min(99, issue + 2));
      if (resultTitle) resultTitle.textContent = '一次性繳費反推結果';
    } else {
      mode.innerHTML = '<option value="minimum">反推最低 GF 年繳保費</option><option value="fixed">指定 GF 年繳保費</option>';
      if (fixedLabel) fixedLabel.textContent = 'GF 年繳保費（USD，5 年繳）';
      premium.value = '20000';
      premium.min = '1';
      if (initial || number(start.value) < issue + 6) start.value = String(Math.min(99, issue + 6));
      if (resultTitle) resultTitle.textContent = '五年供款反推結果';
    }
    mode.value = 'minimum';
    document.getElementById('fixedPremiumWrap').hidden = true;
    document.title = 'GF 醫療融資';
    document.querySelector('.top h1').textContent = 'GF 醫療融資';
    document.querySelector('.top p').textContent = '選擇一次性繳費或五年供款，反推足以支付醫療保費至 99 歲的最低 GF 保費，或測試指定保費。';
    document.querySelector('.badge').textContent = '五年供款＋一次性繳費 · 預期情況';
    const banner = document.querySelector('.experiment-banner');
    if (banner) banner.innerHTML = '<b>計算提示</b><small>兩個計劃均按保單年度年末提款，逐年結果、CSV及PDF只顯示至99歲；正式銷售前請以PGS建議書核對。</small>';
    dispatchInvalidate();
    const notice = document.getElementById('financeNotice');
    notice.textContent = paymentMode === SINGLE
      ? '設定年齡及醫療保費後按「開始計算」。一次性繳費最早由第 2 保單年度年末提款。'
      : '設定年齡及醫療保費後按「開始計算」。五年供款最早由第 6 保單年度年末提款。';
  }

  function installPlanSelector() {
    if (document.getElementById('gfPaymentPlan')) return;
    const form = document.querySelector('.finance-form');
    const heading = form?.querySelector('h2');
    if (!form || !heading) throw new Error('找不到 GF 醫療融資設定表。');
    const wrapper = document.createElement('div');
    wrapper.className = 'gf-payment-plan';
    wrapper.innerHTML = '<label for="gfPaymentPlan">GF 供款方式</label><select id="gfPaymentPlan"><option value="five_year">五年供款</option><option value="single">一次性繳費</option></select><div class="gf-plan-hint" id="gfPaymentPlanHint">兩個計劃使用各自獨立的正式案例校準模型。</div>';
    heading.after(wrapper);
    const selector = document.getElementById('gfPaymentPlan');
    selector.value = FIVE_YEAR;
    selector.addEventListener('change', () => updatePlanUi());
    const oldRun = document.getElementById('financeRun');
    const newRun = oldRun.cloneNode(true);
    oldRun.replaceWith(newRun);
    newRun.addEventListener('click', runCombinedFinance);
    updatePlanUi({ initial: true });
  }

  window.invalidateFinanceResult = invalidateCombinedResult;
  window.renderFinanceResult = renderCombinedResult;
  window.financeCsvText = combinedCsv;
  window.GFPaymentPlan = Object.freeze({
    PLAN_VERSION,
    FIVE_YEAR,
    SINGLE,
    activePlan,
    tagResult,
    readMedicalInputs,
    runCombinedFinance,
    updatePlanUi,
  });

  installPlanSelector();

  if (new URLSearchParams(location.search).get('selftest') === '1') {
    const checks = [
      ['雙計劃選擇已安裝', document.getElementById('gfPaymentPlan')?.options.length === 2],
      ['正式頁預設保留五年供款', activePlan() === FIVE_YEAR],
      ['一次性模型已載入', typeof singleApi.findMinimumSinglePremium === 'function'],
      ['五年模型已保留', typeof fiveCore.findMinimumAnnualPremium === 'function'],
      ['最高醫療保費年齡為99歲', document.getElementById('medicalEndAge')?.max === '99'],
    ];
    const report = document.createElement('pre');
    report.id = 'paymentPlanSelfTestResults';
    report.className = 'panel self-test';
    report.textContent = checks.map(([name, pass]) => `${pass ? 'PASS' : 'FAIL'} ${name}`).join('\n') + `\n\n${checks.filter(([, pass]) => pass).length}/${checks.length} passed`;
    document.querySelector('.advanced')?.before(report);
  }
})();
