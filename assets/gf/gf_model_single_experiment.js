(() => {
  'use strict';

  const D = window.GFSingleOfficialData;
  if (!D) throw new Error('一次性繳費官方校準資料未載入。');
  const MODEL_VERSION = 'gf-single-state-transition-v0.4-2026-08-29';
  const SOURCE_BASIC = D.sourceBasic;
  const MIN_BASIC = D.minimumBasic;
  const SURVIVAL_BASIC = 7501;
  const MIN_BASIC_BRACKET = '官方投保最低 USD 7,500；三個正式提款案例校準持續閘門 USD 7,501';
  const MAX_YEAR = D.baseline.length;
  const CAP_RATE = 1.065;
  const BASE = new Map(D.baseline.map(row => [row[0], { y: row[0], g: row[1], r: row[2], t: row[3] }]));
  const OFFICIAL_ANCHORS = [
    { basic: SOURCE_BASIC, values: BASE },
    D.secondaryAnchor ? { basic: D.secondaryAnchor.sourceBasic, values: new Map(D.secondaryAnchor.baseline.map(row => [row[0], { y: row[0], g: row[1], r: row[2], t: row[3] }])) } : null,
    D.tertiaryAnchor ? { basic: D.tertiaryAnchor.sourceBasic, values: new Map(D.tertiaryAnchor.baseline.map(row => [row[0], { y: row[0], g: row[1], r: row[2], t: row[3] }])) } : null,
  ].filter(Boolean);
  const R_BANDS = [
    { start: 3, end: 9, a: 1.015, b: .0081 },
    { start: 10, end: 39, a: 1.0245, b: .0004 },
    { start: 40, end: 72, a: 1.0045, b: .0045 },
  ];
  const DISCOUNT_TIERS = [
    { low: 0, high: 34999, discount: 0 }, { low: 35000, high: 74999, discount: 10 },
    { low: 75000, high: 149999, discount: 15 }, { low: 150000, high: 249999, discount: 18 },
    { low: 250000, high: 499999, discount: 21 }, { low: 500000, high: 999999, discount: 24 },
    { low: 1000000, high: 1499999, discount: 27 }, { low: 1500000, high: Infinity, discount: 30 },
  ];
  const originalInvalidate = invalidateFinanceResult;

  function tierForBasic(basic) {
    return DISCOUNT_TIERS.find(tier => basic >= tier.low && basic <= tier.high) || DISCOUNT_TIERS.at(-1);
  }
  function premiumForBasic(basic) {
    const tier = tierForBasic(basic);
    return roundMoney(Number(basic) * (1000 - tier.discount) / 1000);
  }
  function inputForIntegerBasic(value, requestedSinglePremium = null) {
    const basic = Math.round(Number(value));
    if (!Number.isFinite(basic) || basic < MIN_BASIC) throw new Error(`GF 投保時基本金額不可低於研究最低門檻 USD ${MIN_BASIC.toLocaleString('en-US')}。`);
    const tier = tierForBasic(basic), premium = premiumForBasic(basic);
    return {
      mode: 'single', value: premium, singlePremium: premium, effectiveSinglePremium: premium,
      singlePremiumCents: Math.round(premium * 100), displaySinglePremium: Math.round(premium),
      requestedSinglePremium: requestedSinglePremium == null ? premium : Number(requestedSinglePremium),
      annual: premium, total: premium, basic,
      discountPerThousand: tier.discount, benefitScale: basic / SOURCE_BASIC,
      premiumScale: premium / SOURCE_BASIC,
      bracket: `${tier.low}-${Number.isFinite(tier.high) ? tier.high : 'up'}`,
    };
  }
  function basicForSinglePremium(value) {
    const requested = Number(value);
    const minimumPremium = Math.ceil(premiumForBasic(MIN_BASIC));
    if (!Number.isFinite(requested) || requested < minimumPremium) throw new Error(`GF 一次性繳費不可低於研究最低門檻 USD ${minimumPremium.toLocaleString('en-US')}。`);
    const candidates = [];
    for (const tier of DISCOUNT_TIERS) {
      const estimate = requested * 1000 / (1000 - tier.discount);
      for (let basic = Math.floor(estimate) - 2; basic <= Math.ceil(estimate) + 2; basic++) {
        if (basic < Math.max(MIN_BASIC, tier.low) || basic > tier.high) continue;
        const input = inputForIntegerBasic(basic, requested);
        candidates.push({ input, difference: Math.abs(input.singlePremium - requested) });
      }
    }
    if (!candidates.length) throw new Error('一次性繳費未能對應有效的大額折扣級距。');
    candidates.sort((a, b) => a.difference - b.difference || b.input.basic - a.input.basic);
    return candidates[0].input;
  }
  function refinedRate(year, component) {
    const observations = OFFICIAL_ANCHORS
      .map(anchor => ({ basic: anchor.basic, value: anchor.values.get(year)?.[component] }))
      .filter(item => Number.isFinite(item.value));
    const low = Math.max(...observations.map(item => (item.value - .5) / item.basic));
    const high = Math.min(...observations.map(item => (item.value + .5) / item.basic));
    if (low <= high) return (low + high) / 2;
    return observations.reduce((sum, item) => sum + item.basic * item.value, 0) /
      observations.reduce((sum, item) => sum + item.basic * item.basic, 0);
  }
  function ratesForYear(year) {
    return { g: refinedRate(year, 'g'), r: refinedRate(year, 'r'), t: refinedRate(year, 't') };
  }
  function validateSingleInputs(issueAge, rows) {
    const issue = Number(issueAge);
    if (!Number.isInteger(issue) || issue < 0 || issue > 98) throw new Error('GF 投保年齡必須是 0 至 98 的整數。');
    if (!Array.isArray(rows) || !rows.length) throw new Error('請輸入至少一年醫療保費。');
    for (const row of rows) {
      const year = mapAgeToPolicyYear(row.age, issue);
      if (year < 1) throw new Error(`${row.age} 歲早於首個完整保單年度。`);
      if (year > MAX_YEAR) throw new Error(`${row.age} 歲對應第 ${year} 保單年度，超出目前單繳官方基準第 ${MAX_YEAR} 年。`);
      if (Number(row.medicalPremium) > 0 && year < 2) throw new Error(`${row.age} 歲對應第 ${year} 保單年度；一次性繳費版本提款最早由第 2 保單年度年末開始。`);
    }
    return true;
  }
  function capBeforeWithdrawal(singlePremium, year, withdrawals) {
    const premiumFutureValue = Number(singlePremium) * Math.pow(CAP_RATE, year);
    const withdrawalFutureValue = withdrawals.reduce((sum, item) => sum + item.amount * Math.pow(CAP_RATE, year - item.year), 0);
    return Math.max(0, premiumFutureValue - withdrawalFutureValue);
  }
  function allocateWithdrawal(requested, preG, preR, preT, q) {
    const available = Math.max(0, preG) + Math.max(0, preR) + Math.max(0, preT);
    const actual = Math.min(Math.max(0, Number(requested) || 0), available);
    if (actual <= 0) return { g: 0, r: 0, t: 0, actual: 0, shortage: Math.max(0, Number(requested) || 0) };
    const terminalForR = Math.min(preT, q * preR), rPool = preR + terminalForR;
    let g = 0, r = 0, t = 0;
    if (actual <= rPool + 1e-9) {
      r = actual / (1 + q); t = actual - r;
    } else {
      const remaining = actual - rPool, terminalForBasic = Math.max(0, preT - terminalForR);
      const qG = preG > 0 ? terminalForBasic / preG : 0;
      g = preG > 0 ? Math.min(preG, remaining / (1 + qG)) : 0;
      r = preR; t = actual - g - r;
    }
    return { g, r, t, actual, shortage: Math.max(0, Number(requested) - actual) };
  }
  function makeRow(input, issueAge, year, requested, actual, fromG, fromR, fromT, basic, postG, postR, postT, cumulative, status, surrendered, force) {
    return {
      case: 'single_state_model', age: Number(issueAge) + year, policy_year: year,
      paid_premium_total: input.singlePremium, requested_withdrawal: requested, withdrawal_total: actual,
      from_guaranteed_cash: fromG, from_reversionary_bonus: fromR, from_terminal_bonus: fromT,
      post_basic_amount: basic, post_guaranteed_cash: postG, post_reversionary_bonus: postR,
      post_terminal_bonus: postT, post_surrender_total: postG + postR + postT,
      cumulative_withdrawal: cumulative, status, surrendered, force,
      source_ref: `${MODEL_VERSION} · ${D.version}`,
    };
  }
  function buildSingleRows(input, issueAge, schedule, minimumBasic = SURVIVAL_BASIC) {
    let basic = input.basic, postR = 0, cumulative = 0, surrendered = false;
    const completedWithdrawals = [], rows = [];
    for (let year = 1; year <= MAX_YEAR; year++) {
      const requested = roundMoney(schedule[year] || 0);
      if (surrendered) {
        rows.push(makeRow(input, issueAge, year, requested, 0, 0, 0, 0, 0, 0, 0, 0, cumulative, '已退保', true, false));
        continue;
      }
      const rates = ratesForYear(year);
      const preG = rates.g * basic;
      const band = R_BANDS.find(item => year >= item.start && year <= item.end);
      const preR = year <= 2 || !band ? 0 : Math.max(0, band.a * postR + band.b * basic);
      const q = rates.t / (rates.r + 1);
      const uncappedT = q * (preR + basic);
      const capTotal = capBeforeWithdrawal(input.singlePremium, year, completedWithdrawals);
      const preT = Math.max(0, Math.min(uncappedT, capTotal - preG - preR));
      let from = allocateWithdrawal(requested, preG, preR, preT, q);
      let postG = Math.max(0, preG - from.g), nextBasic = preG > 0 ? basic * postG / preG : 0;
      let postT = Math.max(0, preT - from.t); postR = Math.max(0, preR - from.r);
      const force = requested > 0 && (from.shortage > .01 || nextBasic < minimumBasic - .01);
      if (force) {
        from = { g: preG, r: preR, t: preT, actual: preG + preR + preT, shortage: 0 };
        postG = 0; postR = 0; postT = 0; nextBasic = 0; surrendered = true;
      }
      basic = nextBasic; cumulative = roundMoney(cumulative + from.actual);
      if (from.actual > 0) completedWithdrawals.push({ year, amount: from.actual });
      const status = force ? '當年退保' : (requested > 0 ? '已提款' : '正常');
      rows.push(makeRow(input, issueAge, year, requested, from.actual, from.g, from.r, from.t, basic, postG, postR, postT, cumulative, status, surrendered, force));
    }
    return rows;
  }
  function simulateSingleInput(input, issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    validateSingleInputs(issueAge, medicalRows);
    input = { ...input, issueAge: Number(issueAge) };
    const schedule = scheduleFromMedicalRows(issueAge, medicalRows);
    const rows = buildSingleRows(input, issueAge, schedule, minimumBasic);
    const surrenderRow = rows.find(row => row.status === '當年退保');
    const result = {
      input, schedule, rows,
      requestedTotal: roundMoney(medicalRows.reduce((sum, row) => sum + Number(row.medicalPremium || 0), 0)),
      actualTotal: rows.length ? Number(rows.at(-1).cumulative_withdrawal) : 0,
      surrendered: Boolean(surrenderRow), surrenderYear: surrenderRow?.policy_year || null,
      baseSource: D.version, baseSourceType: 'official_pgs_expected_scenario', modelVersion: MODEL_VERSION,
      baselineAgeEvidence: '28', issueAgeSupport: Number(issueAge) === 28 ? 'official policy-year baseline' : 'policy-year extrapolation',
      generatedAt: new Date().toISOString(), minimumBasic,
    };
    result.scheduleFingerprint = stableFingerprint(Object.entries(schedule));
    result.resultFingerprint = stableFingerprint(rows.map(row => [row.policy_year,row.requested_withdrawal,row.withdrawal_total,row.post_basic_amount,row.post_guaranteed_cash,row.post_reversionary_bonus,row.post_terminal_bonus,row.status]));
    return result;
  }
  function simulateSinglePremium(singlePremium, issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    return simulateSingleInput(basicForSinglePremium(singlePremium), issueAge, medicalRows, minimumBasic);
  }
  function simulateSingleBasic(basic, issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    return simulateSingleInput(inputForIntegerBasic(basic), issueAge, medicalRows, minimumBasic);
  }
  function canFundSinglePremium(premium, issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    try {
      const result = simulateSinglePremium(premium, issueAge, medicalRows, minimumBasic);
      const lastYear = Math.max(...medicalRows.map(row => mapAgeToPolicyYear(row.age, issueAge)));
      const funded = !result.surrendered && result.rows.filter(row => row.policy_year <= lastYear)
        .every(row => row.withdrawal_total + .01 >= row.requested_withdrawal);
      return { funded, result, reason: funded ? 'funded' : 'insufficient_or_minimum_basic_surrender' };
    } catch (error) { return { funded: false, result: null, reason: error.message }; }
  }
  function canFundSingleBasic(basic, issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    try {
      const result = simulateSingleBasic(basic, issueAge, medicalRows, minimumBasic);
      const lastYear = Math.max(...medicalRows.map(row => mapAgeToPolicyYear(row.age, issueAge)));
      const funded = !result.surrendered && result.rows.filter(row => row.policy_year <= lastYear)
        .every(row => row.withdrawal_total + .01 >= row.requested_withdrawal);
      return { funded, result, reason: funded ? 'funded' : 'insufficient_or_minimum_basic_surrender' };
    } catch (error) { return { funded: false, result: null, reason: error.message }; }
  }
  function findMinimumSinglePremium(issueAge, medicalRows, minimumBasic = SURVIVAL_BASIC) {
    validateSingleInputs(issueAge, medicalRows);
    const tierCandidates = [];
    for (const tier of DISCOUNT_TIERS) {
      let low = Math.max(MIN_BASIC, tier.low), high = Number.isFinite(tier.high) ? tier.high : 100000000;
      if (!canFundSingleBasic(high, issueAge, medicalRows, minimumBasic).funded) continue;
      while (low < high) {
        const mid = Math.floor((low + high) / 2);
        if (canFundSingleBasic(mid, issueAge, medicalRows, minimumBasic).funded) high = mid; else low = mid + 1;
      }
      const pass = canFundSingleBasic(low, issueAge, medicalRows, minimumBasic);
      if (pass.funded) tierCandidates.push({ basic: low, result: pass.result, tier });
    }
    if (!tierCandidates.length) throw new Error('在搜尋上限內仍找不到足夠的一次性繳費，請檢查醫療保費期間。');
    tierCandidates.sort((a, b) => a.result.input.singlePremium - b.result.input.singlePremium || a.basic - b.basic);
    const winner = tierCandidates[0], lowerBasic = winner.basic - 1;
    const lower = lowerBasic >= Math.max(MIN_BASIC, winner.tier.low) ? canFundSingleBasic(lowerBasic, issueAge, medicalRows, minimumBasic) : { funded: false, result: null };
    if (lower.funded) throw new Error('最低一次性繳費的較低基本金額驗證失敗。');
    return {
      singlePremium: winner.result.input.singlePremium, basic: winner.basic, result: winner.result,
      lowerBasic, lowerPremium: lowerBasic >= MIN_BASIC ? premiumForBasic(lowerBasic) : null,
      lowerFails: true,
      tierCandidates: tierCandidates.map(item => ({ basic: item.basic, premium: item.result.input.singlePremium, discount: item.tier.discount })),
    };
  }

  validateMedicalInputs = validateSingleInputs;
  invalidateFinanceResult = function invalidateSingleResult() {
    originalInvalidate();
    const pgs = document.getElementById('pgsTools'); if (pgs) pgs.hidden = true;
    document.getElementById('experimentMethod')?.remove(); document.getElementById('researchNext')?.remove();
  };
  renderFinanceResult = function renderSingleResult(result, proof = null) {
    currentFinanceResult = result; currentMinimumProof = proof;
    const summary = document.getElementById('financeSummary'); clearElement(summary);
    appendMetric(summary, proof ? '最低一次性繳費' : '指定一次性繳費', `USD ${fmt(result.input.singlePremium)}`);
    appendMetric(summary, '投保時基本金額', `USD ${fmt(result.input.basic)}`);
    appendMetric(summary, '大額折扣', `${result.input.discountPerThousand} / 1,000`);
    appendMetric(summary, '醫療保費總額', `USD ${fmt(result.requestedTotal)}`);
    appendMetric(summary, '實際累計提款', `USD ${fmt(result.actualTotal)}`);
    appendMetric(summary, '最終狀態', result.surrendered ? `第 ${result.surrenderYear} 年退保` : '保單持續');
    if (proof) appendMetric(summary, '較低基本金額驗證', `基本金額 USD ${fmt(proof.lowerBasic)}（保費約 USD ${fmt(proof.lowerPremium)}）不足`);
    if (currentMedicalContext?.source === 'table') appendMetric(summary, '醫療保費表', `${currentMedicalContext.formName} · ${currentMedicalContext.deductibleLabel}`);
    const notice = document.getElementById('financeNotice'); notice.className = result.surrendered ? 'notice bad' : 'notice warn';
    const outcome = result.surrendered ? `第 ${result.surrenderYear} 保單年度觸發整份退保。` : proof ? `最低研究值 USD ${fmt(proof.singlePremium)}；較低基本金額 USD ${fmt(proof.lowerBasic)} 未能令保單持續至設定結束年齡。` : '指定一次性繳費情境已完成。';
    notice.textContent = `${outcome} 內部保費保留至小數點後兩位，畫面只顯示四捨五入整數；所有搜尋以整數基本金額進行。最低搜尋使用三份正式提款案例校準的 USD ${SURVIVAL_BASIC.toLocaleString('en-US')} 持續閘門。提款在保單年度年末進行。其他保費、年齡及提款期間仍須再用 PGS 驗證，不可直接作銷售數字。`;
    const table = document.getElementById('financeTable'); clearElement(table);
    const components = financeComponentMode() === 'components';
    const headers = ['年齡','保單年度','實際提款醫療保費','提款後基本金額'];
    if (components) headers.push('保證現金價值','復歸紅利','終期紅利');
    headers.push('提款後總額','狀態','累計提款');
    const thead = document.createElement('thead'), hr = document.createElement('tr'); headers.forEach(v => appendCell(hr,v,'th')); thead.appendChild(hr); table.appendChild(thead);
    const tbody = document.createElement('tbody');
    result.rows.forEach(item => {
      const tr = document.createElement('tr'); if (item.status.includes('退保')) tr.className='force-row'; else if (item.requested_withdrawal>0) tr.className='active-row';
      appendCell(tr,item.age); appendCell(tr,item.policy_year); appendCell(tr,fmt(item.withdrawal_total)); appendCell(tr,fmt(item.post_basic_amount));
      if (components) { appendCell(tr,fmt(item.post_guaranteed_cash)); appendCell(tr,fmt(item.post_reversionary_bonus)); appendCell(tr,fmt(item.post_terminal_bonus)); }
      appendCell(tr,fmt(item.post_surrender_total)); appendCell(tr,item.status,'td',item.status.includes('退保')?'status-surrendered':'status-active'); appendCell(tr,fmt(item.cumulative_withdrawal),'td','cumulative-cell'); tbody.appendChild(tr);
    });
    table.appendChild(tbody); document.getElementById('financeDownload').disabled=false;
    const method=document.createElement('div'); method.id='experimentMethod'; method.className='experiment-method';
    method.innerHTML='<b>一次性繳費三基準狀態模型</b>同時使用三份正式單繳版無提款基準的整數區間、復歸紅利三段累積、6.5% IRR 上限、R/T 優先提款、G/T 後提款及經正式最低案例校準的持續閘門。'; summary.before(method);
  };
  financeCsvText = function singleCsv(result, medicalContext) {
    if (!result) throw new Error('沒有可匯出的單繳研究結果。');
    const active=Object.entries(result.schedule).filter(([,a])=>Number(a)>0).map(([y])=>Number(y));
    const meta=[['schema_version','single-4'],['model_status','research_not_official'],['model_version',MODEL_VERSION],['official_anchor_version',D.version],['official_anchor_count',OFFICIAL_ANCHORS.length],['calculation_mode',result.calculationMode],['generated_at_utc',result.generatedAt],['issue_age',result.input.issueAge],['requested_single_premium_display',result.input.requestedSinglePremium],['effective_single_premium_exact',result.input.singlePremium.toFixed(2)],['effective_single_premium_cents',result.input.singlePremiumCents],['single_premium_display_rounded',result.input.displaySinglePremium],['initial_basic_amount_integer',result.input.basic],['premium_precision','internal_cents_ui_nearest_integer'],['search_axis','integer_basic_amount'],['official_initial_minimum_basic',MIN_BASIC],['calibrated_survival_basic',SURVIVAL_BASIC],['discount_per_thousand',result.input.discountPerThousand],['minimum_basic_selected',result.minimumBasic],['minimum_basic_inferred_bracket',MIN_BASIC_BRACKET],['withdrawal_timing','policy_year_end'],['schedule_first_policy_year',active.length?Math.min(...active):''],['schedule_last_policy_year',active.length?Math.max(...active):''],['surrender_year',result.surrenderYear||''],['lower_basic_tested',currentMinimumProof?.lowerBasic??''],['lower_premium_exact',currentMinimumProof?.lowerPremium?.toFixed?.(2)??''],['lower_basic_fails',currentMinimumProof?'true':''],['schedule_fingerprint',result.scheduleFingerprint],['result_fingerprint',result.resultFingerprint]];
    if(medicalContext?.source==='table')meta.push(['medical_plan',medicalContext.planName],['policy_form',medicalContext.formName],['deductible',medicalContext.deductibleLabel],['rate_effective_date',medicalContext.effectiveDate]);
    const headers=['age','policy_year','medical_premium','actual_withdrawal','from_guaranteed','from_reversionary','from_terminal','post_basic_amount','guaranteed_cash_value','reversionary_bonus','terminal_bonus','total_value','status','cumulative_withdrawal'];
    const rows=result.rows.map(r=>[r.age,r.policy_year,roundMoney(r.requested_withdrawal),roundMoney(r.withdrawal_total),roundMoney(r.from_guaranteed_cash),roundMoney(r.from_reversionary_bonus),roundMoney(r.from_terminal_bonus),roundMoney(r.post_basic_amount),roundMoney(r.post_guaranteed_cash),roundMoney(r.post_reversionary_bonus),roundMoney(r.post_terminal_bonus),roundMoney(r.post_surrender_total),r.status,roundMoney(r.cumulative_withdrawal)]);
    return '\uFEFF'+meta.concat([[]],[headers],rows).map(row=>row.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\r\n');
  };
  function runSingleResearch() {
    invalidateFinanceResult(); const error=document.getElementById('financeError'); error.hidden=true; error.textContent='';
    try {
      const {issueAge,medicalRows,context}=readFinanceForm(); currentMedicalContext=context;
      if(document.getElementById('financeMode').value==='minimum'){
        const proof=findMinimumSinglePremium(issueAge,medicalRows); proof.result.calculationMode='minimum_single'; renderFinanceResult(proof.result,proof);
      } else {
        const premium=Number(document.getElementById('fixedAnnualPremium').value); if(!Number.isInteger(premium))throw new Error('一次性繳費請輸入整數美元。');
        const result=simulateSinglePremium(premium,issueAge,medicalRows); result.calculationMode='fixed_single'; renderFinanceResult(result,null);
      }
    } catch(value){ invalidateFinanceResult(); error.textContent=value instanceof Error?value.message:String(value); error.hidden=false; }
  }
  function configureUi(){
    document.title='GF 一次性繳費醫療融資研究'; document.querySelector('.top h1').textContent='GF 一次性繳費醫療融資研究';
    document.querySelector('.top p').textContent='由現時年齡、醫療保費開始年齡至結束年齡，反推令保單持續到結束年齡的最低一次性繳費，或測試指定一次性繳費。';
    document.querySelector('.badge').textContent='PGS v27.0 單繳個案校準 · 研究版';
    const banner=document.createElement('div'); banner.className='experiment-banner'; banner.innerHTML='研究範圍提示<small>已改用真正的單繳版逐年現金價值及提款規則，不再以「單繳 ÷ 5」代替。不同金額、年齡及期間仍要由你用 PGS 再測試校準。</small>'; document.querySelector('.wrap').before(banner);
    const mode=document.getElementById('financeMode'); mode.innerHTML='<option value="minimum">反推最低 GF 一次性繳費</option><option value="fixed">指定 GF 一次性繳費</option>'; mode.value='minimum'; mode.disabled=false;
    document.getElementById('fixedPremiumWrap').hidden=true; document.querySelector('label[for="fixedAnnualPremium"]').textContent='GF 一次性繳費（USD）';
    const premium=document.getElementById('fixedAnnualPremium'); premium.value='30000'; premium.min=String(Math.ceil(premiumForBasic(MIN_BASIC))); premium.step='1';
    document.querySelector('label[for="issueAge"]').textContent='現時／GF 投保年齡'; document.querySelector('label[for="medicalStartAge"]').textContent='醫療保費開始年齡'; document.querySelector('label[for="medicalEndAge"]').textContent='醫療保費結束年齡';
    document.getElementById('issueAge').value='28'; document.getElementById('medicalStartAge').value='30'; document.getElementById('medicalEndAge').value='99'; document.getElementById('medicalDeductible').value='18000';
    document.querySelector('main.panel h2').textContent='一次性繳費反推結果'; document.getElementById('financeNotice').textContent='設定三個年齡及醫療保費後按「開始計算」。單繳版本最早由第 2 保單年度年末提款。'; document.getElementById('financeDownload').textContent='下載研究 CSV';
    const old=document.getElementById('financeRun'),run=old.cloneNode(true);old.replaceWith(run);run.addEventListener('click',runSingleResearch);
    mode.addEventListener('change',()=>{document.getElementById('fixedPremiumWrap').hidden=mode.value!=='fixed';invalidateFinanceResult();});
    ['financeMode','fixedAnnualPremium','issueAge','medicalStartAge','medicalEndAge','medicalPremiumSource','medicalPlan','medicalPolicyForm','medicalDeductible','medicalPremiumTable'].forEach(id=>document.getElementById(id)?.addEventListener(id==='fixedAnnualPremium'||id==='issueAge'||id==='medicalPremiumTable'?'input':'change',invalidateFinanceResult));
    invalidateFinanceResult(); updateMedicalPremiumSourceUi();
  }
  function runSingleSelfTests(){
    const output=document.createElement('pre');output.id='singleSelfTestResults';output.className='self-test';document.body.prepend(output);
    const tests=[];function assert(value,message){if(!value)throw new Error(message)}function test(name,fn){try{fn();tests.push(['PASS',name])}catch(error){tests.push(['FAIL',name,error.message])}}
    test('官方單繳基準共有72年',()=>assert(D.baseline.length===72,'baseline'));
    test('三份官方無提款基準已載入',()=>assert(OFFICIAL_ANCHORS.length===3&&D.tertiaryAnchor.baseline.length===68,'anchors'));
    test('預設為最低單繳及固定模式並存',()=>{assert(document.getElementById('financeMode').value==='minimum','mode');assert([...document.getElementById('financeMode').options].map(o=>o.value).join(',')==='minimum,fixed','options')});
    test('預設年齡28/30/99',()=>assert([issueAge.value,medicalStartAge.value,medicalEndAge.value].join('/')==='28/30/99','ages'));
    test('30,000單繳對應30,000整數基本金額',()=>assert(basicForSinglePremium(30000).basic===30000,'basic'));
    test('61,945畫面值映射官方62,571基本金額',()=>{const input=basicForSinglePremium(61945);assert(input.basic===62571,'basic='+input.basic);assert(input.singlePremium===61945.29,'premium='+input.singlePremium);assert(input.displaySinglePremium===61945,'display')});
    test('62,571基本金額保留61,945.29內部保費',()=>{const input=inputForIntegerBasic(62571);assert(input.singlePremiumCents===6194529,'cents');assert(input.singlePremium===61945.29,'premium')});
    test('第1年提款拒絕及第2年接受',()=>{let failed=false;try{validateSingleInputs(28,[{age:29,medicalPremium:1}])}catch(_){failed=true}assert(failed,'Y1 accepted');validateSingleInputs(28,[{age:30,medicalPremium:1}])});
    const medical=D.withdrawals.filter(row=>row[4]>0&&row[0]<=67).map(row=>({age:28+row[0],medicalPremium:row[0]===67?D.plannedWithdrawalYear67:row[4]}));
    const official=simulateSinglePremium(30000,28,medical);
    test('官方個案Y1-Y66基本金額顯示全部吻合',()=>{let max=0;for(let y=1;y<=66;y++)max=Math.max(max,Math.abs(Math.round(official.rows[y-1].post_basic_amount)-D.post[y-1][1]));assert(max===0,'max='+max)});
    test('官方個案Y1-Y66總額最大顯示誤差不超過USD3',()=>{let max=0;for(let y=1;y<=66;y++)max=Math.max(max,Math.abs(Math.round(official.rows[y-1].post_surrender_total)-D.post[y-1][5]));assert(max<=3,'max='+max)});
    test('官方個案Y67整份退保',()=>{const row=official.rows[66];assert(row.status==='當年退保',row.status);assert(row.post_surrender_total===0,'post');assert(Math.abs(row.withdrawal_total-392257)<=8,'actual='+row.withdrawal_total)});
    const displayedMaxError=(result,post,field,index)=>post.reduce((max,row)=>Math.max(max,Math.abs(Math.round(result.rows[row[0]-1][field])-row[index])),0);
    const secondaryMedical=medicalRowsFromTable('aia_avsw','basic','8800',45,99);
    const secondaryProof=findMinimumSinglePremium(43,secondaryMedical);
    test('第二正式最低案例反推62,571及61,945.29',()=>{assert(secondaryProof.basic===62571,'basic='+secondaryProof.basic);assert(secondaryProof.singlePremium===61945.29,'premium='+secondaryProof.singlePremium)});
    test('第二正式案例總額最大顯示誤差不超過USD1',()=>assert(displayedMaxError(secondaryProof.result,D.secondaryAnchor.post,'post_surrender_total',5)<=1,'secondary total'));
    const tertiaryMedical=medicalRowsFromTable('aia_avsw','basic','8800',40,99);
    const tertiaryProof=findMinimumSinglePremium(32,tertiaryMedical);
    test('第三正式最低案例反推35,917及35,557.83',()=>{assert(tertiaryProof.basic===35917,'basic='+tertiaryProof.basic);assert(tertiaryProof.singlePremium===35557.83,'premium='+tertiaryProof.singlePremium)});
    test('第三正式案例基本金額68年全部吻合',()=>assert(displayedMaxError(tertiaryProof.result,D.tertiaryAnchor.post,'post_basic_amount',1)===0,'tertiary basic'));
    test('第三正式案例總額最大顯示誤差不超過USD3',()=>assert(displayedMaxError(tertiaryProof.result,D.tertiaryAnchor.post,'post_surrender_total',5)<=3,'tertiary total'));
    test('CSV保留第三案例精確保費及三基準追溯資料',()=>{const csv=financeCsvText(tertiaryProof.result,null);assert(csv.includes('"effective_single_premium_exact","35557.83"'),'premium');assert(csv.includes('"official_anchor_count","3"'),'anchors');assert(csv.includes('"calibrated_survival_basic","7501"'),'survival')});
    test('最低搜尋以整數基本金額驗證',()=>{const rows=[{age:30,medicalPremium:639},{age:31,medicalPremium:645}],proof=findMinimumSinglePremium(28,rows);assert(canFundSingleBasic(proof.basic,28,rows).funded,'pass');assert(!canFundSingleBasic(proof.lowerBasic,28,rows).funded,'lower');assert(Number.isInteger(proof.basic),'integer')});
    test('初始CSV及PGS不可用',()=>{assert(document.getElementById('financeDownload').disabled,'CSV');assert(!document.getElementById('pgsTools')||document.getElementById('pgsTools').hidden,'PGS')});
    const failures=tests.filter(row=>row[0]==='FAIL');output.textContent=tests.map(row=>row.join(' ')).join('\n')+`\n\n${tests.length-failures.length}/${tests.length} passed`;output.dataset.failed=String(failures.length);
  }
  configureUi();
  window.GFSingleExperiment={MODEL_VERSION,premiumForBasic,inputForIntegerBasic,basicForSinglePremium,validateMedicalInputs:validateSingleInputs,simulateSinglePremium,simulateSingleBasic,canFundSinglePremium,canFundSingleBasic,findMinimumSinglePremium,minimumBasic:MIN_BASIC,survivalBasic:SURVIVAL_BASIC,minimumBasicBracket:MIN_BASIC_BRACKET};
  if(new URLSearchParams(location.search).get('selftest')==='1')runSingleSelfTests();
})();
