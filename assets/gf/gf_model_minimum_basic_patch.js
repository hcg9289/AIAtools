(() => {
  'use strict';

  const MODEL_PATCH_VERSION = 'gf-minimum-basic-v1-2026-08-29';
  const MINIMUM_BASIC_BANDS = Object.freeze([
    Object.freeze({ start: 1, end: 4, amount: 20000 }),
    Object.freeze({ start: 5, end: 14, amount: 10000 }),
    Object.freeze({ start: 15, end: 49, amount: 1500 }),
    Object.freeze({ start: 50, end: 100, amount: 1000 }),
  ]);

  const originalBuildStateTransitionRows = window.buildStateTransitionRows;
  const originalSimulateMedicalFinancing = window.simulateMedicalFinancing;
  const originalCanFundMedicalSchedule = window.canFundMedicalSchedule;
  const originalFindMinimumAnnualPremium = window.findMinimumAnnualPremium;
  const originalFinanceCsvText = window.financeCsvText;

  if (
    typeof originalBuildStateTransitionRows !== 'function'
    || typeof originalSimulateMedicalFinancing !== 'function'
    || typeof originalCanFundMedicalSchedule !== 'function'
    || typeof originalFindMinimumAnnualPremium !== 'function'
    || typeof originalFinanceCsvText !== 'function'
  ) {
    throw new Error('GF 最低基本金額補丁無法連接原有計算引擎。');
  }

  function number(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function roundMoney(value) {
    return Math.round((number(value) + Number.EPSILON) * 100) / 100;
  }

  function minimumBasicForPolicyYear(policyYear) {
    const year = number(policyYear);
    const band = MINIMUM_BASIC_BANDS.find((item) => year >= item.start && year <= item.end);
    if (!band) throw new Error(`最低基本金額規則欠缺第 ${policyYear} 保單年度。`);
    return band.amount;
  }

  function fnv1a32(text) {
    let hash = 0x811c9dc5;
    for (let index = 0; index < text.length; index += 1) {
      hash ^= text.charCodeAt(index);
      hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    return `fnv1a32:${hash.toString(16).padStart(8, '0')}`;
  }

  function tracedModelVersion(baseVersion) {
    const base = String(baseVersion || 'unknown');
    return base.includes(MODEL_PATCH_VERSION) ? base : `${base}+${MODEL_PATCH_VERSION}`;
  }

  function fullSurrenderAtMinimumBasic(rows, triggerIndex) {
    const trigger = rows[triggerIndex];
    const previousCumulative = triggerIndex > 0 ? number(rows[triggerIndex - 1].cumulative_withdrawal) : 0;
    const preGuaranteed = number(trigger.from_guaranteed_cash) + number(trigger.post_guaranteed_cash);
    const preReversionary = number(trigger.from_reversionary_bonus) + number(trigger.post_reversionary_bonus);
    const preTerminal = number(trigger.from_terminal_bonus) + number(trigger.post_terminal_bonus);
    const preValue = roundMoney(preGuaranteed + preReversionary + preTerminal);
    const cumulative = roundMoney(previousCumulative + preValue);

    rows[triggerIndex] = {
      ...trigger,
      withdrawal_total: preValue,
      from_guaranteed_cash: preGuaranteed,
      from_reversionary_bonus: preReversionary,
      from_terminal_bonus: preTerminal,
      post_basic_amount: 0,
      post_guaranteed_cash: 0,
      post_reversionary_bonus: 0,
      post_terminal_bonus: 0,
      post_surrender_total: 0,
      cumulative_withdrawal: cumulative,
      status: '當年退保',
      surrendered: true,
      force: true,
      minimum_basic_surrender: true,
      minimum_basic_required: minimumBasicForPolicyYear(trigger.policy_year),
      source_ref: `${trigger.source_ref || ''} · ${MODEL_PATCH_VERSION}`.replace(/^ · /, ''),
    };

    for (let index = triggerIndex + 1; index < rows.length; index += 1) {
      rows[index] = {
        ...rows[index],
        withdrawal_total: 0,
        from_guaranteed_cash: 0,
        from_reversionary_bonus: 0,
        from_terminal_bonus: 0,
        post_basic_amount: 0,
        post_guaranteed_cash: 0,
        post_reversionary_bonus: 0,
        post_terminal_bonus: 0,
        post_surrender_total: 0,
        cumulative_withdrawal: cumulative,
        status: '已退保',
        surrendered: true,
        force: false,
        minimum_basic_surrender: false,
        source_ref: `${rows[index].source_ref || ''} · ${MODEL_PATCH_VERSION}`.replace(/^ · /, ''),
      };
    }
  }

  function patchedBuildStateTransitionRows(...args) {
    const state = originalBuildStateTransitionRows.apply(this, args);
    if (!state || !Array.isArray(state.rows)) return state;

    const rows = state.rows.map((row) => ({ ...row }));
    const triggerIndex = rows.findIndex((row) => (
      number(row.requested_withdrawal) > 0
      && !row.surrendered
      && number(row.post_basic_amount) + 1e-9 < minimumBasicForPolicyYear(row.policy_year)
    ));
    if (triggerIndex < 0) return { ...state, rows };

    fullSurrenderAtMinimumBasic(rows, triggerIndex);
    return {
      ...state,
      rows,
      surrendered: true,
      surrenderYear: number(rows[triggerIndex].policy_year),
      minimumBasicSurrender: true,
      modelPatchVersion: MODEL_PATCH_VERSION,
    };
  }

  function patchedSimulateMedicalFinancing(...args) {
    const result = originalSimulateMedicalFinancing.apply(this, args);
    if (!result) return result;
    result.baseModelVersion = result.baseModelVersion || result.modelVersion;
    result.modelVersion = tracedModelVersion(result.baseModelVersion);
    result.modelPatchVersion = MODEL_PATCH_VERSION;
    result.minimumBasicBands = MINIMUM_BASIC_BANDS.map((band) => ({ ...band }));
    result.minimumBasicSurrender = Boolean(result.rows?.some((row) => row.minimum_basic_surrender));
    return result;
  }

  function patchedCanFundMedicalSchedule(...args) {
    const check = originalCanFundMedicalSchedule.apply(this, args);
    if (check?.result?.minimumBasicSurrender) return { ...check, funded: false, reason: 'minimum_basic_surrender' };
    return check;
  }

  function patchedFindMinimumAnnualPremium(...args) {
    return originalFindMinimumAnnualPremium.apply(this, args);
  }

  function surrenderReason(row) {
    if (row?.minimum_basic_surrender) return 'minimum_basic';
    if (row?.status === '當年退保') return 'insufficient_value';
    if (row?.status === '已退保') return 'post_surrender';
    return '';
  }

  function patchedFinanceCsvText(result, ...args) {
    const csv = originalFinanceCsvText.call(this, result, ...args);
    if (typeof csv !== 'string' || !result || !Array.isArray(result.rows)) return csv;
    const newline = csv.includes('\r\n') ? '\r\n' : '\n';
    const hasBom = csv.startsWith('\uFEFF');
    const content = hasBom ? csv.slice(1) : csv;
    const lines = content.split(/\r?\n/);
    const headerIndex = lines.findIndex((line) => line.startsWith('"age","policy_year",'));
    if (headerIndex < 0) throw new Error('GF CSV欠缺逐年表頭，最低基本金額補丁停止匯出。');

    const oldFingerprintIndex = lines.findIndex((line) => line.startsWith('"engine_parameter_fingerprint",'));
    const oldFingerprint = oldFingerprintIndex >= 0
      ? (lines[oldFingerprintIndex].match(/^"engine_parameter_fingerprint","([^"]*)"/)?.[1] || '')
      : '';
    const combinedFingerprint = fnv1a32(JSON.stringify({
      base: oldFingerprint,
      patch: MODEL_PATCH_VERSION,
      minimumBasicBands: MINIMUM_BASIC_BANDS,
    }));
    if (oldFingerprintIndex >= 0) {
      lines[oldFingerprintIndex] = `"engine_parameter_fingerprint","${combinedFingerprint}"`;
    }

    const metadata = [
      `"model_patch_version","${MODEL_PATCH_VERSION}"`,
      `"minimum_basic_bands","1-4:20000;5-14:10000;15-49:1500;50-100:1000"`,
    ];
    const metadataIndex = headerIndex > 0 && lines[headerIndex - 1] === '' ? headerIndex - 1 : headerIndex;
    lines.splice(metadataIndex, 0, ...metadata);
    const shiftedHeaderIndex = headerIndex + metadata.length;
    lines[shiftedHeaderIndex] += ',"post_basic_amount","surrender_reason"';
    result.rows.forEach((row, index) => {
      const lineIndex = shiftedHeaderIndex + 1 + index;
      if (lineIndex >= lines.length || !lines[lineIndex]) {
        throw new Error(`GF CSV欠缺第 ${row.policy_year} 保單年度資料。`);
      }
      lines[lineIndex] += `,"${roundMoney(row.post_basic_amount)}","${surrenderReason(row)}"`;
    });
    return `${hasBom ? '\uFEFF' : ''}${lines.join(newline)}`;
  }

  window.buildStateTransitionRows = patchedBuildStateTransitionRows;
  window.simulateMedicalFinancing = patchedSimulateMedicalFinancing;
  window.canFundMedicalSchedule = patchedCanFundMedicalSchedule;
  window.findMinimumAnnualPremium = patchedFindMinimumAnnualPremium;
  window.financeCsvText = patchedFinanceCsvText;

  if (window.GFFinanceCore) {
    Object.assign(window.GFFinanceCore, {
      buildStateTransitionRows: patchedBuildStateTransitionRows,
      simulateMedicalFinancing: patchedSimulateMedicalFinancing,
      canFundMedicalSchedule: patchedCanFundMedicalSchedule,
      findMinimumAnnualPremium: patchedFindMinimumAnnualPremium,
      financeCsvText: patchedFinanceCsvText,
      minimumBasicForPolicyYear,
    });
  }

  window.GFMinimumBasicPatch = Object.freeze({
    MODEL_PATCH_VERSION,
    MINIMUM_BASIC_BANDS,
    minimumBasicForPolicyYear,
  });
})();
