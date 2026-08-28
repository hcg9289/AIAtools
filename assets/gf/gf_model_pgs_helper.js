(() => {
  'use strict';

  const EXPORT_FORMAT = 'AIA_GF_PGS_WITHDRAWALS';
  const EXPORT_VERSION = 1;
  const POLICY_YEARS = 72;
  const MAX_WITHDRAWAL = 99999999999;
  const BOOKMARKLET_CODE = 'javascript:(async()=>{const F="AIA_GF_PGS_WITHDRAWALS",V=1,M=99999999999;function P(t){let p;try{p=JSON.parse(String(t).trim())}catch(e){throw Error("資料不是有效的 AIAtools PGS 格式")};if(!p||p.format!==F||p.version!==V||!Array.isArray(p.withdrawals)||p.withdrawals.length!==72)throw Error("資料版本不符或年度不完整");const seen=new Set(),rows=[];for(const r of p.withdrawals){if(!Array.isArray(r)||r.length!==2)throw Error("提款資料結構不正確");const y=r[0],a=r[1];if(!Number.isInteger(y)||y<1||y>72)throw Error("保單年度超出 1 至 72");if(seen.has(y))throw Error("保單年度重複："+y);if(!Number.isInteger(a)||a<0||a>M)throw Error("第 "+y+" 年提款金額不合法");seen.add(y);rows.push([y,a])}for(let y=1;y<=72;y++)if(!seen.has(y))throw Error("欠缺第 "+y+" 保單年度");return rows}let t="";try{t=await navigator.clipboard.readText()}catch(e){}if(!t)t=prompt("瀏覽器未能自動讀取剪貼簿。請貼上由 AIAtools 複製的 PGS 提款資料：");if(t===null)return;try{const rows=P(t),d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,"value");if(!d||!d.set)throw Error("瀏覽器不支援原生輸入更新");let filled=0,disabled=0,missing=0;for(const [y,a] of rows){const el=document.querySelector("input[name=\\"withDraw-"+(y-1)+"\\"]");if(!el){missing++;continue}if(el.disabled){disabled++;continue}d.set.call(el,String(a));el.dispatchEvent(new Event("input",{bubbles:true}));el.dispatchEvent(new Event("change",{bubbles:true}));filled++}alert("PGS 已填入 "+filled+" 格；跳過鎖定 "+disabled+" 格；找不到 "+missing+" 格。請逐年核對後自行按「計算及儲存」。")}catch(e){alert("PGS 一鍵填入失敗："+e.message)}})();';

  function validatePayload(input) {
    let payload;
    try {
      payload = typeof input === 'string' ? JSON.parse(input.trim()) : input;
    } catch (_) {
      throw new Error('資料不是有效的 AIAtools PGS 格式。');
    }
    if (!payload || payload.format !== EXPORT_FORMAT || payload.version !== EXPORT_VERSION) {
      throw new Error('資料不是本工具產生，或版本不受支援。');
    }
    if (!Array.isArray(payload.withdrawals) || payload.withdrawals.length !== POLICY_YEARS) {
      throw new Error('PGS 提款資料必須完整包含第 1–72 保單年度。');
    }
    const seen = new Set();
    const rows = [];
    payload.withdrawals.forEach((item) => {
      if (!Array.isArray(item) || item.length !== 2) throw new Error('PGS 提款資料結構不正確。');
      const [policyYear, amount] = item;
      if (!Number.isInteger(policyYear) || policyYear < 1 || policyYear > POLICY_YEARS) {
        throw new Error('保單年度必須是 1 至 72 的整數。');
      }
      if (seen.has(policyYear)) throw new Error(`保單年度重複：${policyYear}。`);
      if (!Number.isInteger(amount) || amount < 0 || amount > MAX_WITHDRAWAL) {
        throw new Error(`第 ${policyYear} 年提款金額必須是 0 至 ${MAX_WITHDRAWAL} 的整數。`);
      }
      seen.add(policyYear);
      rows.push([policyYear, amount]);
    });
    for (let policyYear = 1; policyYear <= POLICY_YEARS; policyYear += 1) {
      if (!seen.has(policyYear)) throw new Error(`欠缺第 ${policyYear} 保單年度。`);
    }
    return rows.sort((a, b) => a[0] - b[0]);
  }

  function buildPayload(result) {
    if (!result || !result.schedule) throw new Error('目前沒有可複製的反推結果，請先按「開始計算」。');
    const withdrawals = [];
    for (let policyYear = 1; policyYear <= POLICY_YEARS; policyYear += 1) {
      const raw = Number(result.schedule[policyYear] ?? 0);
      if (!Number.isFinite(raw) || raw < 0 || raw > MAX_WITHDRAWAL) {
        throw new Error(`第 ${policyYear} 年原定提款金額超出 PGS 可接受範圍。`);
      }
      withdrawals.push([policyYear, Math.round(raw)]);
    }
    const payload = {
      format: EXPORT_FORMAT,
      version: EXPORT_VERSION,
      timing: 'policy_year_end',
      source: 'requested_withdrawal_schedule',
      withdrawals,
    };
    validatePayload(payload);
    return JSON.stringify(payload);
  }

  function setStatus(message, isError = false) {
    const status = document.getElementById('pgsStatus');
    if (!status) return;
    status.textContent = message;
    status.style.color = isError ? 'var(--red)' : '#285449';
  }

  function setReady(ready) {
    const copyButton = document.getElementById('pgsCopy');
    if (copyButton) copyButton.disabled = !ready;
    setStatus(ready ? '結果已準備好；可複製原定提款排程到 PGS。' : '輸入已變更，請重新計算後再複製。');
  }

  async function copyData() {
    try {
      const result = typeof currentFinanceResult === 'undefined' ? null : currentFinanceResult;
      const text = buildPayload(result);
      try {
        if (!navigator.clipboard?.writeText) throw new Error('clipboard unavailable');
        await navigator.clipboard.writeText(text);
        setStatus('已複製第 1–72 保單年度的 PGS 提款資料。');
      } catch (_) {
        prompt('瀏覽器未能自動複製。請按 Command + C 或 Ctrl + C 複製以下資料：', text);
        setStatus('瀏覽器已顯示資料，請確認已手動複製。');
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  function installUi() {
    if (document.getElementById('pgsTools')) return;
    const table = document.getElementById('financeTable');
    if (!table) return;

    const style = document.createElement('style');
    style.textContent = '.pgs-tools{margin-top:16px;padding:16px;border:1px solid #b9d4c7;border-radius:10px;background:#eef8f3}.pgs-tools h3{margin:0 0 8px}.pgs-tools ol{margin:12px 0 0;padding-left:22px}.pgs-tools li+li{margin-top:5px}.pgs-bookmark-setup{margin-top:14px;padding:14px;border:2px solid #bd7b18;border-radius:10px;background:#fffaf0}.pgs-bookmark-setup h4{margin:0 0 8px;color:#6b4815}.pgs-drag-warning{margin:0 0 12px;font-weight:700;color:#9b341c}.pgs-bookmarklet{display:inline-block;padding:11px 18px;border-radius:9px;background:#176b57;color:#fff!important;font-weight:800;text-decoration:none;cursor:grab;box-shadow:0 2px 5px rgba(0,0,0,.16)}.pgs-bookmarklet:active{cursor:grabbing}.pgs-bookmark-help{margin:10px 0 0;color:#5b5142;font-size:13px}.pgs-status{min-height:1.55em;margin-top:9px;font-size:13px;color:#285449}@media(max-width:620px){.pgs-tools .toolbar>*{width:100%;text-align:center}.pgs-bookmarklet{display:block;text-align:center}}';
    document.head.appendChild(style);

    const section = document.createElement('section');
    section.className = 'pgs-tools';
    section.id = 'pgsTools';
    section.setAttribute('aria-labelledby', 'pgsToolsTitle');
    section.innerHTML = '<h3 id="pgsToolsTitle">填入 AIA PGS 提款設定</h3><div class="hint">工具會複製第 1–72 保單年度的原定提款排程；不會把年齡當作保單年度，也不會代你按 PGS 的「計算及儲存」。</div><div class="toolbar"><button class="primary" id="pgsCopy" type="button" disabled>複製 PGS 提款資料</button></div><div class="pgs-bookmark-setup"><h4>首次設定：把一鍵填入工具加入書籤列</h4><p class="pgs-drag-warning">請勿直接點擊下方按鈕。請用滑鼠按住它，拖到瀏覽器最上方的書籤列，再放開滑鼠。</p><a class="pgs-bookmarklet" id="pgsBookmarklet" draggable="true" href="#" title="按住並拖到瀏覽器書籤列">PGS 一鍵填入｜按住拖到書籤列</a><p class="pgs-bookmark-help">看不到書籤列？Chrome／Edge：Mac 按 Command + Shift + B；Windows 按 Ctrl + Shift + B。這項設定每部電腦只需完成一次。</p></div><ol><li>先把上方「PGS 一鍵填入」拖到瀏覽器書籤列。</li><li>每次在本頁完成計算後，按「複製 PGS 提款資料」。</li><li>轉到 PGS「進階設定」頁，再點書籤列內的「PGS 一鍵填入」。</li><li>核對逐年金額後，由你親自按 PGS 的「計算及儲存」。</li></ol><div class="pgs-status" id="pgsStatus" role="status">請先完成反推計算。</div>';
    table.closest('.tableWrap')?.after(section);

    const copyButton = document.getElementById('pgsCopy');
    const bookmarklet = document.getElementById('pgsBookmarklet');
    copyButton?.addEventListener('click', copyData);
    bookmarklet?.setAttribute('href', BOOKMARKLET_CODE);
    bookmarklet?.addEventListener('click', (event) => {
      event.preventDefault();
      window.alert('這個按鈕不是直接點擊使用。請用滑鼠按住它，拖到瀏覽器最上方的書籤列，再放開滑鼠。加入後，才在 PGS 頁面點擊書籤列內的「PGS 一鍵填入」。');
    });
  }

  installUi();

  if (typeof invalidateFinanceResult === 'function') {
    const originalInvalidate = invalidateFinanceResult;
    invalidateFinanceResult = function wrappedInvalidateFinanceResult(...args) {
      const value = originalInvalidate.apply(this, args);
      setReady(false);
      return value;
    };
  }
  if (typeof renderFinanceResult === 'function') {
    const originalRender = renderFinanceResult;
    renderFinanceResult = function wrappedRenderFinanceResult(...args) {
      const value = originalRender.apply(this, args);
      setReady(true);
      return value;
    };
  }

  window.GFPgsHelper = {
    BOOKMARKLET_CODE,
    buildPayload,
    validatePayload,
  };
})();
