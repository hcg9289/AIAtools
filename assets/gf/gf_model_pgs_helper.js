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

  function revealInstallGuide() {
    const panel = document.getElementById('pgsInstallPanel');
    const code = document.getElementById('pgsBookmarkletCode');
    if (panel) panel.hidden = false;
    if (code) code.value = BOOKMARKLET_CODE;
  }

  async function copyBookmarkletCode() {
    const code = document.getElementById('pgsBookmarkletCode');
    if (!code) return false;
    try {
      if (!navigator.clipboard?.writeText) throw new Error('clipboard unavailable');
      await navigator.clipboard.writeText(BOOKMARKLET_CODE);
      setStatus('書籤程式已複製，但尚未加入瀏覽器；請完成下方步驟。');
      return true;
    } catch (_) {
      code.focus();
      code.select();
      setStatus('瀏覽器未能自動複製。書籤程式已選取，請按 Command + C 或 Ctrl + C；此時仍未加入書籤。', true);
      return false;
    }
  }

  async function requestBookmarkInstall() {
    const confirmed = window.confirm('基於 Chrome／Safari 的安全限制，普通網頁不能直接新增或修改你的書籤。按「確定」後，本頁會複製書籤程式並顯示一次性的安裝步驟；這不代表書籤已經加入。');
    if (!confirmed) return false;
    revealInstallGuide();
    await copyBookmarkletCode();
    return true;
  }

  function installUi() {
    if (document.getElementById('pgsTools')) return;
    const table = document.getElementById('financeTable');
    if (!table) return;

    const style = document.createElement('style');
    style.textContent = '.pgs-tools{margin-top:16px;padding:16px;border:1px solid #b9d4c7;border-radius:10px;background:#eef8f3}.pgs-tools h3{margin:0 0 8px}.pgs-tools ol{margin:12px 0 0;padding-left:22px}.pgs-tools li+li{margin-top:5px}.pgs-install-panel{margin-top:14px;padding:14px;border:1px solid #cdbfaa;border-radius:8px;background:#fffdf8}.pgs-install-panel[hidden]{display:none!important}.pgs-install-panel h4{margin:12px 0 5px}.pgs-install-warning{margin:0;padding:10px;border-left:4px solid #bd7b18;background:#fff4d8;color:#6b4815}.pgs-code{width:100%;min-height:90px;margin-top:10px;resize:vertical;font:12px/1.4 ui-monospace,SFMono-Regular,Menlo,monospace}.pgs-status{min-height:1.55em;margin-top:9px;font-size:13px;color:#285449}@media(max-width:620px){.pgs-tools .toolbar>*{width:100%;text-align:center}}';
    document.head.appendChild(style);

    const section = document.createElement('section');
    section.className = 'pgs-tools';
    section.id = 'pgsTools';
    section.setAttribute('aria-labelledby', 'pgsToolsTitle');
    section.innerHTML = '<h3 id="pgsToolsTitle">填入 AIA PGS 提款設定</h3><div class="hint">工具會複製第 1–72 保單年度的原定提款排程；不會把年齡當作保單年度，也不會代你按 PGS 的「計算及儲存」。</div><div class="toolbar"><button class="primary" id="pgsCopy" type="button" disabled>複製 PGS 提款資料</button><button class="secondary" id="pgsInstall" type="button">加入書籤</button></div><div class="pgs-install-panel" id="pgsInstallPanel" hidden><p class="pgs-install-warning"><strong>尚未加入：</strong>Chrome 和 Safari 不允許普通網頁直接改動你的書籤。下列設定只需在每部電腦完成一次。</p><h4>Chrome／Edge</h4><ol><li>Mac 按 <strong>Command + Shift + B</strong>；Windows 按 <strong>Ctrl + Shift + B</strong>，顯示書籤列。</li><li>在書籤列空白位置按右鍵，選擇「新增網頁／Add page」。</li><li>名稱輸入「PGS 一鍵填入」，網址貼上已複製的書籤程式，然後儲存。</li></ol><h4>Safari（Mac）</h4><ol><li>先按 <strong>Command + D</strong> 建立一個書籤，再開啟「書籤」→「編輯書籤」。</li><li>把名稱改成「PGS 一鍵填入」，並把網址改為已複製的書籤程式。若你的 Safari 版本阻止 JavaScript 書籤，請改用 Chrome。</li></ol><textarea class="pgs-code" id="pgsBookmarkletCode" readonly aria-label="PGS 書籤程式"></textarea><div class="toolbar"><button class="secondary" id="pgsCopyBookmarklet" type="button">再次複製書籤程式</button><button class="secondary" id="pgsCloseInstall" type="button">關閉說明</button></div></div><ol><li>每次先在本頁完成計算，按「複製 PGS 提款資料」。</li><li>轉到 PGS「進階設定」頁，點書籤列的「PGS 一鍵填入」。若瀏覽器拒絕讀取剪貼簿，按提示手動貼上。</li><li>核對逐年金額後，由你親自按 PGS 的「計算及儲存」。</li></ol><div class="pgs-status" id="pgsStatus" role="status">請先完成反推計算。</div>';
    table.closest('.tableWrap')?.after(section);

    const copyButton = document.getElementById('pgsCopy');
    const installButton = document.getElementById('pgsInstall');
    const copyBookmarkletButton = document.getElementById('pgsCopyBookmarklet');
    const closeInstallButton = document.getElementById('pgsCloseInstall');
    copyButton?.addEventListener('click', copyData);
    installButton?.addEventListener('click', requestBookmarkInstall);
    copyBookmarkletButton?.addEventListener('click', copyBookmarkletCode);
    closeInstallButton?.addEventListener('click', () => {
      const panel = document.getElementById('pgsInstallPanel');
      if (panel) panel.hidden = true;
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
    copyBookmarkletCode,
    requestBookmarkInstall,
    revealInstallGuide,
    validatePayload,
  };
})();
