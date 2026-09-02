# GF 醫療融資反推工具：理想架構

## 1. 文件目的

本文件記錄 GF 醫療融資反推工具日後產品化的理想結構，讓系統可以安全地加入：

- 其他醫療保險計劃及其官方保費表；
- 同一醫療計劃的不同保單形式、自付費及其他選項；
- 其他儲蓄保險產品；
- 同一儲蓄產品的不同供款模式；
- 不同產品專屬的官方文件、保障內容、提款規則及驗證案例。

本文件是未來重構藍圖，不代表現時程式已完成以下結構。

## 2. 不可破壞的原則

1. 重構只改結構，不可改變已驗證的計算結果。
2. 醫療保費、儲蓄模型、介面及 PDF 必須分開管理。
3. 「儲蓄產品」與「供款模式」必須是兩個不同概念。
4. 每個產品必須以固定 ID 識別，不可依靠畫面文字或產品名稱猜測。
5. PDF 不可寫死某一產品的保障內容、保費表檔案或頁碼。
6. 新產品未通過官方建議書交叉驗證前，必須標示為研究／未校準狀態。
7. 用戶設定的開始及結束年齡必須完整輸出；不得靜默減少年份。
8. 內部計算保留所需小數精度，畫面顯示的四捨五入不可影響計算。
9. 所有下載輸出必須能追溯產品版本、保費表版本、模型版本及輸入設定。
10. 現有五年供款及一次性繳費模型要先建立黃金基準，才能開始搬遷。

## 3. 概念模型

```text
醫療融資方案
├── 醫療產品
│   ├── 產品資料
│   ├── 可配置選項
│   ├── 逐年保費率
│   ├── 保障內容
│   └── 官方保費表文件
│
├── 儲蓄產品
│   ├── 產品資料
│   ├── 支援的供款模式
│   ├── 無提款基準
│   ├── 提款後狀態轉移模型
│   ├── 最低基本金額／退保規則
│   └── 大額保費折扣規則
│
├── 融資計算服務
│   ├── 指定供款額計算
│   ├── 反推最低供款額
│   ├── 年齡與保單年度映射
│   └── 完整年期驗證
│
└── 輸出服務
    ├── 網頁結果
    ├── CSV
    ├── PGS 輸入助手
    └── PDF 建議書
```

## 4. 建議目錄

```text
gf_medical_financing/
├── catalog/
│   ├── medical/
│   │   ├── aia_avsw/
│   │   │   ├── manifest.json
│   │   │   ├── rates.json
│   │   │   ├── coverage.json
│   │   │   ├── official-premium-table.pdf
│   │   │   └── fixtures/
│   │   └── another_medical_plan/
│   │       └── ...
│   └── savings/
│       ├── aia_gf/
│       │   ├── manifest.json
│       │   ├── five_year/
│       │   │   ├── model.js
│       │   │   ├── calibration.json
│       │   │   └── fixtures/
│       │   └── single/
│       │       ├── model.js
│       │       ├── calibration.json
│       │       └── fixtures/
│       └── another_savings_plan/
│           └── ...
│
├── core/
│   ├── product_registry.js
│   ├── medical_rate_engine.js
│   ├── financing_engine.js
│   ├── minimum_solver.js
│   ├── range_validator.js
│   ├── result_schema.js
│   └── money.js
│
├── ui/
│   ├── page.html
│   ├── controller.js
│   ├── selectors.js
│   ├── result_table.js
│   └── styles.css
│
├── exports/
│   ├── csv_exporter.js
│   ├── pgs_exporter.js
│   └── pdf/
│       ├── builder.py
│       ├── product_assets.py
│       └── templates/
│
└── tests/
    ├── golden/
    ├── medical_products/
    ├── savings_products/
    ├── exports/
    └── integration/
```

實際目錄名稱可以配合現有 AIAtools 調整，但責任分界應保持一致。

## 5. 醫療產品資料格式

醫療產品不應只固定支援「保單形式」和「自付費」。產品可以自行聲明所需選項，例如性別、吸煙狀況、保障地區、病房級別或保障級別。

`manifest.json` 範例：

```json
{
  "id": "aia_avsw",
  "name": "AIA 自願醫保睿選計劃",
  "version": "2025-10-27",
  "currency": "USD",
  "frequency": "annual",
  "ageRange": { "min": 0, "max": 99 },
  "dimensions": [
    {
      "id": "policy_form",
      "label": "保單形式",
      "options": [
        { "id": "basic", "label": "基本計劃" },
        { "id": "rider", "label": "附加契約" }
      ]
    },
    {
      "id": "deductible",
      "label": "自付費",
      "options": [
        { "id": "0", "label": "HKD 0／USD 0" },
        { "id": "8800", "label": "HKD 8,800／USD 1,100" }
      ]
    }
  ],
  "ratesFile": "rates.json",
  "coverageFile": "coverage.json",
  "officialDocument": {
    "file": "official-premium-table.pdf",
    "pageMap": {
      "policy_form=basic&deductible=0": [0, 1],
      "policy_form=basic&deductible=8800": [2, 3]
    }
  }
}
```

`rates.json` 應以選項組合連接逐年保費，不應把所有數字塞入主頁 HTML：

```json
{
  "policy_form=basic&deductible=8800": {
    "0": 566,
    "1": 566,
    "99": 13024
  }
}
```

載入時必須驗證：

- 年齡沒有遺漏或重複；
- 金額是有限且不小於零的數字；
- 所有選項組合均有唯一保費表；
- 保費表版本及生效日期存在；
- 官方PDF頁碼有效；
- 用戶要求的結束年齡不超過該保費表範圍。

## 6. 儲蓄產品介面

每個儲蓄產品應由 registry 註冊，每種供款模式實作同一份介面：

```javascript
{
  productId: "aia_gf",
  paymentModeId: "five_year",
  name: "AIA GF 五年供款",
  modelVersion: "...",
  validationStatus: "officially-cross-checked",
  minimumWithdrawalPolicyYear: 6,
  validateInput(input) {},
  calculateFixedContribution(input, medicalSchedule) {},
  findMinimumContribution(input, medicalSchedule) {},
  buildPgsPayload(result) {},
  explainResult(result) {}
}
```

一次性繳費是 GF 的另一個 `paymentModeId`，不是另一個醫療計劃，也不應被誤當成另一個儲蓄產品。

如果日後加入新儲蓄產品，必須建立其自己的：

- 基本保額與保費關係；
- 保證現金價值、復歸紅利及終期紅利模型；
- 提款來源及提款後增長規律；
- 最低基本金額及強制退保條件；
- 大額保費折扣；
- 最早提款年度；
- 指定供款額與反推最低供款額算法；
- 官方建議書交叉驗證案例。

不可把 GF 的比例、紅利增長或退保門檻直接套用到另一產品。

## 7. 共用計算結果格式

所有儲蓄產品均應輸出相同欄位，介面、CSV及PDF才能共用：

```json
{
  "schemaVersion": "1",
  "medicalProductId": "aia_avsw",
  "medicalRateVersion": "2025-10-27",
  "savingsProductId": "aia_gf",
  "paymentModeId": "five_year",
  "modelVersion": "...",
  "input": {
    "issueAge": 29,
    "medicalStartAge": 35,
    "medicalEndAge": 99
  },
  "rows": [
    {
      "age": 30,
      "policyYear": 1,
      "requestedWithdrawal": 0,
      "actualWithdrawal": 0,
      "cumulativeWithdrawal": 0,
      "withdrawalFromGuaranteed": 0,
      "withdrawalFromReversionary": 0,
      "withdrawalFromTerminal": 0,
      "postBasicAmount": 0,
      "postGuaranteedCash": 0,
      "postReversionaryBonus": 0,
      "postTerminalBonus": 0,
      "postTotalValue": 0,
      "status": "active"
    }
  ]
}
```

核心層應檢查：

- 第一行年齡等於投保年齡加一；
- 年齡及保單年度逐年連續；
- 最後一行等於用戶指定結束年齡；
- 每年醫療保費已正確映射到提款排程；
- 未退保情況下實際提款足以支付當年要求；
- 退保後不可再出現重新增長的保單價值；
- 累計提款不可倒退；
- 所有金額均為有效數字。

## 8. 介面

介面只從產品 registry 產生選項，不應在 HTML 寫死產品：

1. 選擇醫療產品；
2. 根據該產品的 `dimensions` 動態產生選項；
3. 選擇儲蓄產品；
4. 顯示該產品支援的供款模式；
5. 根據供款模式顯示指定供款／反推最低供款；
6. 所有輸入改動均使舊結果及下載按鈕失效；
7. 計算成功後才顯示CSV、PDF及PGS工具。

## 9. PDF與其他輸出

PDF產生器應接收產品ID和版本，再由伺服器載入可信產品資料；不可相信瀏覽器自行提交的產品名稱、保障內容或官方PDF頁碼。

每個醫療產品自行提供：

- 正式產品名稱；
- 保障內容；
- 官方保費表檔案；
- 不同選項組合對應的PDF頁面；
- 免責聲明及生效日期。

每個儲蓄產品自行提供：

- 產品名稱；
- 供款方式顯示文字；
- 年繳／一次性／總供款摘要；
- 建議書免責聲明；
- 模型驗證狀態。

PDF、CSV及PGS輸出必須帶有一致的版本資料，避免網頁結果與下載內容使用不同模型。

## 10. 驗證策略

### 10.1 黃金基準

重構前保存現有已確認案例的未四捨五入結果，包括：

- 逐年提款；
- 保證現金價值；
- 復歸紅利；
- 終期紅利；
- 基本金額；
- 總額；
- 累計提款；
- 退保狀態及退保年度。

現有29歲投保、35歲開始融資、99歲結束的70年比較案例，可作第一份完整回歸基準。

### 10.2 每個醫療產品

- 每種選項組合至少測試首年、中段及最後一年保費；
- 檢查所有年齡完整；
- 檢查下載PDF加入正確官方頁面；
- 檢查保障內容屬於正確產品；
- 與官方保費表抽樣核對。

### 10.3 每個儲蓄產品及供款模式

- 至少三份無提款官方基準；
- 多份不同開始年齡、提款年期及供款額的提款案例；
- 指定供款與反推最低供款均須驗證；
- 驗證低一個搜尋單位確實不足；
- 驗證內部小數、畫面整數與官方顯示規則；
- 驗證退保門檻及最後可提款年度。

### 10.4 發布門檻

任何重構或新產品發布前必須：

1. 現有黃金基準逐欄零差異；
2. 新產品資料驗證全部通過；
3. 網頁、CSV及PDF的年期和金額一致；
4. PDF逐頁渲染並視覺檢查；
5. 正式環境只更新1008，不影響其他Docker；
6. 測試完成並獲確認後才推送GitHub。

## 11. 新增產品的理想流程

### 新增醫療保險

1. 建立產品資料夾及固定產品ID；
2. 放入官方保費表原檔；
3. 匯入並核對所有逐年保費；
4. 定義產品選項維度；
5. 加入保障內容及PDF頁碼對應；
6. 執行資料完整性及PDF測試；
7. registry註冊後，介面自動出現新產品。

不應修改核心計算引擎。

### 新增儲蓄產品

1. 建立新產品及供款模式資料夾；
2. 取得足夠官方無提款及提款建議書；
3. 建立獨立模型，不沿用GF未經證明的規律；
4. 實作共用產品介面；
5. 輸出標準結果格式；
6. 建立固定官方交叉驗證案例；
7. registry註冊後，介面自動出現新產品。

不應修改既有GF模型。

## 12. 建議重構順序

1. 先凍結及匯出現有黃金基準；
2. 把醫療保費資料從HTML移到產品目錄；
3. 建立醫療產品registry並由資料生成選單；
4. 把PDF的產品資料及頁碼移到產品manifest；
5. 定義共用儲蓄模型介面及結果schema；
6. 先接回GF五年供款，確認零差異；
7. 再接回GF一次性繳費，確認零差異；
8. 移除HTML字串替換及補丁式注入；
9. 完成全套網頁、CSV、PDF及PGS回歸測試；
10. 完成人手驗收後，才開始加入第二個產品。

## 13. 完成定義

當以下條件同時成立，才可稱為可擴展架構：

- 新醫療產品只需增加產品包及測試，不需要修改核心計算；
- 新儲蓄產品只需實作標準模型介面及測試，不需要修改既有產品；
- 所有選單由registry動態生成；
- PDF沒有任何AVSW或GF專屬硬編碼；
- 舊GF案例重構前後內部數值逐欄完全相同；
- 新產品未校準時不能被標示為正式或100%準確；
- 每份輸出均可追溯至明確的產品、保費表及模型版本。
