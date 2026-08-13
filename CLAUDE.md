# 和和研電商出貨彙整系統（HEHEYEN Recon）

React 18 + Vite 5 單檔架構，把各電商平台訂單報表彙整成統整出貨明細，並提供**七條**金流路徑對帳。

## 目錄結構
```
heheyen-recon/          ← 實際原始碼（注意：外層同名資料夾只放建置文件）
├── src/
│   ├── App.jsx         ← 全部 UI（190KB+，三個頂層 component：App / ReconPage / GatewayWorkspace）
│   ├── parsers.js      ← 四平台出貨報表解析器 + detectPlatform() + excelDate()
│   ├── transform.js    ← expandItems() / buildBlocks() → 三區塊輸出
│   ├── recon_parsers.js← 七條金流撥款明細解析器 + detectGateway() + 欄位格式檢查
│   └── reconcile.js    ← reconcile() 比對 shipping_orders；previewInvoice / applyInvoice
├── public/
│   └── recon-guide.png ← 金流對帳說明圖（說明頁籤用）
└── sql/                ← Supabase schema 與 migration（01~06）
```

## 技術棧
- React 18 + Vite 5；vite base = `/heheyen-recon/`
- Supabase PostgreSQL — URL: `https://geirbvjkwsewglvvrfmg.supabase.co`
  Key: `sb_publishable_yDgLU7V2PcL_2QmrQkxo2w_WZGEbP63`（repo 為 public，靠 RLS）
- SheetJS (xlsx)：`sheet_to_json(ws, { header: 1 })` array mode；`{ defval: '' }` object mode
- GitHub Pages + Actions 自動部署（push main 即部署，約 40~60 秒）

## 指令
```bash
npm install
npm run build   # 編譯，部署前必跑確認無錯
# 部署：git push origin main（GitHub Actions 自動接手）
```

**本地可以跑 dev server 驗證**（`npm run dev`，vite base 是 `/heheyen-recon/`，
所以要開 `http://localhost:5173/heheyen-recon/`，開根目錄會是空白頁）。
驗證部署是否生效：比對 `dist/assets/index-*.js` 的檔名與線上 HTML 引用的檔名。

## 密碼保護
`<PasswordGate>` 包住整個 App，通過後存 localStorage（`hhy_auth = '1'`）。
密碼明文比對，僅基本防護。**開發驗證時直接設 localStorage 即可，不要輸入密碼。**

## 三個主頁籤
1. **出貨轉換** — 上傳平台報表 → 解析 → 套商品對照表 → 匯出天心格式 Excel
2. **商品對照表** — 維護 `product_mapping`（Supabase）
3. **金流對帳** — 七條金流路徑對帳 + 發票核對

---

## 金流對帳頁的結構

### 天心銷貨單：跨通路共用，位置在通路頁籤「之上」
`handleTianxin` 撈訂單時**不帶 platform 條件**，一次上傳就比對全部平台的訂單編號，
把 SA 單號與訂單發票號碼寫進對應訂單。上傳元件是獨立卡片（`tianxinPanel`），
狀態存在 `ReconPage`，切換頁籤不會掉檔案。回填後遞增 `txVersion` 通知當前頁籤重載。

各頁籤的「銷貨單號」步驟卡只是**進度指示**（顯示本通路尚缺幾筆），沒有上傳欄位。

### 步驟順序（2026-08-06 依實際作業流程重排）
```
① 上傳該渠道的對帳單／撥款明細
② 天心銷貨單號
③ 手續費發票核對
④ 玉山銀行對帳（收尾，七條金流都有）
```
渠道專屬步驟依性質靠邊：費用類緊接撥款明細，發票類靠發票核對。
**要再調整順序前請先跟使用者確認，勿依設計稿或舊版線上順序自行重排。**

| 頁籤 | 步驟 |
|---|---|
| 酷澎 / 官網›信用卡 / LINE商城×2 / 兆豐 | 撥款明細 → 銷貨單號 → 發票核對 → 玉山 |
| 蝦皮 | 訂單匯入 → 撥款明細 → 銷貨單號 → 訂單發票 → 發票核對 → 玉山 |
| 官網›LINE Pay | 撥款明細 → 銷貨單號 → 手續費發票 → 交易處理費發票 → 玉山 |

官網›LINE Pay 的「發票核對」步驟顯示為**手續費發票**（該通路有兩種發票，需區隔）。
原本的獨立「交易處理費」步驟已移除（D-2 已逐筆提供 tx_fee，屬重複輸入）。

### 入帳差異註記（`diff_*` 欄位，sql/07）
明細表「差異」欄有數字時可點，開啟視窗記錄原因與對應發票
（號碼／日期／金額／附件）。附件用既有的 Storage bucket `invoices`，路徑前綴 `diff/`，
選檔即上傳並寫回 DB，不等按「儲存」。

**不可改用 `note`** —— 那是出貨報表帶進來的買家訂單備註（「星期一至五收包裹」之類），
且 `reconcile` 自動建檔會寫入「對帳單匯入」。
也不可沿用 `fee_invoice_*` / `order_invoice_*` —— 那兩組是群組層欄位（同張發票每筆重複存同值），
`diff_invoice_*` 記的是單筆訂單的對應發票。

蝦皮明細表最下方有固定說明：每月最後一筆入帳會扣掉前月代開發票總額，該筆差異屬正常。

### 發票分組卡片（`InvoiceGroupCards`）
代開／手續費／交易處理費三種發票共用同一個元件，卡片區塊可收合，
狀態存 `localStorage['invcards_open_{gateway}_{kind}']`（`'0'` = 收合），
依通路與發票種類各記各的。預設展開。

### 玉山銀行對帳
七條金流都採**手動勾選訂單**比對（`isManualSelection` 永遠為真；
早期依撥款日自動篩選的分支保留但不會走到）。
酷澎／兆豐的報表沒有預計撥款日，`payoutRows` 對這兩條金流回傳 `[]`。

「已確認入帳」群組（依 `in_date` 彙總）可**解除整組**或**移除單筆**，
會清空 `in_date` / `actual_in` / `bank_deposit` 並把狀態改回「平台已結算」。

---

## 業務規則（勿刪、勿「整理」）

### 出貨轉換
- 商品對照：`match_text` 關鍵字包含比對；多個命中取最長者（最精準）
- 一筆訂單 → 主商品列 + N 贈品列；贈品列只填編碼/品項/數量，單價=0
- 主商品單價 = 總額 ÷ mainQty
- 蝦皮/酷澎地址填 `蝦皮隱碼`/`酷澎隱碼`；LINE商城/官網填真實值
- 酷澎/官網若報表自帶 `direct_code`，對照表未命中時直接使用

### 七條金流的比對鑰匙與手續費

| 路徑 | 對帳單 | 比對鑰匙 | 手續費 |
|------|--------|----------|-------|
| 蝦皮 | 我的進帳（新格式） | 訂單編號 | J/K/L/P/Q/R/U/V/W/X 加總取絕對值 |
| 蝦皮 | 撥款明細（舊格式） | 訂單編號 | 成交手續費＋其他服務費＋金流處理費 |
| 官網›信用卡 | PayUni 入帳表／交易入帳表 | 商店訂單編號（去槓） | 手續費欄取絕對值 |
| 官網›LINE Pay | D-1(LinePay)＋D-2(PayUni) | D-2商店訂單編號（去槓）→ D-1交易號碼**前15碼** | D-1 手續費合計 |
| LINE商城›LINE Pay | LINE Pay 撥款明細 | 訂單號碼 | 手續費合計（含稅） |
| LINE商城›信用卡 | 藍新金流撥款 | 商店訂單編號 | 主支付手續費 |
| 酷澎 | 酷澎出帳明細 | 訂單編號（合併商品＋DELIVERY_FEE兩列） | 手續費總額＋固定4元 |
| 兆豐福利網 | 訂單明細報表＋手續費報表（雙檔） | 訂單編號 | 手續費報表的手續費（固定 6%） |

### 兆豐福利網（雙檔，2026-08-06 新增）
- **應收 = 商品金額 ＋ 運費**。不可用「付款總金額」：客戶全額以福利金折抵時該欄為 0，
  但兆豐仍照訂單金額扣 6% 撥款。「訂單總金額」欄位在未動用福利金時是空的，同樣不可用。
- 手續費報表**沒有訂單編號欄**，需從備註 `[訂單編號：607020108] 商品名…` 以 regex 抽取。
- 訂單明細的訂單編號**帶前導空白**，務必 trim。
- 內建 6% 費率檢核：查無手續費或費率偏離 6% 即示警。
- 兆豐訂單只從對帳單進來（沒有出貨轉換那條路），故 `reconcile` 會一併寫入 order_date。

### 訂單日期回填（各通路，一律「只補空白」）
對帳單自帶訂單日期時寫入 `order_date`，但**僅補原本沒有日期的訂單**，
由 parser 回傳 `order_date_fill_only: true` 控制，`reconcile` 判斷：
```js
if (row.order_date && !(row.order_date_fill_only && order.order_date)) { ... }
```
理由：出貨報表帶入的下單日才是正確的；對帳單的日期是付款/交易日，
且退款列填的是退款日，覆蓋會失真。

| 通路 | 來源欄位 |
|---|---|
| 蝦皮 | 訂單成立日期 |
| 官網›信用卡 | 付款(退款)日期 |
| 官網›LINE Pay | 交易日期（D-2） |
| LINE商城›LINE Pay | 交易日期 |
| 兆豐 | 訂單時間（覆蓋，非 fill-only） |

### 官網›LINE Pay：D-2 有兩種格式
「交易動態明細」用 `付款金額`，「交易查詢表」用 `收款金額`／`訂單金額` — pick 全都列入，
否則 total 會是 0，D-1 沒勾稽到時 payable 會變成負的手續費。
**只處理交易狀態為「已付款」的列**；付款取消等未成立交易若照寫，勾自動建檔會產生
應收 0／應入帳 0 的垃圾訂單。被略過的列掛在回傳陣列的 `.skipped`，UI 會列出來不靜默丟棄。
舊版報表無「交易狀態」欄時不過濾，維持相容。

### 蝦皮：訂單（代開）發票群組化
一個月開立一張代開發票跟蝦皮請款，依**發票號碼**群組多筆訂單。

**應開發票金額 = I欄商品原價 − M欄賣家負擔優惠券**，該值恰等於 `payable + fee_total`
（因 `Y = I + M − 手續費`），故直接由既有欄位回推，不需另存欄位、既有訂單立即套用。

`order_invoice_amount` 語意為「**該組實際開立的代開發票金額**」（每筆重複存同一值，
作法同 `fee_invoice_amount`），由使用者在訂單發票步驟輸入。
**parser 不再寫入此欄**（原本存 I~O 加總的代收金額，會在重傳報表時覆蓋使用者填的金額）。

明細表「代開發票金額」欄：**已開立**（有 `order_invoice_no`）顯示該組實開金額；
**未開立**顯示 `payable + fee_total` 回推的應開金額（灰字加註「應開」）。
未開立時刻意不讀 `order_invoice_amount` —— 舊版 parser 寫入的殘值仍留在 DB
（6 月三筆未開立訂單），不同步且會讓同樣未開立的訂單有的有值有的空白。

⚠️ 已知邊界：`payable + fee_total` 實為 I~O 欄加總。目前報表 J/K/L/N/O 全為 0 故與 `I − M`
一致；若日後出現退款、蝦幣回饋券或買家支付運費則會分歧，屆時需改為逐欄取值並另存欄位。

### 蝦皮玉山銀行對帳
上傳玉山 XLS，篩選條件：備註含「SHOPEE」**或**帳號含「808/0370979139156」。
確認入帳時 `actual_in = br.deposit`（銀行實際存入金額，非訂單應入帳）。

### 官網 LINE Pay 雙層費用發票
兩張月結總額發票分開核對：LINE Pay 手續費（備註含 LINE Pay）／PayUni 服務費（備註含信用卡）。

---

## 檔案上傳

### 支援格式
`.xlsx` / `.xls` / `.csv`（9 個上傳欄位都支援）。

### 表頭不在第一列 —— 兩個獨立的動態偵測
1. **`sheetToRows(ws)`**（對帳單通用，`readFile` 使用）
   先用預設方式解析，欄位裡找得到鑰匙欄位（訂單編號／商店訂單編號／訂單號碼／
   交易號碼／客戶訂單）就直接採用；找不到才往下尋找含鑰匙欄位的那一列當表頭。
   **蝦皮「我的進帳」表頭在第 6 列**（前面有賣家帳號區塊、空列、小計列），
   不做這個處理會整份檔案解析出 0 筆。
2. **`parseTianxinSheet(ws)`**（天心銷貨明細專用）
   找同時含「客戶訂單」「單號」的列當表頭。**每格先 trim 再比對**，
   否則欄位帶空白時 `findIndex` 回 -1 會靜默退回錯誤的第一列。

組物件時 key 用 trim 過的 header、**value 取原始列**，否則數字欄會被轉成字串。

### 欄位格式檢查（`checkReconColumns` / `checkDualReconColumns` / `checkMegabankColumns`）
每條金流列出 parser 真正會讀的關鍵欄位，parse 前先檢查，不符即擋下並提示缺什麼欄位。
**沒有用 `detectGateway()` 來擋** —— 它是關鍵字啟發式，容易誤傷合法檔案。
UI 有「略過格式檢查」checkbox 供平台改欄位名時強制執行。

> 起因：誤把酷澎「營業稅明細 VAT History Report」當成「出帳明細」上傳，
> 該報表沒有手續費／金額欄位，parser 一路讀成 0，再套上寫死的 4 元其他服務費，
> 產生「應收 0 / 應入帳 -4」這種看似成功實則垃圾的結果，還被自動建檔寫進 DB。

### CSV 注意事項
- PayUni 匯出的 CSV 欄位帶 Excel 強制文字語法（`"=""P4WZMRHB74KV"""`），
  SheetJS 會正確還原成純字串。
- 若匯出時**沒有**該語法，19 位數的「支付方式對應碼」會被當數字而**掉精度**（末幾位跑掉）。
  勾稽用前 15 碼所以不受影響，但存進 DB 的值會與原始值不同。
- SheetJS 讀 CSV 預設 UTF-8。若遇到 Big5 編碼的 CSV 會是亂碼，屆時需補 `codepage: 950`
  （目前無樣本可驗，未預先寫）。

---

## Supabase `shipping_orders` 主要欄位
`id, ref_no, sa_no, platform, total, fee_total, payable, actual_in, in_date, bank_deposit,
order_date, pay_method, note, recon_status, invoice_check, fee_invoice_no, fee_invoice_date,
fee_invoice_amount, order_invoice_no, order_invoice_date, order_invoice_amount, tx_code,
tx_fee, tx_fee_invoice_no, fee_invoice_pdf_url, tx_fee_invoice_pdf_url,
diff_note, diff_invoice_no, diff_invoice_date, diff_invoice_amount, diff_invoice_pdf_url`

- `recon_status`：`待出貨` / `已出貨` / `平台已結算` / `已入帳` / `已對帳`
- `platform`：`蝦皮` / `LINE商城` / `酷澎` / `官網` / `兆豐`
- 數字欄位不可傳空字串，先轉 0 或 null
- `gateway_sops` 表：教學 SOP 內容，以 gateway 為 key，`html_content` 存 `{v:1, steps:{stepKey:{html,img}}}`

---

## 常見陷阱
- `isShopee`、`isPayuniCC` 等 flag 只在 `GatewayWorkspace` 作用域內有效
- **狀態與欄位要連動**：從訂單編輯視窗清掉 `in_date` 不會改 `recon_status`，
  會變成「沒有入帳日卻仍是已入帳」。清單篩選已改為
  `o.recon_status !== '已入帳' || !o.in_date` 以容錯，但正確做法是用「解除整組／移除」按鈕。
- `detectGateway()` 依 headers 特徵自動判斷金流，新增格式需同步更新（目前實際未用於擋檔）
- Supabase client 方法不可命名為保留字（如 `in`），會導致 esbuild 失敗
- SOP 內容以 **step key** 儲存且 `collectSopSteps` 是 `{...base}` 疊加式（只增不刪），
  步驟從畫面移除不會刪掉 DB 內容
- repo 是 public；untracked 的 `03/04/05/06 sql seed` 含客戶 PII，**不可 git add**
  （目前未列入 .gitignore，`git add .` 會誤傷 —— 建議補上）
- `scripts/` 與 `.claude/` 已在 .gitignore（一次性修資料腳本含真實訂單／發票號碼）

## 重要注意事項
- **確認功能正常後再告知完成，不要只靠 build 成功就宣告 done**
- 部署前先問是否要 push（預設不自動 push）
- 直接寫入 Supabase 前需確認（production DB，無沙箱）。
  一次性修資料腳本放 `scripts/`，由使用者自行執行
- 密碼與 credentials 不要輸入任何欄位
- 使用者常看到瀏覽器快取的舊版 —— 部署後提醒強制重新整理（⌘⇧R）
