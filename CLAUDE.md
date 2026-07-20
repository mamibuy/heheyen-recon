# 和和研電商出貨彙整系統（HEHEYEN Recon）

React 18 + Vite 5 單檔架構，把各電商平台訂單報表彙整成統整出貨明細，並提供六條金流路徑對帳。

## 目錄結構
```
heheyen-recon/          ← 實際原始碼（注意：外層同名資料夾只放建置文件）
├── src/
│   ├── App.jsx         ← 全部 UI（140KB+，三個頂層 component：App / ReconPage / GatewayWorkspace）
│   ├── parsers.js      ← 四平台出貨報表解析器 + detectPlatform() + excelDate()
│   ├── transform.js    ← expandItems() / buildBlocks() → 三區塊輸出
│   ├── recon_parsers.js← 六條金流撥款明細解析器 + detectGateway()
│   └── reconcile.js    ← reconcile() 比對 shipping_orders；previewInvoice / applyInvoice
├── public/
│   └── recon-guide.png ← 金流對帳說明圖（說明頁籤用）
└── sql/                ← Supabase schema 與 migration（01~06）
```

## 技術棧
- React 18 + Vite 5；vite base = `/heheyen-recon/`
- Supabase PostgreSQL — URL: `https://geirbvjkwsewglvvrfmg.supabase.co`
- SheetJS (xlsx)：`sheet_to_json(ws, { header: 1 })` for array mode；`{ defval: '' }` for object mode
- GitHub Pages + Actions 自動部署（push main 即部署）

## 指令
```bash
npm install
npm run build   # 編譯，部署前必跑確認無錯
# 部署：git push origin main（GitHub Actions 自動接手）
```
沒有本地測試環境，一律靠部署驗證。

## 密碼保護
前端有 `<PasswordGate>` 包住整個 App，密碼存在 localStorage（`hhy_auth`）。密碼：imheheyen（明文比對，僅基本防護）。

## 三個主頁籤
1. **出貨轉換** — 上傳平台報表 → 解析 → 套商品對照表 → 匯出天心格式 Excel
2. **商品對照表** — 維護 `product_mapping`（Supabase）
3. **金流對帳** — 六條金流路徑對帳 + 發票核對，子頁籤：說明 / 酷澎 / 蝦皮 / 官網>信用卡 / 官網>LINE Pay / LINE商城>LINE Pay / LINE商城>信用卡

## 業務規則（勿刪、勿「整理」）

### 出貨轉換
- 商品對照：`match_text` 關鍵字包含比對；多個命中取最長者（最精準）
- 一筆訂單 → 主商品列 + N 贈品列；贈品列只填編碼/品項/數量，單價=0
- 主商品單價 = 總額 ÷ mainQty
- 蝦皮/酷澎地址填 `蝦皮隱碼`/`酷澎隱碼`；LINE商城/官網填真實值
- 酷澎/官網若報表自帶 `direct_code`，對照表未命中時直接使用

### 金流對帳（六條路徑）
每條路徑的手續費公式和比對鑰匙不同，改任何一條前先查 `/RECON_BUILD.md`。

| 路徑 | 對帳單 | 比對鑰匙 | 手續費欄位 |
|------|--------|----------|-----------|
| 蝦皮 | 我的進帳（新格式） | B欄訂單編號 | J/K/L/P/Q/R/U/V/W/X 加總 |
| 蝦皮 | 撥款明細（舊格式） | 訂單編號 | 成交手續費+其他服務費+金流處理費 |
| 官網>信用卡 | PayUni 入帳表 | 商店訂單編號（去槓） | 手續費欄（取絕對值） |
| 官網>LINE Pay | D-1(LinePay)+D-2(PayUni) | D-2商店訂單編號（去槓）→D-1交易號碼前15碼 | D-1手續費合計 |
| LINE商城>LINE Pay | 舊 LINE Pay 撥款 | 訂單號碼 | 手續費合計（含稅） |
| LINE商城>信用卡 | 藍新金流撥款 | 商店訂單編號 | 主支付手續費 |
| 酷澎 | 酷澎出帳明細 | 訂單編號（合併商品+DELIVERY_FEE兩列） | 手續費總額+固定4元 |

蝦皮新格式額外欄位：
- Y欄（錢包入帳金額）→ `payable`（應入帳）
- I~O欄加總 → `order_invoice_amount`（代收付發票金額）

### 蝦皮玉山銀行對帳
上傳玉山 XLS，篩選條件：備註含「SHOPEE」**或**帳號含「808/0370979139156」。
確認入帳時 `actual_in = br.deposit`（銀行實際存入金額，非訂單應入帳）。

### 官網 LINE Pay 雙層費用發票
兩張月結總額發票分開核對：LINE Pay 手續費（備注含 LINE Pay）/ PayUni 服務費（備注含信用卡）。

## Supabase shipping_orders 主要欄位
`id, ref_no, sa_no, platform, total, fee_total, payable, actual_in, in_date, bank_deposit, order_date, pay_method, note, recon_status, invoice_check, fee_invoice_no, fee_invoice_date, fee_invoice_amount, order_invoice_no, order_invoice_date, order_invoice_amount, tx_code, tx_fee, tx_fee_invoice_no, fee_invoice_pdf_url, tx_fee_invoice_pdf_url`

- `recon_status`：`未對帳` / `平台已結算` / `已入帳` / `已對帳`
- 數字欄位不可傳空字串，先轉 0 或 null

## 常見陷阱
- `isShopee`、`isPayuniCC` 等 flag 只在 `GatewayWorkspace` 作用域內有效，`ReconPage` 用 `activeGateway !== 'shopee'` 判斷
- SheetJS 讀玉山 XLS：header 在 raw array 第 5 列（index 5），`sheet_to_json` 無須設 range
- `detectGateway()` 依 headers 特徵自動判斷金流，新增格式需同步更新
- Supabase client 方法不可命名為保留字（如 `in`），會導致 esbuild 失敗
- repo 是 public；untracked 的 `03/04/06 sql seed` 含客戶 PII，不可 git add

## 重要注意事項
- 確認功能正常後再告知完成，不要只靠 build 成功就宣告 done
- 部署前先問是否要 push（預設不自動 push）
- 直接寫入 Supabase 前需確認（production DB，無沙箱）
- 密碼與 credentials 不要輸入任何欄位
