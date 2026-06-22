# 和和研電商出貨彙整系統

獨立專案。把各電商平台訂單報表彙整成「統整出貨明細」格式，餵 ERP 產生 SA 銷貨單號。

## 技術棧
- React 18 + Vite 5（單檔 App.jsx 架構）
- Supabase（PostgreSQL）— URL: https://kyzyozfdjqlhpcxtzzfp.supabase.co
- xlsx (SheetJS) 解析/匯出
- GitHub Pages + Actions 部署；vite base = `/heheyen-recon/`

## 指令
- `npm install` 安裝
- `npm run build` 編譯（部署前先跑這個確認沒錯）
- 部署：push 到 main，GitHub Actions 自動部署

## 架構
- `src/parsers.js` — 四平台（蝦皮/LINE商城/酷澎/官網）解析器、平台自動辨識、Excel 日期轉換
- `src/transform.js` — 套商品對照表拆主商品+贈品、產出三區塊
- `src/App.jsx` — UI：出貨轉換頁 + 商品對照表後台
- `sql/` — Supabase 建表與初始資料

## 業務規則（修改時務必遵守）
- 商品對照用「關鍵字包含比對」，存 Supabase 的 product_mapping，使用者可自維護
- 一筆訂單拆「主商品列 + N 贈品列」；贈品列只填編碼/品項/數量
- 主商品單價 = 總額 ÷ 主商品數量
- 贈品編碼不限 FREEGIFT 開頭；贈品數量可 > 1
- 隱碼：蝦皮填「蝦皮隱碼」、酷澎填「酷澎隱碼」；LINE商城/官網用真實值
- 三區塊：訂單總表（二）、商品彙總（按編碼加總）、全部物流資料

## 注意事項
- repo 是 public，靠 Supabase anon key + RLS（policy 開放讀寫）
- 改 Supabase client 時，不要把方法命名為保留字（如 `in`），會導致 esbuild 失敗
- 數字欄位不要傳空字串給 Supabase，先轉 0 或 null
- 沒有本地測試環境時，一律靠 GitHub Pages 部署驗證
