# 和和研電商出貨彙整系統（階段 1）

> 獨立專案。把各平台訂單報表彙整成「統整出貨明細」格式，餵 ERP 產生 SA 銷貨單號。

## 功能
- **出貨轉換**：上傳平台報表（蝦皮 / LINE商城 / 酷澎 / 官網）→ 自動辨識平台 → 套商品對照表拆主商品+贈品 → 產出三區塊（訂單總表二、商品彙總、全部物流資料）→ 下載 Excel
- **商品對照表後台**：新增/編輯/刪除對照規則，平台上新商品時自己維護，不需改程式
- 未對應商品會標記提醒

## 第一次設定

### 1. 建立 Supabase 資料表
到 Supabase 專案的 SQL Editor，依序貼上執行：
- `sql/01_schema.sql`（建表）
- `sql/02_seed_mapping.sql`（灌入已驗證的對照規則）

### 2. 填入 Supabase anon key
打開 `src/App.jsx`，把第 10 行的 `PLACEHOLDER_REPLACE_WITH_REAL_ANON_KEY`
換成你的 Supabase anon key（與 Mamibuy 同一專案，可從 Mamibuy 的 App.jsx 複製）。

### 3. 部署到 GitHub Pages
建立一個新的 public repo 叫 `heheyen-recon`，然後：
```bash
cd heheyen-recon
npm install
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/你的帳號/heheyen-recon.git
git push -u origin main
```
到 repo 的 Settings → Pages → Source 選 **GitHub Actions**，
等 Actions 跑完，網址在 `https://你的帳號.github.io/heheyen-recon/`

## 技術棧
React 18 + Vite 5 + Supabase + SheetJS(xlsx) + GitHub Pages
