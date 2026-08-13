-- 入帳差異註記：點明細表「差異」欄開啟的視窗所使用的欄位
--
-- diff_note 不沿用 note —— 那欄是出貨報表帶進來的買家訂單備註
-- （「星期一至五收包裹」之類），且對帳單自動建檔會寫入「對帳單匯入」，
-- 兩者混用會互相蓋掉。
--
-- diff_invoice_* 也不沿用 fee_invoice_* / order_invoice_*：那兩組是「群組層」
-- 欄位（同一張發票的每筆訂單重複存同值），這裡要記的是單筆訂單的差異對應發票。
--
-- 附件沿用既有的 Storage bucket `invoices`，路徑前綴 diff/。
ALTER TABLE shipping_orders
  ADD COLUMN IF NOT EXISTS diff_note            text,
  ADD COLUMN IF NOT EXISTS diff_invoice_no      text,
  ADD COLUMN IF NOT EXISTS diff_invoice_date    date,
  ADD COLUMN IF NOT EXISTS diff_invoice_amount  numeric,
  ADD COLUMN IF NOT EXISTS diff_invoice_pdf_url text;
