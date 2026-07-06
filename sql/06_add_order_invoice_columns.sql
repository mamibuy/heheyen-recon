-- 新增訂單發票日期與金額欄位（供蝦皮訂單發票輸入功能使用）
ALTER TABLE shipping_orders
  ADD COLUMN IF NOT EXISTS order_invoice_date  date,
  ADD COLUMN IF NOT EXISTS order_invoice_amount numeric;
