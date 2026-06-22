-- ============================================================
-- 和和研電商對帳系統 — 階段 1 資料表
-- 在 Supabase SQL Editor 貼上執行
-- ============================================================

-- ---------- 1. 商品對照表 ----------
-- 每一列：某平台商品「識別字串」→ 主商品/贈品的編碼與數量
-- 設計為「一個編碼一列」：同一個平台商品若含主商品+多贈品，就有多列，用 group_key 綁在一起
create table if not exists product_mapping (
  id          bigint generated always as identity primary key,
  platform    text not null,              -- 平台：蝦皮 / LINE商城 / 酷澎 / 官網
  match_text  text not null,              -- 平台商品識別字串（關鍵字包含比對）
  group_key   text not null,              -- 同一平台商品的綁定鍵（同 match_text 共用）
  role        text not null default 'main', -- main=主商品 / gift=贈品
  code        text not null,              -- 內部編碼（如 001100POWA001 / FREEGIFT00002）
  item_name   text,                       -- 品項名稱（顯示用）
  qty         numeric not null default 1, -- 數量（主商品=盒數；贈品可>1）
  sort_order  int not null default 0,     -- 顯示排序（主商品先、贈品後）
  active      boolean not null default true,
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

create index if not exists idx_pm_platform on product_mapping(platform);
create index if not exists idx_pm_match on product_mapping(match_text);

-- ---------- 2. 出貨單（彙整後，每筆訂單一列；明細存 JSON） ----------
create table if not exists shipping_orders (
  id            bigint generated always as identity primary key,
  platform      text not null,            -- 來源平台
  ref_no        text not null,            -- 參照編號（平台訂單編號）
  order_date    text,                     -- 訂單日期
  contact       text,                     -- 聯絡人
  address       text,                     -- 地址
  phone         text,                     -- 電話
  email         text,
  pay_method    text,                     -- 付款方式
  note          text,                     -- 備註
  store         text,                     -- 商店
  pkg_count     int default 1,            -- 件數
  tracking_no   text,                     -- 託運單號
  total         numeric,                  -- 總計
  shipping_fee  numeric,                  -- 運費
  -- 明細：主商品+贈品列（陣列），格式 [{role,code,item_name,qty,unit_price,subtotal}]
  items_json    jsonb,
  -- 對帳欄位（階段 2+ 回填）
  fee_total     numeric,                  -- 手續費（金流與系統處理費）
  fee_rate      text,                     -- 費率
  payable       numeric,                  -- 應付/應入帳
  actual_in     numeric,                  -- 實際入帳
  in_date       text,                     -- 入帳日
  recon_status  text default '待出貨',     -- 待出貨/已出貨/平台已結算/已入帳/已對帳
  batch_tag     text,                     -- 月份/批次標籤
  created_at    timestamptz default now(),
  updated_at    timestamptz default now(),
  unique(platform, ref_no)
);

create index if not exists idx_so_platform on shipping_orders(platform);
create index if not exists idx_so_ref on shipping_orders(ref_no);
create index if not exists idx_so_status on shipping_orders(recon_status);

-- ---------- RLS：比照 Mamibuy，開放匿名讀寫（repo public，靠 anon key） ----------
alter table product_mapping enable row level security;
alter table shipping_orders enable row level security;

drop policy if exists pm_all on product_mapping;
create policy pm_all on product_mapping for all using (true) with check (true);

drop policy if exists so_all on shipping_orders;
create policy so_all on shipping_orders for all using (true) with check (true);
