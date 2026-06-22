-- ============================================================
-- 商品對照表 — 初始資料（已驗證規則）
-- 在 01_schema.sql 之後執行
-- ============================================================

-- 蝦皮：純LGG x3盒 贈5入旅行組+14入體驗組
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('蝦皮', '純LGG活菌益生菌x3盒', 'sp_3box', 'main', '001100POWA001', '純LGG活菌益生菌', 3, 0),
('蝦皮', '純LGG活菌益生菌x3盒', 'sp_3box', 'gift', '001100POWA002', '純LGG活菌益生菌-14入體驗組', 1, 1),
('蝦皮', '純LGG活菌益生菌x3盒', 'sp_3box', 'gift', 'FREEGIFT00002', '純LGG活菌益生菌-5入旅行組', 1, 2);

-- LINE商城：買5送1贈玩沙玩具組
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('LINE商城', '買5送1贈玩沙玩具組', 'lm_5send1', 'main', '001100POWA001', '純LGG活菌益生菌', 5, 0),
('LINE商城', '買5送1贈玩沙玩具組', 'lm_5send1', 'gift', 'FREEGIFT00001', '夏日玩沙玩具組', 1, 1);

-- LINE商城：五入組贈5入旅行組x2+夏日玩沙組x1
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('LINE商城', '五入組', 'lm_5set', 'main', '001100POWA001', '純LGG活菌益生菌', 5, 0),
('LINE商城', '五入組', 'lm_5set', 'gift', 'FREEGIFT00002', '純LGG活菌益生菌-5入旅行組', 2, 1),
('LINE商城', '五入組', 'lm_5set', 'gift', 'FREEGIFT00001', '夏日玩沙玩具組', 1, 2);

-- 酷澎：Heheyen 純LGG活菌益生菌 60g 1個
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('酷澎', 'Heheyen 和和研 純LGG活菌益生菌', 'cp_1', 'main', '001100POWA001', '純LGG活菌益生菌', 1, 0);

-- 官網：三入組送旅行組+兒童保暖斗篷（範例，依實際調整）
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('官網', '三入組送旅行組', 'ma_3set', 'main', '001100POWA001', '純LGG活菌益生菌', 3, 0),
('官網', '三入組送旅行組', 'ma_3set', 'gift', 'FREEGIFT00002', '純LGG活菌益生菌-5入旅行組', 1, 1);

-- 官網：14入體驗組
insert into product_mapping (platform, match_text, group_key, role, code, item_name, qty, sort_order) values
('官網', '14入體驗組', 'ma_14', 'main', '001100POWA002', '純LGG活菌益生菌-14入體驗組', 1, 0);
