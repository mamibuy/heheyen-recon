// ============================================================
// 轉換核心 — 套商品對照表，把訂單拆成主商品+贈品列，產出三區塊
// ============================================================
import { num } from './parsers.js';

// 用對照表比對 product_text，回傳該商品的明細列（主商品+贈品）
// mapping：product_mapping 全部資料；order：單筆訂單
export function expandItems(order, mapping) {
  // 1. 酷澎/官網若報表自帶編碼，且對照表沒命中，可用直帶編碼
  const platformRules = mapping.filter((m) => m.platform === order.platform && m.active !== false);

  // 2. 關鍵字包含比對：找出 match_text 出現在 product_text 裡的 group
  const text = String(order.product_text || '');
  const hitGroups = {};
  for (const m of platformRules) {
    if (m.match_text && text.includes(m.match_text)) {
      if (!hitGroups[m.group_key]) hitGroups[m.group_key] = [];
      hitGroups[m.group_key].push(m);
    }
  }

  // 取命中字串最長的 group（最精準），避免「五入組」誤命中
  let bestGroup = null, bestLen = -1;
  for (const gk in hitGroups) {
    const mlen = Math.max(...hitGroups[gk].map((m) => m.match_text.length));
    if (mlen > bestLen) { bestLen = mlen; bestGroup = gk; }
  }

  const lines = [];
  if (bestGroup) {
    const rows = hitGroups[bestGroup].sort((a, b) => a.sort_order - b.sort_order);
    const orderQty = order.qty_platform || 1;
    for (const m of rows) {
      lines.push({
        role: m.role,
        code: m.code,
        item_name: m.item_name,
        // 主商品數量 = 對照表qty（已是該商品的盒數）；如平台一單買多組可乘 orderQty
        qty: num(m.qty),
        matched: true,
      });
    }
  } else if (order.direct_code) {
    // 對照表沒命中，但報表自帶編碼（酷澎/官網）
    lines.push({
      role: 'main',
      code: order.direct_code,
      item_name: order.product_text || '',
      qty: order.qty_platform || 1,
      matched: true,
    });
  } else {
    // 完全未對應 → 標記，讓使用者去後台補
    lines.push({
      role: 'main',
      code: '',
      item_name: order.product_text || '(未對應)',
      qty: order.qty_platform || 1,
      matched: false,
    });
  }

  // 計算主商品單價：總額 ÷ 主商品數量（總額平攤）
  const mainLine = lines.find((l) => l.role === 'main');
  const mainQty = mainLine ? mainLine.qty : 1;
  const unitPrice = mainQty > 0 ? Math.round((order.total / mainQty) * 100) / 100 : order.total;

  return lines.map((l) => ({
    ...l,
    unit_price: l.role === 'main' ? unitPrice : 0,
    subtotal: l.role === 'main' ? order.total : 0,
  }));
}

// 把所有訂單轉成三區塊
export function buildBlocks(orders, mapping) {
  const block1 = []; // 訂單總表（二）
  const block2map = {}; // 商品彙總（按編碼加總）
  const block3 = []; // 全部物流資料
  let seq = 0;
  const unmatched = [];

  for (const o of orders) {
    const items = expandItems(o, mapping);
    seq += 1;

    items.forEach((it, idx) => {
      if (!it.matched) unmatched.push({ ref_no: o.ref_no, text: o.product_text });
      // 區塊一：主列填滿，贈品列只填編碼/品項/數量
      if (idx === 0) {
        block1.push({
          編號: seq, 參照編號: o.ref_no, 日期: o.order_date,
          聯絡人: o.contact, 地址: o.address, 電話: o.phone, Email: o.email,
          付款方式: o.pay_method, 後五碼: '', 備註: o.note, 統編: '',
          商店: o.store, 件數: o.pkg_count, 託運單號: o.tracking_no, 出貨日期: '',
          編碼: it.code, 品項: it.item_name, 數量: it.qty,
          含稅單價: it.unit_price, 小計: it.subtotal, 總計: o.total,
          含稅購物金: '', 含稅折扣: '', Status: '正式', 'Dispatch Status': '待出貨',
        });
      } else {
        block1.push({
          編號: '', 參照編號: '', 日期: '', 聯絡人: '', 地址: '', 電話: '', Email: '',
          付款方式: '', 後五碼: '', 備註: '', 統編: '', 商店: '', 件數: '',
          託運單號: '', 出貨日期: '',
          編碼: it.code, 品項: it.item_name, 數量: it.qty,
          含稅單價: 0, 小計: 0, 總計: '', 含稅購物金: '', 含稅折扣: '',
          Status: '', 'Dispatch Status': '',
        });
      }

      // 區塊二：按編碼加總
      if (it.code) {
        if (!block2map[it.code]) {
          block2map[it.code] = { 數量: 0, 品項: it.item_name, 編碼: it.code, 品牌: '', '包裝尺寸 (長 x 寬 x 高)': '', 重量: '', 含稅額: 0 };
        }
        block2map[it.code].數量 += num(it.qty);
        block2map[it.code].含稅額 += num(it.subtotal);
      }
    });

    // 區塊三：全部物流資料，一筆訂單一列
    block3.push({
      聯絡人: o.contact, 地址: o.address, 電話: o.phone, 付款方式: o.pay_method,
      總計: o.total, 未付款總額: 0, 件數: o.pkg_count, 備註: o.note,
      參照編號: o.ref_no, 託運單號: o.tracking_no,
      Transporter: '新竹物流', 'Transporter Account Number': '',
    });
  }

  const block2 = Object.values(block2map);
  const total2 = block2.reduce((s, r) => s + num(r.含稅額), 0);
  block2.push({ 數量: '', 品項: 'total', 編碼: '', 品牌: '', '包裝尺寸 (長 x 寬 x 高)': '', 重量: '', 含稅額: total2 });

  return { block1, block2, block3, unmatched };
}
