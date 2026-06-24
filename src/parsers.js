// ============================================================
// 平台解析器 — 各通路訂單報表 → 統一中間格式
// 每個解析器吃 rows（xlsx 轉的物件陣列），吐 orders[]
// order = { platform, ref_no, order_date, contact, address, phone, email,
//           pay_method, note, store, pkg_count, tracking_no, total,
//           shipping_fee, product_text }
// product_text 用來比對商品對照表
// ============================================================

// 安全取值：從多個可能欄位名取第一個有值的
function pick(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
      return row[k];
    }
  }
  return '';
}

const num = (v) => {
  const n = parseFloat(String(v).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
};

// Excel 日期序號 → YYYY-MM-DD（xlsx 常把日期讀成數字）
function excelDate(v) {
  if (v === '' || v === null || v === undefined) return '';
  // 已是日期字串：統一換成 YYYY-MM-DD（斜線換橫線、截掉時間部分）
  if (typeof v === 'string' && /\d{4}[-/]\d/.test(v)) return v.replace(/\//g, '-').slice(0, 10);
  const n = parseFloat(v);
  if (isNaN(n) || n < 1 || n > 90000) return String(v);
  // Excel epoch: 1899-12-30
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  if (isNaN(d.getTime())) return String(v);
  const p = (x) => String(x).padStart(2, '0');
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`;
}

// ---------- 蝦皮 ----------
function parseShopee(rows) {
  return rows.map((r) => {
    const total = num(pick(r, ['買家總支付金額', '買家實際支付金額']));
    return {
      platform: '蝦皮',
      ref_no: pick(r, ['訂單編號']),
      order_date: excelDate(pick(r, ['訂單成立日期', '訂單日期'])),
      contact: pick(r, ['買家帳號', '買家用戶名稱']) || '蝦皮隱碼',
      address: '蝦皮隱碼',
      phone: '蝦皮隱碼',
      email: '蝦皮隱碼',
      pay_method: pick(r, ['付款方式']) || '蝦皮',
      note: '蝦皮',
      store: '蝦皮',
      pkg_count: 1,
      tracking_no: pick(r, ['包裹查詢號碼', '貨運單號', '追蹤編號']),
      total,
      shipping_fee: num(pick(r, ['買家支付的運費', '買家支付運費'])),
      product_text: pick(r, ['商品名稱', '商品選項名稱']),
      qty_platform: num(pick(r, ['數量'])) || 1,
    };
  }).filter((o) => o.ref_no);
}

// ---------- LINE商城 ----------
function parseLineMall(rows) {
  return rows.map((r) => {
    const dateRaw = pick(r, ['訂單下單日期', '訂單成立日期', '訂單日期'])
    return {
      platform: 'LINE商城',
      ref_no: pick(r, ['訂單編號']),
      order_date: typeof dateRaw === 'string' && dateRaw.length >= 10 ? dateRaw.slice(0, 10) : excelDate(dateRaw),
      contact: pick(r, ['收件人姓名', '收件者姓名']),
      address: pick(r, ['配送地址', '收件地址']),
      phone: pick(r, ['收件人聯絡電話', '收件者電話']),
      email: pick(r, ['收件人電子郵件', 'Email', '電子郵件']),
      pay_method: pick(r, ['付款方式']),
      note: 'LINE商城',
      store: 'LINE商城',
      pkg_count: 1,
      tracking_no: pick(r, ['物流單編號', '配送編號']),
      total: num(pick(r, ['總付款金額'])),
      shipping_fee: num(pick(r, ['合計運費', '運費'])),
      discount: num(pick(r, ['訂單優惠券折抵金額'])),
      product_text: pick(r, ['商品名稱']),
      qty_platform: num(pick(r, ['數量'])) || 1,
    };
  }).filter((o) => o.ref_no);
}

// ---------- 酷澎 ----------
function parseCoupang(rows) {
  return rows.map((r) => {
    return {
      platform: '酷澎',
      ref_no: pick(r, ['訂單編號']),
      order_date: excelDate(pick(r, ['訂購日期', '訂單日期'])),
      contact: pick(r, ['訂購人姓名', '收件人姓名']),  // 半遮蔽如 廖*姐，直接帶
      address: '酷澎隱碼',
      phone: '酷澎隱碼',
      email: '酷澎隱碼',
      pay_method: '酷澎',
      note: '酷澎訂單',
      store: '酷澎',
      pkg_count: 1,
      tracking_no: pick(r, ['寄件單號', '運送單號']),
      total: num(pick(r, ['銷售價格', '訂單金額'])),
      shipping_fee: num(pick(r, ['運費', '配送費'])),
      // 酷澎原始報表有公司商品代碼，可直接用；也保留商品名供比對
      product_text: pick(r, ['顯示產品名稱', '商品名稱']),
      direct_code: pick(r, ['公司商品代碼', '賣家商品代碼']),
      qty_platform: num(pick(r, ['數量'])) || 1,
    };
  }).filter((o) => o.ref_no);
}

// ---------- 官網（大筆） ----------
function parseOfficial(rows) {
  return rows.map((r) => {
    return {
      platform: '官網',
      ref_no: pick(r, ['參照編號', '訂單編號']),
      order_date: excelDate(pick(r, ['日期', '訂單日期'])),
      contact: pick(r, ['聯絡人', '訂購人', '收件人姓名']),
      address: pick(r, ['地址']),
      phone: pick(r, ['電話']),
      email: pick(r, ['Email', '電子郵件']),
      pay_method: pick(r, ['付款方式']),  // 線上刷卡 / Line Pay / Apple Pay
      note: pick(r, ['備註']) || '官網',
      store: pick(r, ['商店']) || 'HEHEYEN和和研',
      pkg_count: num(pick(r, ['件數'])) || 1,
      tracking_no: pick(r, ['託運單號']),
      total: num(pick(r, ['總計', '訂單金額'])),
      shipping_fee: num(pick(r, ['運費'])),
      product_text: pick(r, ['品項', '商品名稱', '商品資訊']),
      direct_code: pick(r, ['編碼']),
      qty_platform: num(pick(r, ['數量'])) || 1,
    };
  }).filter((o) => o.ref_no);
}

export const PARSERS = {
  '蝦皮': parseShopee,
  'LINE商城': parseLineMall,
  '酷澎': parseCoupang,
  '官網': parseOfficial,
};

// 自動偵測平台：依欄位特徵
export function detectPlatform(headers) {
  const h = headers.join('|');
  if (h.includes('買家總支付金額') || h.includes('買家帳號')) return '蝦皮';
  if (h.includes('配送地址') || h.includes('收件人聯絡電話')) return 'LINE商城';
  if (h.includes('訂購人姓名') || h.includes('顯示產品名稱')) return '酷澎';
  if (h.includes('參照編號') || h.includes('商店')) return '官網';
  return '';
}

export { num, pick, excelDate };
