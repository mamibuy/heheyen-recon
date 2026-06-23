// ============================================================
// 金流撥款明細解析器
// 每個解析器吃 xlsx rows，吐出 { key, key_type, fee, payable, actual_in, in_date, payout_date }[]
// key_type: 'ref_no' 直接比對 / 'ref_no_nodash' 去槓號後比對
// ============================================================
import { num, pick, excelDate } from './parsers.js'

export function stripDash(s) { return String(s).replace(/-/g, '') }

// 1. 蝦皮（逐筆）
function parseShopeeRecon(rows) {
  return rows.map(r => {
    const fee = num(pick(r, ['成交手續費'])) + num(pick(r, ['其他服務費'])) + num(pick(r, ['金流與系統處理費']))
    const total = num(pick(r, ['買家實際支付金額', '買家總支付金額', '訂單金額']))
    const actual_in = num(pick(r, ['銀行實際收款金額'])) || null
    const in_date = excelDate(pick(r, ['銀行實際收款日'])) || null
    return {
      key: String(pick(r, ['訂單編號'])).trim(),
      key_type: 'ref_no',
      fee,
      payable: total - fee,
      actual_in,
      in_date,
      payout_date: in_date,
    }
  }).filter(r => r.key)
}

// 2. LINE商城 - LINE Pay（E-1）
function parseLinePayRecon(rows) {
  return rows.map(r => {
    const fee = num(pick(r, ['手續費合計（含營業稅）', '手續費合計']))
    const total = num(pick(r, ['支付總額']))
    const payout_date = excelDate(pick(r, ['預計撥款日'])) || null
    return {
      key: String(pick(r, ['訂單號碼', '訂單編號'])).trim(),
      key_type: 'ref_no',
      fee,
      payable: total - fee,
      actual_in: null,
      in_date: null,
      payout_date,
    }
  }).filter(r => r.key)
}

// 3. LINE商城 - 信用卡（E-2 蘭新金流）
function parseLanxinRecon(rows) {
  return rows.map(r => {
    const fee = num(pick(r, ['主支付手續費']))
    const total = num(pick(r, ['訂單總金額']))
    const payout_date = excelDate(pick(r, ['預計撥款日'])) || null
    return {
      key: String(pick(r, ['商店訂單編號'])).trim(),
      key_type: 'ref_no',
      fee,
      payable: total - fee,
      actual_in: null,
      in_date: null,
      payout_date,
    }
  }).filter(r => r.key)
}

// 4. 酷澎（合併商品列 + DELIVERY_FEE列）
function parseCoupangRecon(rows) {
  const map = {}
  for (const r of rows) {
    const key = String(pick(r, ['訂單編號'])).trim()
    if (!key) continue
    if (!map[key]) map[key] = { key, fee: 0, delivery: 0, total: 0, actual_in: null, in_date: null }
    const o = map[key]
    const isDelivery = String(pick(r, ['商品名稱', '品項名稱', ''])).toUpperCase().includes('DELIVERY_FEE')
    const amt = num(pick(r, ['買家總支付', '銷售價格', '訂單金額']))
    if (isDelivery) {
      o.delivery += amt
    } else {
      o.fee += num(pick(r, ['手續費總額', '手續費']))
      o.total += amt
    }
    const ai = num(pick(r, ['銀行實際收款金額']))
    if (ai) o.actual_in = (o.actual_in || 0) + ai
    const id = excelDate(pick(r, ['銀行實際收款日']))
    if (id) o.in_date = id
  }
  return Object.values(map).map(o => ({
    key: o.key,
    key_type: 'ref_no',
    fee: o.fee + 4,  // 固定其他服務費 4 元
    payable: o.total - o.fee - 4 - o.delivery,
    actual_in: o.actual_in || null,
    in_date: o.in_date || null,
    payout_date: null,
  }))
}

// 5. 官網信用卡（C PayUni 入帳表）— 去槓號比對
function parsePayuniCCRecon(rows) {
  return rows.map(r => {
    const fee = Math.abs(num(pick(r, ['手續費'])))
    const total = num(pick(r, ['收款金額']))
    const payable = num(pick(r, ['入帳金額'])) || (total - fee)
    const in_date = excelDate(pick(r, ['入帳日期'])) || null
    return {
      key: stripDash(String(pick(r, ['商店訂單編號'])).trim()),
      key_type: 'ref_no_nodash',
      fee,
      payable,
      actual_in: payable,
      in_date,
      payout_date: in_date,
    }
  }).filter(r => r.key)
}

// 6. 官網 LINE Pay（D-2 PayUni 電子錢包）— 去槓號比對
// D-1 LINE Pay 手續費需另行上傳勾稽；此處先以 D-2 付款金額計算應入帳
function parseOfficialLinePayRecon(rows) {
  return rows.map(r => {
    const total = num(pick(r, ['付款金額', '交易金額']))
    const fee = num(pick(r, ['手續費合計', '手續費']))
    const in_date = excelDate(pick(r, ['入帳日期', '撥款日期'])) || null
    return {
      key: stripDash(String(pick(r, ['商店訂單編號'])).trim()),
      key_type: 'ref_no_nodash',
      fee,
      payable: total - fee,
      actual_in: null,
      in_date,
      payout_date: in_date,
    }
  }).filter(r => r.key)
}

export const RECON_PARSERS = {
  shopee: parseShopeeRecon,
  linepay: parseLinePayRecon,
  lanxin: parseLanxinRecon,
  coupang: parseCoupangRecon,
  payuni_cc: parsePayuniCCRecon,
  payuni_linepay: parseOfficialLinePayRecon,
}

export const GATEWAY_LABELS = {
  shopee: '蝦皮',
  linepay: 'LINE商城-LinePay',
  lanxin: 'LINE商城-信用卡',
  coupang: '酷澎',
  payuni_cc: '官網-信用卡',
  payuni_linepay: '官網-LinePay',
}

export function detectGateway(headers) {
  const h = headers.join('|')
  if (h.includes('成交手續費') || h.includes('金流與系統處理費')) return 'shopee'
  if (h.includes('手續費合計') && h.includes('LINE Pay優惠')) return 'linepay'
  if (h.includes('主支付手續費') || h.includes('藍新金流交易序號')) return 'lanxin'
  if (h.includes('手續費總額') && (h.includes('DELIVERY_FEE') || h.includes('捆綁配送編號') || h.includes('應付金額'))) return 'coupang'
  if (h.includes('uni序號') || (h.includes('入帳金額') && h.includes('商店訂單編號'))) return 'payuni_cc'
  if (h.includes('交易處理費') || h.includes('支付方式對應碼')) return 'payuni_linepay'
  return ''
}

// 官網 LINE Pay 雙檔解析：D-1（新LinePay）+ D-2（PayUni電子錢包）
// D-1 提供 交易號碼 → 手續費合計
// D-2 提供 商店訂單編號（去槓）→ ref_no、支付方式對應碼 → D-1 交易號碼（勾稽取費用）
export function parseOfficialLinePayReconDual(d1rows, d2rows) {
  // 建 D-1 查找表：交易號碼 → 手續費
  const feeByTx = {}
  for (const r of d1rows) {
    const tx = String(pick(r, ['交易號碼', '訂單號碼'])).trim()
    if (tx) feeByTx[tx] = num(pick(r, ['手續費合計', '手續費']))
  }
  // 從 D-2 組裝每筆訂單
  return d2rows.map(r => {
    const key = stripDash(String(pick(r, ['商店訂單編號'])).trim())
    const txCode = String(pick(r, ['支付方式對應碼'])).trim()
    const total = num(pick(r, ['付款金額', '交易金額']))
    const fee = feeByTx[txCode] ?? 0
    const in_date = excelDate(pick(r, ['入帳日期', '撥款日期'])) || null
    return { key, key_type: 'ref_no_nodash', fee, payable: total - fee, actual_in: null, in_date, payout_date: in_date }
  }).filter(r => r.key)
}
