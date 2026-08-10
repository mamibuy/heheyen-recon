// ============================================================
// 金流撥款明細解析器
// 每個解析器吃 xlsx rows，吐出 { key, key_type, fee, payable, actual_in, in_date, payout_date }[]
// key_type: 'ref_no' 直接比對 / 'ref_no_nodash' 去槓號後比對
// ============================================================
import { num, pick, excelDate } from './parsers.js'

export function stripDash(s) { return String(s).replace(/-/g, '') }

// 1. 蝦皮（逐筆）
// 新格式（蝦皮報表，有「錢包入帳金額」）：payable = 錢包入帳金額，fee = abs(sum(I:W))
// 舊格式（撥款明細，有「銀行實際收款金額」）：沿用原有邏輯
function parseShopeeRecon(rows) {
  if (!rows.length) return []
  const headers = Object.keys(rows[0])
  const isNewFormat = headers.includes('錢包入帳金額')

  if (isNewFormat) {
    // 手續費欄位：J/K/L/P/Q/R/U/V/W/X（排除M/N/O/S/T/Y）
    const FEE_COL_NAMES = new Set([
      '賣場商品促銷折扣', '退款金額', '蝦皮補貼金額',
      '蝦皮補助運費', '蝦皮代付運費', '退貨運費',
      'AMS推廣費用', '成交手續費', '其他服務費', '金流與系統處理費',
    ])
    // 代收付發票金額欄位：I~O（商品原價、賣場商品促銷折扣、退款金額、蝦皮補貼金額、賣家負擔優惠券、賣家負擔蝦幣回饋券、買家支付運費）
    const INV_COL_NAMES = new Set([
      '商品原價', '賣場商品促銷折扣', '退款金額', '蝦皮補貼金額',
      '賣家負擔優惠券', '賣家負擔蝦幣回饋券', '買家支付運費',
    ])
    const feeCols = headers.filter(h => FEE_COL_NAMES.has(h))
    const invCols  = headers.filter(h => INV_COL_NAMES.has(h))
    return rows.map(r => {
      const feeSum = feeCols.reduce((s, col) => {
        const v = Number(r[col])   // Number('2.50%') = NaN，避免費率欄被誤算
        return s + (isNaN(v) ? 0 : v)
      }, 0)
      const invSum = invCols.reduce((s, col) => {
        const v = Number(r[col])
        return s + (isNaN(v) ? 0 : v)
      }, 0)
      const payable = num(pick(r, ['錢包入帳金額']))  // Y欄（應入帳）
      const total   = num(pick(r, ['商品原價']))       // I欄（應收）
      return {
        key: String(pick(r, ['訂單編號'])).trim(),
        key_type: 'ref_no',
        fee: Math.abs(feeSum),
        payable,
        total,
        order_invoice_amount: Math.round(invSum * 100) / 100,
        actual_in: null,
        in_date: null,
        payout_date: null,
      }
    }).filter(r => r.key)
  }

  // 舊格式（撥款明細）—— Excel 欄位為負數，取絕對值後存正數
  return rows.map(r => {
    const feeRaw = num(pick(r, ['成交手續費'])) + num(pick(r, ['其他服務費'])) + num(pick(r, ['金流與系統處理費']))
    const fee = Math.abs(feeRaw)
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
      in_date: payout_date,
      payout_date,
    }
  }).filter(r => r.key)
}

// 3. LINE商城 - 信用卡（E-2 蘭新金流）
function parseLanxinRecon(rows) {
  return rows.map(r => {
    const fee = num(pick(r, ['主支付手續費']))
    const total = num(pick(r, ['訂單總金額']))
    const payout_date = (excelDate(pick(r, ['預計撥款日'])) || '').slice(0, 10) || null
    return {
      key: String(pick(r, ['商店訂單編號'])).trim(),
      key_type: 'ref_no',
      fee,
      payable: total - fee,
      actual_in: null,
      in_date: payout_date,
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
// C 欄「付款(退款)日期」→ order_date，但僅在訂單日期原本是空的時候才補
// （見 reconcile 的 order_date_fill_only）。已從官網出貨報表帶入正確下單日的訂單不覆蓋：
// 該欄在退款列填的是退款日，蓋上去會讓訂單日期完全失真。
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
      order_date: excelDate(pick(r, ['付款(退款)日期', '付款日期'])) || null,
      order_date_fill_only: true,
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

// 7. 兆豐福利網（雙檔：訂單明細報表 + 手續費報表）
//
// 應收一律用「商品金額 + 運費」，不可用「付款總金額」：
// 客戶若全額以福利金折抵，付款總金額會是 0（實付 0），但兆豐仍照訂單金額扣 6% 撥款給我們。
// 「訂單總金額」欄位也不可靠 —— 沒有動用福利金的訂單那欄是空的。
//
// 手續費報表沒有獨立的訂單編號欄，訂單編號埋在備註文字裡：
//   「[訂單編號：607020108] 【和和研】純LGG活菌益生菌-五入組-五入組 * 1」
//
// 兩份報表都沒有撥款日／實際入帳，實際入帳日由使用者對帳時自行填寫。
export function parseMegabankRecon(ordRows, feeRows) {
  const feeByRef = {}
  for (const r of (feeRows || [])) {
    const m = String(pick(r, ['備註']) || '').match(/訂單編號[：:]\s*([^\]\s]+)/)
    if (!m) continue   // 抽不到訂單編號的列（非交易手續費）直接略過
    const ref = m[1].trim()
    feeByRef[ref] = (feeByRef[ref] || 0) + num(pick(r, ['手續費']))
  }
  return (ordRows || []).map(r => {
    const key = String(pick(r, ['訂單編號']) || '').trim()   // 檔案裡帶前導空白，務必 trim
    const total = num(pick(r, ['商品金額'])) + num(pick(r, ['運費']))
    const fee = feeByRef[key] ?? 0
    return {
      key,
      key_type: 'ref_no',
      fee,
      payable: Math.round((total - fee) * 100) / 100,
      total,
      order_date: excelDate(pick(r, ['訂單時間', '訂單日期'])) || null,
      actual_in: null,
      in_date: null,
      payout_date: null,
    }
  }).filter(r => r.key)
}

// 兆豐手續費固定為訂單金額的 6%。算出來不符代表兩份報表對不起來
// （最常見：手續費報表的期間沒涵蓋到某些訂單），回傳待人工確認的清單。
export const MEGABANK_FEE_RATE = 0.06

export function megabankRateWarnings(parsed) {
  const noFee = [], oddRate = []
  for (const r of parsed) {
    if (!r.total) continue
    if (!r.fee) { noFee.push(r.key); continue }
    const rate = r.fee / r.total
    if (Math.abs(rate - MEGABANK_FEE_RATE) > 0.0005) {
      oddRate.push({ key: r.key, rate: (rate * 100).toFixed(2), fee: r.fee, total: r.total })
    }
  }
  return (noFee.length || oddRate.length) ? { noFee, oddRate } : null
}

// 兆豐雙檔的欄位檢查
export function checkMegabankColumns(ordRows, feeRows) {
  const bad = []
  const ho = new Set(Object.keys(ordRows?.[0] || {}).map(h => String(h).trim()))
  const hf = new Set(Object.keys(feeRows?.[0] || {}).map(h => String(h).trim()))
  if (!ho.has('訂單編號')) bad.push('訂單明細報表缺少「訂單編號」')
  if (!ho.has('商品金額')) bad.push('訂單明細報表缺少「商品金額」')
  if (!hf.has('手續費')) bad.push('手續費報表缺少「手續費」')
  if (!hf.has('備註')) bad.push('手續費報表缺少「備註」（訂單編號來源）')
  return bad.length ? bad : null
}

// ============================================================
// 上傳檔案格式檢查
// 每條金流列出 parser 真正會讀的關鍵欄位；每組至少要命中一個名稱。
// 目的：擋掉上傳錯報表的情況 —— parser 找不到欄位時 num()/pick() 一路回 0，
// 再套上寫死的固定費用（如酷澎 4 元），會產出「應收 0 / 應入帳 -4」這種
// 看似成功、實則垃圾的結果，還會被自動建檔寫進 DB。
// （實例：酷澎「營業稅明細 VAT History Report」被當成「出帳明細」上傳）
// ============================================================
const REQUIRED_COLS = {
  shopee: [
    { label: '訂單編號', names: ['訂單編號'] },
    { label: '手續費或入帳金額', names: ['成交手續費', '金流與系統處理費', '錢包入帳金額', '銀行實際收款金額'] },
  ],
  linepay: [
    { label: '訂單號碼', names: ['訂單號碼', '訂單編號'] },
    { label: '手續費合計', names: ['手續費合計（含營業稅）', '手續費合計'] },
  ],
  lanxin: [
    { label: '商店訂單編號', names: ['商店訂單編號'] },
    { label: '主支付手續費', names: ['主支付手續費'] },
  ],
  coupang: [
    { label: '訂單編號', names: ['訂單編號'] },
    { label: '手續費總額', names: ['手續費總額', '手續費'] },
    { label: '訂單金額', names: ['買家總支付', '銷售價格', '訂單金額'] },
  ],
  payuni_cc: [
    { label: '商店訂單編號', names: ['商店訂單編號'] },
    { label: '手續費', names: ['手續費'] },
  ],
  payuni_linepay: [
    { label: '商店訂單編號', names: ['商店訂單編號'] },
    { label: '交易處理費或對應碼', names: ['交易處理費', '支付方式對應碼'] },
  ],
}

// 回傳 null 代表通過；否則回傳缺少欄位的描述陣列
export function checkReconColumns(gateway, rows) {
  const spec = REQUIRED_COLS[gateway]
  if (!spec || !rows || !rows.length) return null
  const headers = new Set(Object.keys(rows[0]).map(h => String(h).trim()))
  const missing = spec.filter(g => !g.names.some(n => headers.has(n)))
  return missing.length ? missing.map(g => `${g.label}（${g.names.join('／')}）`) : null
}

// 官網 LINE Pay 雙檔：只檢查各自的比對鑰匙，欄位名變異較多故不做嚴格檢查
export function checkDualReconColumns(d1rows, d2rows) {
  const bad = []
  const h1 = new Set(Object.keys(d1rows?.[0] || {}).map(h => String(h).trim()))
  const h2 = new Set(Object.keys(d2rows?.[0] || {}).map(h => String(h).trim()))
  if (!['交易號碼', '訂單號碼'].some(n => h1.has(n))) bad.push('D-1 缺少「交易號碼」')
  if (!h2.has('商店訂單編號')) bad.push('D-2 缺少「商店訂單編號」')
  return bad.length ? bad : null
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
  megabank: '兆豐福利網',
}

export function detectGateway(headers) {
  const h = headers.join('|')
  // 兆豐擺最前面：兩份報表的特徵欄位（福利金）其他金流都沒有，先攔下避免被後面的規則誤判
  if (h.includes('福利金折扣') || (h.includes('福利金') && h.includes('類別') && h.includes('備註'))) return 'megabank'
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
  // D-2 的支付方式對應碼為數字，超過 JS 安全整數後精度遺失（後幾位變 0）
  // D-1 的交易號碼為字串（精確）。兩者前 15 位相同，用前綴比對。
  const feeByTx = {}
  const payableByTx = {}
  for (const r of d1rows) {
    const tx = String(pick(r, ['交易號碼', '訂單號碼'])).trim()
    if (tx && tx !== '-') {
      const key = tx.slice(0, 15)
      feeByTx[key] = num(pick(r, ['手續費合計', '手續費']))
      payableByTx[key] = num(pick(r, ['排定的各項目撥款']))
    }
  }
  // 從 D-2 組裝每筆訂單
  return d2rows.map(r => {
    const key = stripDash(String(pick(r, ['商店訂單編號'])).trim())
    const txCode = String(pick(r, ['支付方式對應碼'])).trim()
    const txKey = txCode.slice(0, 15)
    const total = num(pick(r, ['付款金額', '交易金額']))
    const fee = feeByTx[txKey] ?? 0
    const payable = payableByTx[txKey] ?? (total - fee)
    const txFee = num(pick(r, ['交易處理費']))
    const in_date = excelDate(pick(r, ['入帳日期', '撥款日期'])) || null
    return { key, key_type: 'ref_no_nodash', fee, payable, actual_in: null, in_date, payout_date: in_date, tx_code: txCode || null, tx_fee: txFee }
  }).filter(r => r.key)
}
