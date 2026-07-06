// ============================================================
// 金流對帳回填邏輯
// 用解析出的撥款明細比對 shipping_orders，回填手續費/應入帳/入帳日/狀態
// ============================================================
import { stripDash } from './recon_parsers.js'

// 發票核對：gateway → DB 的 platform 欄位
const GATEWAY_PLATFORM = {
  shopee: '蝦皮', linepay: 'LINE商城', lanxin: 'LINE商城',
  coupang: '酷澎', payuni_cc: '官網', payuni_linepay: '官網',
}

// 同一平台下用 pay_method 區分金流（LINE商城/官網各有兩條）
function matchesGateway(order, gateway) {
  const pm = String(order.pay_method || '').toLowerCase().replace('_', ' ')
  if (gateway === 'linepay') return pm.includes('line')
  if (gateway === 'lanxin') return !pm.includes('line')
  if (gateway === 'payuni_linepay') return pm.includes('line pay') || pm.includes('linepay')
  if (gateway === 'payuni_cc') return !(pm.includes('line pay') || pm.includes('linepay'))
  return true
}

// 撈出某金流、某期間的訂單並加總 fee_total，供使用者確認後套用
export async function previewInvoice(supabase, { gateway, dateFrom, dateTo }) {
  const platform = GATEWAY_PLATFORM[gateway]
  if (!platform) throw new Error('未知金流：' + gateway)
  const { data, error } = await supabase
    .from('shipping_orders')
    .select('id,ref_no,fee_total,in_date,order_date,pay_method')
    .eq('platform', platform)
  if (error) throw new Error(error.message)
  const orders = (data || []).filter(o => {
    if (!matchesGateway(o, gateway)) return false
    const d = o.in_date || o.order_date || ''
    if (dateFrom && d < dateFrom) return false
    if (dateTo && d > dateTo) return false
    return true
  })
  const feeSum = Math.round(orders.reduce((s, o) => s + (o.fee_total || 0), 0) * 100) / 100
  return { orders, feeSum }
}

// 撈出指定子分類的訂單（按 platform + pay_method 篩選）
export async function loadGatewayOrders(supabase, gateway) {
  const platform = GATEWAY_PLATFORM[gateway]
  if (!platform) return []
  const { data } = await supabase
    .from('shipping_orders')
    .select('id,ref_no,sa_no,platform,total,fee_total,payable,actual_in,in_date,bank_deposit,order_date,pay_method,note,recon_status,invoice_check,fee_invoice_no,fee_invoice_date,fee_invoice_amount,account_fee_note,order_invoice_no,order_invoice_date,order_invoice_amount,tx_code,tx_fee,tx_fee_invoice_no,fee_invoice_note,tx_fee_invoice_note,fee_invoice_pdf_url,tx_fee_invoice_pdf_url')
    .eq('platform', platform)
    .order('created_at', { ascending: false })
  return (data || []).filter(o => matchesGateway(o, gateway))
}

// 把進項發票資訊寫入這批訂單，並標記 invoice_check
export async function applyInvoice(supabase, { orderIds, invoiceNo, invoiceDate, invoiceAmount, isMatch }) {
  const { error } = await supabase
    .from('shipping_orders')
    .update({
      fee_invoice_no: invoiceNo || null,
      fee_invoice_date: invoiceDate || null,
      fee_invoice_amount: invoiceAmount || null,
      invoice_check: isMatch ? '相符' : '有差異',
    })
    .in('id', orderIds)
  if (error) throw new Error(error.message)
  return orderIds.length
}

export async function reconcile(supabase, gateway, parsedRows) {
  const { data: orders, error } = await supabase
    .from('shipping_orders')
    .select('id,ref_no,total,recon_status')
  if (error) throw new Error('撈取訂單失敗：' + error.message)

  // 建兩張查找表：原始 ref_no 和去槓號後的 ref_no
  const byRefNo = {}
  const byRefNoDash = {}
  for (const o of (orders || [])) {
    byRefNo[String(o.ref_no)] = o
    byRefNoDash[stripDash(String(o.ref_no))] = o
  }

  const matched = [], unmatched = [], updated = []
  let feeTotal = 0, payableTotal = 0

  for (const row of parsedRows) {
    const order = row.key_type === 'ref_no_nodash'
      ? byRefNoDash[row.key]
      : byRefNo[row.key]

    if (!order) {
      unmatched.push(row.key)
      continue
    }
    matched.push(row.key)

    const fee_total = row.fee != null ? row.fee : 0
    const payable = row.payable != null ? row.payable : 0

    // 只有報表本身含實際入帳金額（蝦皮）才回填 actual_in / in_date / 狀態改已入帳
    // 其他金流（CC / LINE Pay）入帳日只能從銀行對帳單確認後寫入，不從撥款報表覆蓋
    const updates = { fee_total, payable }
    // 蝦皮進帳報表提供正確應收（商品原價 - 賣家自負折扣），一併更新 total
    if (row.total != null) updates.total = row.total
    if (row.actual_in != null) {
      updates.actual_in = row.actual_in
      updates.in_date = row.in_date || null
      updates.recon_status = '已入帳'
    } else if (order.recon_status !== '已入帳' && order.recon_status !== '已對帳') {
      updates.recon_status = '平台已結算'
    }
    if (row.tx_code !== undefined) updates.tx_code = row.tx_code ?? null
    if (row.tx_fee !== undefined) updates.tx_fee = row.tx_fee ?? null

    const { error: updateError } = await supabase
      .from('shipping_orders')
      .update(updates)
      .eq('id', order.id)

    if (!updateError) {
      updated.push(row.key)
      feeTotal += fee_total
      payableTotal += payable
    }
  }

  return {
    matched: matched.length,
    unmatched,
    updated: updated.length,
    feeTotal: Math.round(feeTotal * 100) / 100,
    payableTotal: Math.round(payableTotal * 100) / 100,
  }
}
