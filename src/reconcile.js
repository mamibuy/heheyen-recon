// ============================================================
// 金流對帳回填邏輯
// 用解析出的撥款明細比對 shipping_orders，回填手續費/應入帳/入帳日/狀態
// ============================================================
import { stripDash } from './recon_parsers.js'

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
    const actual_in = row.actual_in != null ? row.actual_in : null
    const in_date = row.in_date || null
    const recon_status = actual_in != null ? '已入帳' : '平台已結算'

    const { error: updateError } = await supabase
      .from('shipping_orders')
      .update({ fee_total, payable, actual_in, in_date, recon_status })
      .eq('id', order.id)

    if (!updateError) updated.push(row.key)
  }

  return { matched: matched.length, unmatched, updated: updated.length }
}
