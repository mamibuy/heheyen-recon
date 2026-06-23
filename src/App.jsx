import React, { useState, useEffect, useRef } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import { PARSERS, detectPlatform } from './parsers.js'
import { buildBlocks } from './transform.js'
import { RECON_PARSERS, parseOfficialLinePayReconDual } from './recon_parsers.js'
import { reconcile, previewInvoice, applyInvoice, loadGatewayOrders } from './reconcile.js'

// ====== Supabase（沿用 Mamibuy 專案）======
const supabase = createClient(
  'https://geirbvjkwsewglvvrfmg.supabase.co',
  'sb_publishable_yDgLU7V2PcL_2QmrQkxo2w_WZGEbP63'
)

const C = {
  bg: '#f7f8fa', card: '#ffffff', line: '#e3e6ea', ink: '#1f2933',
  sub: '#6b7682', brand: '#1d7a5f', brandBg: '#e8f4ef',
  warn: '#b4541a', warnBg: '#fbeee2', danger: '#c0392b',
}

const PLATFORMS = ['蝦皮', 'LINE商城', '酷澎', '官網']

const GATEWAY_LIST = [
  { key: 'coupang',        label: '酷澎' },
  { key: 'shopee',         label: '蝦皮' },
  { key: 'payuni_cc',      label: '官網 › 信用卡' },
  { key: 'payuni_linepay', label: '官網 › LINE Pay', twoFile: true },
  { key: 'linepay',        label: 'LINE商城 › LINE Pay' },
  { key: 'lanxin',         label: 'LINE商城 › 信用卡' },
]

export default function App() {
  const [tab, setTab] = useState('convert')
  return (
    <div style={{ minHeight: '100vh', background: C.bg, color: C.ink,
      fontFamily: '-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC","Microsoft JhengHei",sans-serif' }}>
      <header style={{ background: C.card, borderBottom: `1px solid ${C.line}`, padding: '14px 20px' }}>
        <div style={{ maxWidth: 1100, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 16 }}>
          <strong style={{ fontSize: 17 }}>和和研 · 電商出貨彙整</strong>
          <nav style={{ display: 'flex', gap: 4, marginLeft: 8 }}>
            <TabBtn active={tab === 'convert'} onClick={() => setTab('convert')}>出貨轉換</TabBtn>
            <TabBtn active={tab === 'mapping'} onClick={() => setTab('mapping')}>商品對照表</TabBtn>
            <TabBtn active={tab === 'recon'} onClick={() => setTab('recon')}>金流對帳</TabBtn>
          </nav>
        </div>
      </header>
      <main style={{ maxWidth: 1100, margin: '0 auto', padding: 20 }}>
        {tab === 'convert' && <ConvertPage />}
        {tab === 'mapping' && <MappingPage />}
        {tab === 'recon' && <ReconPage />}
      </main>
    </div>
  )
}

function TabBtn({ active, onClick, children }) {
  return (
    <button onClick={onClick} style={{
      padding: '6px 14px', borderRadius: 8, border: 'none', cursor: 'pointer',
      fontSize: 14, fontWeight: active ? 700 : 500,
      background: active ? C.brandBg : 'transparent', color: active ? C.brand : C.sub,
    }}>{children}</button>
  )
}

// ============================================================
// 出貨轉換頁
// ============================================================
function ConvertPage() {
  const [mapping, setMapping] = useState([])
  const [platform, setPlatform] = useState('')
  const [orders, setOrders] = useState([])
  const [blocks, setBlocks] = useState(null)
  const [fileName, setFileName] = useState('')
  const [msg, setMsg] = useState('')
  const fileRef = useRef(null)

  useEffect(() => { loadMapping() }, [])
  async function loadMapping() {
    const { data } = await supabase.from('product_mapping').select('*')
    setMapping(data || [])
  }

  function handleFile(e) {
    const f = e.target.files?.[0]
    if (!f) return
    setFileName(f.name)
    const reader = new FileReader()
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' })
      const headers = rows.length ? Object.keys(rows[0]) : []
      const detected = detectPlatform(headers)
      const usePlatform = platform || detected
      if (!usePlatform) {
        setMsg('無法自動辨識平台，請手動從上方選擇平台後重新上傳。')
        return
      }
      setPlatform(usePlatform)
      const parsed = PARSERS[usePlatform](rows)
      setOrders(parsed)
      const b = buildBlocks(parsed, mapping)
      setBlocks(b)
      setMsg(`已辨識為「${usePlatform}」，共 ${parsed.length} 筆訂單。`)
    }
    reader.readAsArrayBuffer(f)
  }

  function download() {
    if (!blocks) return
    const wb = XLSX.utils.book_new()
    // 訂單總表（二）+ 商品彙總 放同一張表（上下排）
    const ws1 = XLSX.utils.json_to_sheet(blocks.block1)
    // 在 block1 下方空兩列接 block2
    XLSX.utils.sheet_add_json(ws1, blocks.block2, { origin: -1, skipHeader: false })
    XLSX.utils.book_append_sheet(wb, ws1, '訂單總表（二）')
    const ws3 = XLSX.utils.json_to_sheet(blocks.block3)
    XLSX.utils.book_append_sheet(wb, ws3, '全部物流資料')
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '')
    XLSX.writeFile(wb, `和和研出貨明細_${platform}_${today}.xlsx`)
  }

  async function saveToDb() {
    if (!orders.length) return
    const rows = orders.map((o) => {
      const items = blocks.block1.filter((b) => b.參照編號 === o.ref_no || b.編號 === '')
      return {
        platform: o.platform, ref_no: o.ref_no, order_date: String(o.order_date || ''),
        contact: o.contact, address: o.address, phone: String(o.phone || ''), email: o.email,
        pay_method: o.pay_method, note: o.note, store: o.store, pkg_count: o.pkg_count || 1,
        tracking_no: String(o.tracking_no || ''), total: o.total || 0, shipping_fee: o.shipping_fee || 0, discount: o.discount ?? null,
        recon_status: '已出貨',
      }
    })
    const { error } = await supabase.from('shipping_orders').upsert(rows, { onConflict: 'platform,ref_no' })
    setMsg(error ? `存檔失敗：${error.message}` : `已存入 ${rows.length} 筆到資料庫。`)
  }

  const unmatchedCount = blocks?.unmatched?.length || 0

  return (
    <div>
      <Card>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <label style={{ fontSize: 13, color: C.sub }}>平台（留空自動辨識）</label>
          <select value={platform} onChange={(e) => setPlatform(e.target.value)}
            style={inp}>
            <option value="">自動辨識</option>
            {PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile}
            style={{ display: 'none' }} />
          <button onClick={() => fileRef.current.click()} style={btnPrimary}>上傳平台報表</button>
          {fileName && <span style={{ fontSize: 13, color: C.sub }}>{fileName}</span>}
        </div>
        {msg && <p style={{ marginTop: 12, marginBottom: 0, fontSize: 13,
          color: msg.includes('失敗') || msg.includes('無法') ? C.danger : C.brand }}>{msg}</p>}
      </Card>

      {unmatchedCount > 0 && (
        <Card style={{ background: C.warnBg, borderColor: C.warn }}>
          <strong style={{ color: C.warn }}>⚠ {unmatchedCount} 筆商品未對應</strong>
          <p style={{ fontSize: 13, color: C.ink, margin: '6px 0 0' }}>
            這些商品在對照表找不到規則，編碼會空白。請到「商品對照表」分頁新增規則後重新上傳。
          </p>
          <ul style={{ fontSize: 12, color: C.sub, margin: '8px 0 0', paddingLeft: 18 }}>
            {blocks.unmatched.slice(0, 5).map((u, i) => (
              <li key={i}>{u.ref_no}：{u.text}</li>
            ))}
          </ul>
        </Card>
      )}

      {blocks && (
        <Card>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <strong>預覽 · 訂單總表（二）</strong>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={saveToDb} style={btnGhost}>存入資料庫</button>
              <button onClick={download} style={btnPrimary}>下載 Excel</button>
            </div>
          </div>
          <PreviewTable rows={blocks.block1.slice(0, 12)} />
          <p style={{ fontSize: 12, color: C.sub, marginTop: 8 }}>
            僅顯示前 12 列；下載的 Excel 含完整三區塊（訂單總表、商品彙總、全部物流資料）。
          </p>
        </Card>
      )}
    </div>
  )
}

function PreviewTable({ rows }) {
  if (!rows.length) return null
  const cols = ['參照編號', '聯絡人', '付款方式', '編碼', '品項', '數量', '含稅單價', '總計']
  return (
    <div style={{ overflowX: 'auto' }}>
      <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 13 }}>
        <thead>
          <tr>{cols.map((c) => <th key={c} style={th}>{c}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ background: r.編號 === '' ? '#fafbfc' : '#fff' }}>
              {cols.map((c) => <td key={c} style={td}>{String(r[c] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ============================================================
// 商品對照表後台
// ============================================================
function MappingPage() {
  const [rows, setRows] = useState([])
  const [loading, setLoading] = useState(true)
  const [filter, setFilter] = useState('')
  const [editing, setEditing] = useState(null)

  useEffect(() => { load() }, [])
  async function load() {
    setLoading(true)
    const { data } = await supabase.from('product_mapping').select('*').order('platform').order('group_key').order('sort_order')
    setRows(data || [])
    setLoading(false)
  }

  async function saveRow(r) {
    const payload = { ...r }
    delete payload.id; delete payload.created_at; delete payload.updated_at
    payload.qty = parseFloat(payload.qty) || 0
    payload.sort_order = parseInt(payload.sort_order) || 0
    const { error } = r.id
      ? await supabase.from('product_mapping').update(payload).eq('id', r.id)
      : await supabase.from('product_mapping').insert([payload])
    if (error) { alert('儲存失敗：' + error.message); return }
    setEditing(null); load()
  }

  async function delRow(id) {
    if (!confirm('確定刪除這列？')) return
    await supabase.from('product_mapping').delete().eq('id', id)
    load()
  }

  function newRow() {
    setEditing({ platform: '蝦皮', match_text: '', group_key: '', role: 'main', code: '', item_name: '', qty: 1, sort_order: 0, active: true })
  }

  const shown = rows.filter((r) =>
    !filter || r.platform.includes(filter) || (r.match_text || '').includes(filter) || (r.code || '').includes(filter))

  return (
    <div>
      <Card>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <input placeholder="搜尋平台 / 商品 / 編碼" value={filter}
              onChange={(e) => setFilter(e.target.value)} style={{ ...inp, width: 220 }} />
            <span style={{ fontSize: 13, color: C.sub }}>{shown.length} 列</span>
          </div>
          <button onClick={newRow} style={btnPrimary}>+ 新增規則</button>
        </div>
      </Card>

      <Card>
        {loading ? <p style={{ color: C.sub }}>載入中…</p> : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 13 }}>
              <thead>
                <tr>
                  {['平台', '商品識別字串', '綁定鍵', '角色', '編碼', '品項', '數量', '排序', ''].map((c) =>
                    <th key={c} style={th}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {shown.map((r) => (
                  <tr key={r.id}>
                    <td style={td}>{r.platform}</td>
                    <td style={td}>{r.match_text}</td>
                    <td style={{ ...td, color: C.sub }}>{r.group_key}</td>
                    <td style={td}>{r.role === 'main' ? '主商品' : '贈品'}</td>
                    <td style={{ ...td, fontFamily: 'monospace' }}>{r.code}</td>
                    <td style={td}>{r.item_name}</td>
                    <td style={{ ...td, textAlign: 'right' }}>{r.qty}</td>
                    <td style={{ ...td, textAlign: 'right', color: C.sub }}>{r.sort_order}</td>
                    <td style={td}>
                      <button onClick={() => setEditing(r)} style={miniBtn}>編輯</button>
                      <button onClick={() => delRow(r.id)} style={{ ...miniBtn, color: C.danger }}>刪除</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {editing && <EditModal row={editing} onClose={() => setEditing(null)} onSave={saveRow} />}
    </div>
  )
}

// ============================================================
// 金流對帳頁 — 六個子分類工作區
// ============================================================
function ReconPage() {
  const [activeGateway, setActiveGateway] = useState('coupang')
  const [txRows, setTxRows] = useState(null)
  const [txFileName, setTxFileName] = useState('')
  const [txMsg, setTxMsg] = useState('')
  const [txResult, setTxResult] = useState(null)
  const txFileRef = useRef(null)

  function readTxFile(e) {
    const f = e.target.files?.[0]
    if (!f) return
    setTxFileName(f.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      setTxRows(XLSX.utils.sheet_to_json(ws, { defval: '' }))
    }
    reader.readAsArrayBuffer(f)
  }

  async function handleTianxin() {
    if (!txRows) { setTxMsg('請先上傳檔案'); return }
    setTxMsg('比對中…'); setTxResult(null)

    // 每個客戶訂單只取第一個 SA 單號（一筆訂單可能有多列商品）
    const map = {}
    for (const r of txRows) {
      const ref = String(r['客戶訂單'] || '').trim()
      const sa = String(r['單號'] || '').trim()
      if (ref && sa.startsWith('SA') && !map[ref]) map[ref] = sa
    }
    const pairs = Object.entries(map)
    if (!pairs.length) { setTxMsg('找不到 SA 開頭的單號，請確認欄位名稱'); return }

    const { data: allOrders, error } = await supabase.from('shipping_orders').select('id,ref_no')
    if (error) { setTxMsg('錯誤：' + error.message); return }

    const byRef = {}
    for (const o of (allOrders || [])) byRef[o.ref_no] = o.id

    const unmatched = []
    let updated = 0
    for (const [ref, sa] of pairs) {
      const id = byRef[ref]
      if (!id) { unmatched.push(ref); continue }
      const { error: ue } = await supabase.from('shipping_orders').update({ sa_no: sa }).eq('id', id)
      if (!ue) updated++
    }

    setTxResult({ total: pairs.length, updated, unmatched })
    setTxMsg(`${pairs.length} 筆訂單，回填 ${updated} 筆銷貨單號，未對應 ${unmatched.length} 筆`)
  }

  return (
    <div>
      <Card>
        <strong style={{ fontSize: 14 }}>上傳天心銷貨單（回填銷貨單號）</strong>
        <p style={{ fontSize: 12, color: C.sub, margin: '4px 0 10px' }}>
          比對「客戶訂單」與平台訂單編號，將 SA 單號寫入銷貨單號欄位（適用所有平台）
        </p>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <input ref={txFileRef} type="file" accept=".xlsx,.xls" onChange={readTxFile} style={{ display: 'none' }} />
          <button onClick={() => txFileRef.current.click()} style={btnGhost}>{txFileName || '選擇天心銷貨單'}</button>
          {txRows && <span style={{ fontSize: 12, color: C.brand }}>✓ {txRows.length} 列</span>}
          <button onClick={handleTianxin} style={btnPrimary}>比對回填</button>
        </div>
        {txMsg && (
          <p style={{ marginTop: 8, marginBottom: 0, fontSize: 13,
            color: txMsg.includes('錯誤') || txMsg.includes('找不到') ? C.danger : C.brand }}>
            {txMsg}
          </p>
        )}
        {txResult?.unmatched?.length > 0 && (
          <p style={{ marginTop: 4, marginBottom: 0, fontSize: 12, color: C.warn }}>
            未對應：{txResult.unmatched.slice(0, 8).join('、')}
            {txResult.unmatched.length > 8 && `…等 ${txResult.unmatched.length} 筆`}
          </p>
        )}
      </Card>

      <Card style={{ marginBottom: 0, borderRadius: '12px 12px 0 0' }}>
        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
          {GATEWAY_LIST.map(g => (
            <TabBtn key={g.key} active={activeGateway === g.key} onClick={() => setActiveGateway(g.key)}>
              {g.label}
            </TabBtn>
          ))}
        </div>
      </Card>
      <GatewayWorkspace gateway={activeGateway} key={activeGateway} />
    </div>
  )
}

function GatewayWorkspace({ gateway }) {
  const gwInfo = GATEWAY_LIST.find(g => g.key === gateway) || {}
  const isTwoFile = !!gwInfo.twoFile
  const isLinePayOfficial = gateway === 'payuni_linepay'
  const STATUSES = ['待出貨', '已出貨', '平台已結算', '已入帳', '已對帳']

  const [rows1, setRows1] = useState(null)
  const [rows2, setRows2] = useState(null)
  const [fileName1, setFileName1] = useState('')
  const [fileName2, setFileName2] = useState('')
  const [reconMsg, setReconMsg] = useState('')
  const [reconResult, setReconResult] = useState(null)
  const fileRef1 = useRef(null)
  const fileRef2 = useRef(null)

  const [orders, setOrders] = useState([])
  const [filterStatus, setFilterStatus] = useState('')
  const [filterMonth, setFilterMonth] = useState('')
  const [onlyDiff, setOnlyDiff] = useState(false)
  const [sortCol, setSortCol] = useState('order_date')
  const [sortDir, setSortDir] = useState('desc')
  const [selectedIds, setSelectedIds] = useState(new Set())
  const [deleteMsg, setDeleteMsg] = useState('')
  const [editOrder, setEditOrder] = useState(null)
  const [editMsg, setEditMsg] = useState('')
  const [viewInvKey, setViewInvKey] = useState(null)

  const [invMethod, setInvMethod] = useState('auto')
  const [invNo, setInvNo] = useState('')
  const [invDate, setInvDate] = useState('')
  const [invAmount, setInvAmount] = useState('')
  const [invFrom, setInvFrom] = useState('')
  const [invTo, setInvTo] = useState('')
  const [invPreview, setInvPreview] = useState(null)
  const [invMsg, setInvMsg] = useState('')
  const [checkedIds, setCheckedIds] = useState(new Set())

  const [inv2No, setInv2No] = useState('')
  const [inv2Date, setInv2Date] = useState('')
  const [inv2Amount, setInv2Amount] = useState('')
  const [inv2Msg, setInv2Msg] = useState('')

  useEffect(() => { loadOrders() }, [])

  async function loadOrders() {
    const data = await loadGatewayOrders(supabase, gateway)
    setOrders(data)
  }

  function readFile(e, setRows, setFileName) {
    const f = e.target.files?.[0]
    if (!f) return
    setFileName(f.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      setRows(XLSX.utils.sheet_to_json(ws, { defval: '' }))
    }
    reader.readAsArrayBuffer(f)
  }

  async function handleReconcile() {
    setReconMsg('比對中…')
    try {
      let parsed
      if (isTwoFile) {
        if (!rows1 || !rows2) { setReconMsg('請分別上傳 D-1 和 D-2 兩份對帳單'); return }
        parsed = parseOfficialLinePayReconDual(rows1, rows2)
      } else {
        if (!rows1) { setReconMsg('請先上傳對帳單'); return }
        parsed = RECON_PARSERS[gateway](rows1)
      }
      const result = await reconcile(supabase, gateway, parsed)
      setReconResult(result)
      setReconMsg(`比對 ${parsed.length} 筆、回填 ${result.updated} 筆、未對應 ${result.unmatched.length} 筆`)
      loadOrders()
    } catch(e) { setReconMsg('錯誤：' + e.message) }
  }

  function calcDiff(o) {
    if (o.actual_in != null && o.payable != null)
      return Math.round((o.actual_in - o.payable) * 100) / 100
    return null
  }

  async function runInvPreviewAuto() {
    if (!invFrom || !invTo) { setInvMsg('請填寫涵蓋期間'); return }
    setInvMsg('查詢中…'); setInvPreview(null)
    try {
      const result = await previewInvoice(supabase, { gateway, dateFrom: invFrom, dateTo: invTo })
      setInvPreview({ ...result, method: 'auto' })
      setInvMsg(result.orders.length ? '' : '查無符合期間的訂單')
    } catch(e) { setInvMsg('錯誤：' + e.message) }
  }

  const manualOrders = orders.filter(o => !o.fee_invoice_no)
  const checkedOrders = manualOrders.filter(o => checkedIds.has(o.id))
  const manualFeeSum = Math.round(checkedOrders.reduce((s, o) => s + (o.fee_total || 0), 0) * 100) / 100

  function toggleCheck(id) {
    setCheckedIds(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n })
  }
  function toggleAll() {
    setCheckedIds(prev => prev.size === manualOrders.length ? new Set() : new Set(manualOrders.map(o => o.id)))
  }
  function switchMethod(v) { setInvMethod(v); setInvPreview(null); setCheckedIds(new Set()); setInvMsg('') }

  async function runApplyInvoice() {
    const amount = parseFloat(invAmount) || 0
    const isAutoMode = invPreview?.method === 'auto'
    const orderIds = isAutoMode ? invPreview.orders.map(o => o.id) : [...checkedIds]
    const feeSum = isAutoMode ? invPreview.feeSum : manualFeeSum
    if (!orderIds.length) { setInvMsg('請先選取訂單'); return }
    const isMatch = amount > 0 && Math.abs(amount - feeSum) < 0.01
    try {
      await applyInvoice(supabase, { orderIds, invoiceNo: invNo, invoiceDate: invDate, invoiceAmount: amount || null, isMatch })
      setInvMsg(`已套用至 ${orderIds.length} 筆（${isMatch ? '相符' : '有差異'}）`)
      setInvPreview(null); setCheckedIds(new Set()); loadOrders()
    } catch(e) { setInvMsg('錯誤：' + e.message) }
  }

  async function applyPayuniAccountFee() {
    if (!inv2No) { setInv2Msg('請填寫發票號碼'); return }
    const note = `PayUni服務費 ${inv2No} ${inv2Date} $${inv2Amount}`
    const ids = orders.map(o => o.id)
    if (!ids.length) { setInv2Msg('無訂單可套用'); return }
    const { error } = await supabase.from('shipping_orders').update({ account_fee_note: note }).in('id', ids)
    setInv2Msg(error ? '錯誤：' + error.message : `已記錄至 ${ids.length} 筆`)
    loadOrders()
  }

  async function saveEditOrder(updates) {
    setEditMsg('儲存中…')
    const { error } = await supabase
      .from('shipping_orders')
      .update({
        sa_no: updates.sa_no || null,
        recon_status: updates.recon_status || null,
        note: updates.note || null,
        fee_invoice_no: updates.fee_invoice_no || null,
      })
      .eq('id', updates.id)
    if (error) { setEditMsg('錯誤：' + error.message); return }
    setEditMsg('')
    setEditOrder(null)
    loadOrders()
  }

  function toggleSelect(id) {
    setSelectedIds(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n })
  }
  function toggleSelectAll() {
    setSelectedIds(prev => prev.size === shownOrders.length ? new Set() : new Set(shownOrders.map(o => o.id)))
  }
  async function deleteSelected() {
    if (!selectedIds.size) return
    if (!window.confirm(`確定要刪除選取的 ${selectedIds.size} 筆訂單？此操作無法復原。`)) return
    setDeleteMsg('刪除中…')
    const { error } = await supabase.from('shipping_orders').delete().in('id', [...selectedIds])
    if (error) { setDeleteMsg('錯誤：' + error.message); return }
    setDeleteMsg(`已刪除 ${selectedIds.size} 筆`)
    setSelectedIds(new Set())
    loadOrders()
  }

  function exportOrders() {
    const data = shownOrders.map(o => ({
      銷貨單號: o.sa_no ?? '', 平台訂單編號: o.ref_no, 訂單日期: o.order_date ?? '', 應收: o.total, 手續費: o.fee_total ?? '',
      應入帳: o.payable ?? '', 實際入帳: o.actual_in ?? '', 入帳日: o.in_date ?? '',
      差異: calcDiff(o) ?? '', 狀態: o.recon_status, 發票號碼: o.fee_invoice_no ?? '',
    }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), gwInfo.label || '對帳')
    XLSX.writeFile(wb, `對帳_${gwInfo.label}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  const months = [...new Set(orders.map(o => (o.order_date || '').slice(0, 7)).filter(Boolean))].sort().reverse()

  const SORT_KEY = {
    '銷貨單號': 'sa_no', '平台訂單編號': 'ref_no', '訂單日期': 'order_date',
    '應收': 'total', '手續費': 'fee_total', '應入帳': 'payable',
    '實際入帳': 'actual_in', '入帳日': 'in_date', '差異': '_diff',
    '狀態': 'recon_status', '發票號碼': 'fee_invoice_no',
  }

  function handleSort(col) {
    if (sortCol === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc')
    else { setSortCol(col); setSortDir('asc') }
  }

  function getVal(o, key) {
    if (key === '_diff') return calcDiff(o) ?? -Infinity
    const v = o[key]
    return v == null ? '' : v
  }

  const shownOrders = orders
    .filter(o => {
      if (filterStatus && o.recon_status !== filterStatus) return false
      if (filterMonth && (o.order_date || '').slice(0, 7) !== filterMonth) return false
      if (onlyDiff) { const d = calcDiff(o); if (d == null || d === 0) return false }
      return true
    })
    .sort((a, b) => {
      const va = getVal(a, sortCol), vb = getVal(b, sortCol)
      const cmp = typeof va === 'number' && typeof vb === 'number'
        ? va - vb
        : String(va).localeCompare(String(vb), 'zh-Hant')
      return sortDir === 'asc' ? cmp : -cmp
    })

  // 發票分組：計算每張發票在目前清單中的合計手續費與筆數
  const invoiceGroups = {}
  shownOrders.forEach(o => {
    if (!o.fee_invoice_no) return
    if (!invoiceGroups[o.fee_invoice_no]) invoiceGroups[o.fee_invoice_no] = {
      invAmount: o.fee_invoice_amount ?? null,
      invDate: o.fee_invoice_date ?? null,
      invoiceCheck: o.invoice_check ?? null,
      feeSum: 0, count: 0,
    }
    invoiceGroups[o.fee_invoice_no].feeSum += o.fee_total || 0
    invoiceGroups[o.fee_invoice_no].count++
  })
  const INV_BG = ['#e6f4f0', '#e8edf8']   // 交替淡綠 / 淡藍
  const invColorIdx = {}
  let _ci = 0
  Object.keys(invoiceGroups).forEach(k => { invColorIdx[k] = _ci++ % 2 })

  const invFeeSum = invPreview?.feeSum ?? (invMethod === 'manual' ? manualFeeSum : null)
  const invAmountNum = parseFloat(invAmount) || 0
  const invDiff = invAmountNum > 0 && invFeeSum != null ? Math.round((invAmountNum - invFeeSum) * 100) / 100 : null
  const invIsMatch = invDiff != null && Math.abs(invDiff) < 0.01
  const hasInvOrders = invPreview?.orders?.length > 0 || (invMethod === 'manual' && checkedIds.size > 0)

  return (
    <div>
      {/* 上傳撥款明細 */}
      <Card>
        <strong style={{ fontSize: 14 }}>上傳撥款明細</strong>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginTop: 10, alignItems: 'center' }}>
          {isTwoFile ? (
            <>
              <div>
                <div style={{ fontSize: 12, color: C.sub, marginBottom: 4 }}>D-1 新 LINE Pay 對帳單</div>
                <input ref={fileRef1} type="file" accept=".xlsx,.xls" onChange={e => readFile(e, setRows1, setFileName1)} style={{ display: 'none' }} />
                <button onClick={() => fileRef1.current.click()} style={btnGhost}>{fileName1 || '選擇 D-1 檔案'}</button>
                {rows1 && <span style={{ fontSize: 12, color: C.brand, marginLeft: 6 }}>✓ {rows1.length} 列</span>}
              </div>
              <div>
                <div style={{ fontSize: 12, color: C.sub, marginBottom: 4 }}>D-2 PayUni 電子錢包對帳單</div>
                <input ref={fileRef2} type="file" accept=".xlsx,.xls" onChange={e => readFile(e, setRows2, setFileName2)} style={{ display: 'none' }} />
                <button onClick={() => fileRef2.current.click()} style={btnGhost}>{fileName2 || '選擇 D-2 檔案'}</button>
                {rows2 && <span style={{ fontSize: 12, color: C.brand, marginLeft: 6 }}>✓ {rows2.length} 列</span>}
              </div>
            </>
          ) : (
            <>
              <input ref={fileRef1} type="file" accept=".xlsx,.xls" onChange={e => readFile(e, setRows1, setFileName1)} style={{ display: 'none' }} />
              <button onClick={() => fileRef1.current.click()} style={btnGhost}>{fileName1 || '選擇對帳單'}</button>
              {rows1 && <span style={{ fontSize: 12, color: C.brand }}>✓ {rows1.length} 列</span>}
            </>
          )}
          <button onClick={handleReconcile} style={btnPrimary}>比對回填</button>
        </div>
        {reconMsg && (
          <p style={{ marginTop: 8, marginBottom: 0, fontSize: 13,
            color: reconMsg.includes('錯誤') || reconMsg.includes('請') ? C.danger : C.brand }}>
            {reconMsg}
          </p>
        )}
        {reconResult && reconResult.updated > 0 && (
          <div style={{ marginTop: 10, padding: '10px 14px', background: C.brandBg, borderRadius: 8,
            display: 'flex', gap: 24, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 13 }}>
              手續費合計：<strong>{reconResult.feeTotal.toLocaleString()}</strong>
              <span style={{ fontSize: 11, color: C.sub, marginLeft: 4 }}>（供發票核對）</span>
            </span>
            <span style={{ fontSize: 13 }}>
              預計撥款金額：<strong>{reconResult.payableTotal.toLocaleString()}</strong>
              <span style={{ fontSize: 11, color: C.sub, marginLeft: 4 }}>（供玉山對帳單核對）</span>
            </span>
          </div>
        )}
        {reconResult?.unmatched?.length > 0 && (
          <p style={{ marginTop: 6, marginBottom: 0, fontSize: 12, color: C.warn }}>
            未對應：{reconResult.unmatched.slice(0, 5).join('、')}
            {reconResult.unmatched.length > 5 && `…等 ${reconResult.unmatched.length} 筆`}
          </p>
        )}
      </Card>

      {/* 對帳狀態清單 */}
      <Card>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', marginBottom: 10 }}>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <select value={filterMonth} onChange={e => setFilterMonth(e.target.value)} style={{ ...inp, width: 'auto' }}>
              <option value="">全部月份</option>
              {months.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
            <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={{ ...inp, width: 'auto' }}>
              <option value="">全部狀態</option>
              {STATUSES.map(s => <option key={s}>{s}</option>)}
            </select>
            <label style={{ fontSize: 13, display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
              <input type="checkbox" checked={onlyDiff} onChange={e => setOnlyDiff(e.target.checked)} />
              只看差異
            </label>
            <span style={{ fontSize: 13, color: C.sub }}>{shownOrders.length} 筆</span>
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {selectedIds.size > 0 && (
              <button onClick={deleteSelected} style={{ ...btnGhost, color: C.danger, borderColor: C.danger }}>
                刪除 {selectedIds.size} 筆
              </button>
            )}
            {deleteMsg && <span style={{ fontSize: 12, color: deleteMsg.includes('錯誤') ? C.danger : C.sub }}>{deleteMsg}</span>}
            <button onClick={loadOrders} style={btnGhost}>重新整理</button>
            <button onClick={exportOrders} style={btnPrimary}>匯出</button>
          </div>
        </div>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 13 }}>
            <thead>
              <tr>
                <th style={th}></th>
                <th style={th}>
                  <input type="checkbox"
                    checked={shownOrders.length > 0 && selectedIds.size === shownOrders.length}
                    onChange={toggleSelectAll} />
                </th>
                {['銷貨單號', '平台訂單編號', '訂單日期', '應收', '手續費', '應入帳', '實際入帳', '入帳日', '差異', '狀態', '發票號碼'].map(c => {
                  const key = SORT_KEY[c]
                  const active = sortCol === key
                  return (
                    <th key={c} style={{ ...th, cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' }}
                      onClick={() => key && handleSort(key)}>
                      {c}{active ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {(() => {
                const seenInv = new Set()
                return shownOrders.map((o, i) => {
                  const d = calcDiff(o); const hasDiff = d != null && d !== 0
                  const invBg = o.fee_invoice_no ? INV_BG[invColorIdx[o.fee_invoice_no]] : undefined
                  const rowBg = selectedIds.has(o.id) ? C.brandBg : invBg
                  const isFirstInv = o.fee_invoice_no && !seenInv.has(o.fee_invoice_no)
                  if (o.fee_invoice_no) seenInv.add(o.fee_invoice_no)
                  const grp = o.fee_invoice_no ? invoiceGroups[o.fee_invoice_no] : null
                  return (
                    <tr key={i} style={{ background: rowBg }}>
                      <td style={td}>
                        <input type="checkbox" checked={selectedIds.has(o.id)} onChange={() => toggleSelect(o.id)} />
                      </td>
                      <td style={td}>
                        <button onClick={() => { setEditOrder({ ...o }); setEditMsg('') }}
                          style={{ fontSize: 12, padding: '2px 8px', border: `1px solid ${C.line}`,
                            borderRadius: 6, background: '#fff', cursor: 'pointer', color: C.sub }}>
                          編輯
                        </button>
                      </td>
                      <td style={{ ...td, fontFamily: 'monospace', fontSize: 12 }}>{o.sa_no || '—'}</td>
                      <td style={{ ...td, fontFamily: 'monospace', fontSize: 12 }}>{o.ref_no}</td>
                      <td style={td}>{o.order_date ? o.order_date.slice(0, 10) : '—'}</td>
                      <td style={{ ...td, textAlign: 'right' }}>{o.total?.toLocaleString()}</td>
                      <td style={{ ...td, textAlign: 'right' }}>{o.fee_total != null ? o.fee_total.toLocaleString() : '—'}</td>
                      <td style={{ ...td, textAlign: 'right' }}>{o.payable != null ? o.payable.toLocaleString() : '—'}</td>
                      <td style={{ ...td, textAlign: 'right' }}>{o.actual_in != null ? o.actual_in.toLocaleString() : '—'}</td>
                      <td style={td}>{o.in_date || '—'}</td>
                      <td style={{ ...td, textAlign: 'right', color: hasDiff ? C.danger : C.ink, fontWeight: hasDiff ? 600 : 400 }}>
                        {d != null ? d.toLocaleString() : '—'}
                      </td>
                      <td style={td}>
                        <span style={{ padding: '2px 8px', borderRadius: 99, fontSize: 12,
                          background: statusBg(o.recon_status), color: statusColor(o.recon_status) }}>
                          {o.recon_status || '—'}
                        </span>
                      </td>
                      <td
                        style={{ ...td, fontFamily: 'monospace', fontSize: 12, cursor: o.fee_invoice_no ? 'pointer' : 'default' }}
                        onClick={o.fee_invoice_no ? () => setViewInvKey(o.fee_invoice_no) : undefined}
                      >
                        {isFirstInv && grp ? (
                          <div>
                            <div style={{ color: C.brand, textDecoration: 'underline' }}>{o.fee_invoice_no}</div>
                            <div style={{ fontSize: 11, color: C.sub, marginTop: 2, whiteSpace: 'nowrap' }}>
                              {grp.count} 筆・發票金額 {grp.invAmount != null ? grp.invAmount.toLocaleString() : '—'}
                            </div>
                          </div>
                        ) : o.fee_invoice_no ? <span style={{ fontSize: 11, color: C.sub }}>↑</span> : '—'}
                      </td>
                    </tr>
                  )
                })
              })()}
              {shownOrders.length === 0 && (
                <tr><td colSpan={13} style={{ ...td, textAlign: 'center', color: C.sub, padding: 24 }}>沒有資料</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {/* 發票核對 */}
      <Card>
        <strong style={{ fontSize: 14 }}>發票核對（手續費進項發票）</strong>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 12, alignItems: 'flex-end' }}>
          <Field label="發票號碼">
            <input value={invNo} onChange={e => setInvNo(e.target.value)} placeholder="AB-12345678" style={inp} />
          </Field>
          <Field label="發票日期">
            <input type="date" value={invDate} onChange={e => setInvDate(e.target.value)} style={inp} />
          </Field>
          <Field label="發票金額">
            <input type="number" value={invAmount} onChange={e => setInvAmount(e.target.value)} placeholder="0" style={inp} />
          </Field>
        </div>

        <div style={{ display: 'flex', margin: '12px 0', gap: 0 }}>
          {[['auto', '方式 A — 期間篩選'], ['manual', '方式 B — 手動勾選']].map(([v, lbl], i) => (
            <button key={v} onClick={() => switchMethod(v)} style={{
              padding: '6px 14px', border: `1px solid ${C.line}`, cursor: 'pointer', fontSize: 13,
              background: invMethod === v ? C.brand : '#fff', color: invMethod === v ? '#fff' : C.sub,
              borderRadius: i === 0 ? '8px 0 0 8px' : '0 8px 8px 0',
            }}>{lbl}</button>
          ))}
        </div>

        {invMethod === 'auto' && (
          <div style={{ display: 'flex', gap: 10, alignItems: 'flex-end', flexWrap: 'wrap' }}>
            <Field label="期間起"><input type="date" value={invFrom} onChange={e => setInvFrom(e.target.value)} style={inp} /></Field>
            <Field label="期間訖"><input type="date" value={invTo} onChange={e => setInvTo(e.target.value)} style={inp} /></Field>
            <div style={{ paddingBottom: 2 }}><button onClick={runInvPreviewAuto} style={btnPrimary}>查詢</button></div>
          </div>
        )}

        {invMethod === 'manual' && (
          <div style={{ overflowX: 'auto', maxHeight: 240, overflowY: 'auto', marginBottom: 8 }}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
              <thead>
                <tr>
                  <th style={th}>
                    <input type="checkbox"
                      checked={manualOrders.length > 0 && checkedIds.size === manualOrders.length}
                      onChange={toggleAll} />
                  </th>
                  {['銷貨單號', '入帳日', '手續費'].map(c => <th key={c} style={th}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {manualOrders.length === 0 && (
                  <tr><td colSpan={4} style={{ ...td, textAlign: 'center', color: C.sub }}>所有訂單都已歸發票</td></tr>
                )}
                {manualOrders.map((o, i) => (
                  <tr key={i} style={{ background: checkedIds.has(o.id) ? C.brandBg : '#fff' }}>
                    <td style={td}><input type="checkbox" checked={checkedIds.has(o.id)} onChange={() => toggleCheck(o.id)} /></td>
                    <td style={{ ...td, fontFamily: 'monospace' }}>{o.ref_no}</td>
                    <td style={td}>{o.in_date || o.order_date || '—'}</td>
                    <td style={{ ...td, textAlign: 'right' }}>{(o.fee_total ?? 0).toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {hasInvOrders && (
          <div style={{ padding: '10px 14px', borderRadius: 8, marginTop: 10,
            background: invIsMatch ? C.brandBg : invDiff != null ? C.warnBg : '#f5f5f5',
            display: 'flex', gap: 20, alignItems: 'center', flexWrap: 'wrap' }}>
            {invFeeSum != null && <span style={{ fontSize: 13 }}>手續費加總：<strong>{invFeeSum.toLocaleString()}</strong></span>}
            {invAmountNum > 0 && <span style={{ fontSize: 13 }}>發票金額：<strong>{invAmountNum.toLocaleString()}</strong></span>}
            {invDiff != null && (
              <span style={{ fontSize: 13, color: invIsMatch ? C.brand : C.danger, fontWeight: 600 }}>
                差異：{invDiff.toLocaleString()}　{invIsMatch ? '✓ 相符' : '✗ 有差異'}
              </span>
            )}
            <button onClick={runApplyInvoice} style={btnPrimary}>套用</button>
          </div>
        )}
        {invMsg && (
          <p style={{ marginTop: 8, marginBottom: 0, fontSize: 13,
            color: invMsg.includes('錯誤') ? C.danger : invMsg.includes('相符') ? C.brand : C.sub }}>
            {invMsg}
          </p>
        )}
      </Card>

      {/* 編輯 modal */}
      {editOrder && (
        <div style={overlay} onClick={() => setEditOrder(null)}>
          <div style={modal} onClick={e => e.stopPropagation()}>
            <h3 style={{ marginTop: 0, fontSize: 15 }}>編輯訂單</h3>
            <p style={{ fontSize: 12, color: C.sub, margin: '0 0 12px' }}>平台訂單編號：{editOrder.ref_no}</p>
            <Field label="銷貨單號（ERP SA 單號）">
              <input value={editOrder.sa_no || ''} onChange={e => setEditOrder(p => ({ ...p, sa_no: e.target.value }))}
                placeholder="SA-XXXXXXXX" style={inp} />
            </Field>
            <Field label="狀態">
              <select value={editOrder.recon_status || ''} onChange={e => setEditOrder(p => ({ ...p, recon_status: e.target.value }))} style={inp}>
                {['待出貨', '已出貨', '平台已結算', '已入帳', '已對帳'].map(s => <option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="發票號碼">
              <input value={editOrder.fee_invoice_no || ''} onChange={e => setEditOrder(p => ({ ...p, fee_invoice_no: e.target.value }))}
                placeholder="AB-12345678" style={inp} />
            </Field>
            <Field label="備註">
              <input value={editOrder.note || ''} onChange={e => setEditOrder(p => ({ ...p, note: e.target.value }))} style={inp} />
            </Field>
            {editMsg && <p style={{ fontSize: 13, color: editMsg.includes('錯誤') ? C.danger : C.sub, margin: '4px 0' }}>{editMsg}</p>}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
              <button onClick={() => setEditOrder(null)} style={btnGhost}>取消</button>
              <button onClick={() => saveEditOrder(editOrder)} style={btnPrimary}>儲存</button>
            </div>
          </div>
        </div>
      )}

      {/* 發票資訊 modal */}
      {viewInvKey && (() => {
        const grp = invoiceGroups[viewInvKey]
        if (!grp) return null
        const checkColor = grp.invoiceCheck === '相符' ? C.brand : grp.invoiceCheck === '有差異' ? C.danger : C.sub
        const rows = [
          ['發票號碼', viewInvKey],
          ['發票日期', grp.invDate || '—'],
          ['發票金額', grp.invAmount != null ? `NT$ ${Number(grp.invAmount).toLocaleString()}` : '—'],
          ['手續費合計', `NT$ ${Math.round(grp.feeSum * 100) / 100}`],
          ['包含訂單', `${grp.count} 筆`],
          ['核對結果', grp.invoiceCheck || '—'],
        ]
        return (
          <div style={overlay} onClick={() => setViewInvKey(null)}>
            <div style={{ ...modal, width: 380 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 15 }}>發票資訊</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <tbody>
                  {rows.map(([label, val]) => (
                    <tr key={label} style={{ borderBottom: '1px solid #f0f0f0' }}>
                      <td style={{ padding: '8px 0', color: C.sub, width: 100 }}>{label}</td>
                      <td style={{ padding: '8px 0', fontWeight: label === '核對結果' ? 600 : 400, color: label === '核對結果' ? checkColor : '#222', fontFamily: label === '發票號碼' ? 'monospace' : 'inherit' }}>{val}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 20 }}>
                <button onClick={() => setViewInvKey(null)} style={btnGhost}>關閉</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* 官網 LINE Pay 第二層：PayUni 帳戶層服務費 */}
      {isLinePayOfficial && (
        <Card>
          <strong style={{ fontSize: 14 }}>PayUni 服務費發票（帳戶層，0.2%）</strong>
          <p style={{ fontSize: 12, color: C.sub, margin: '4px 0 4px' }}>
            此費用為月結帳戶層，不逐筆計入 fee_total，請登記於備註欄供查閱。
          </p>
          <p style={{ fontSize: 12, color: C.sub, margin: '0 0 12px', fontStyle: 'italic' }}>
            提示：和和研信箱中搜尋「CYICI.60481225.EC 電子發票開立通知」可找到對應發票。
          </p>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
            <Field label="發票號碼"><input value={inv2No} onChange={e => setInv2No(e.target.value)} placeholder="XC19745594" style={inp} /></Field>
            <Field label="發票日期"><input type="date" value={inv2Date} onChange={e => setInv2Date(e.target.value)} style={inp} /></Field>
            <Field label="金額"><input type="number" value={inv2Amount} onChange={e => setInv2Amount(e.target.value)} placeholder="0" style={inp} /></Field>
            <div style={{ paddingBottom: 2 }}><button onClick={applyPayuniAccountFee} style={btnGhost}>記錄備註</button></div>
          </div>
          {inv2Msg && <p style={{ marginTop: 8, marginBottom: 0, fontSize: 13, color: inv2Msg.includes('錯誤') ? C.danger : C.brand }}>{inv2Msg}</p>}
        </Card>
      )}
    </div>
  )
}

function statusBg(s) {
  if (s === '已對帳') return C.brandBg
  if (s === '已入帳') return '#e6f4f1'
  if (s === '平台已結算') return '#edf2fb'
  if (s === '已出貨') return '#f5f5f5'
  return C.warnBg
}
function statusColor(s) {
  if (s === '已對帳') return C.brand
  if (s === '已入帳') return '#1d7a6f'
  if (s === '平台已結算') return '#2c5282'
  if (s === '已出貨') return C.sub
  return C.warn
}

function EditModal({ row, onClose, onSave }) {
  const [r, setR] = useState(row)
  const set = (k, v) => setR((p) => ({ ...p, [k]: v }))
  return (
    <div style={overlay} onClick={onClose}>
      <div style={modal} onClick={(e) => e.stopPropagation()}>
        <h3 style={{ marginTop: 0 }}>{row.id ? '編輯規則' : '新增規則'}</h3>
        <Field label="平台">
          <select value={r.platform} onChange={(e) => set('platform', e.target.value)} style={inp}>
            {PLATFORMS.map((p) => <option key={p}>{p}</option>)}
          </select>
        </Field>
        <Field label="商品識別字串（關鍵字，會用包含比對）">
          <input value={r.match_text} onChange={(e) => set('match_text', e.target.value)} style={inp} />
        </Field>
        <Field label="綁定鍵（同一商品的主商品+贈品共用，如 sp_3box）">
          <input value={r.group_key} onChange={(e) => set('group_key', e.target.value)} style={inp} />
        </Field>
        <Field label="角色">
          <select value={r.role} onChange={(e) => set('role', e.target.value)} style={inp}>
            <option value="main">主商品</option>
            <option value="gift">贈品</option>
          </select>
        </Field>
        <Field label="編碼（如 001100POWA001 或 FREEGIFT00002）">
          <input value={r.code} onChange={(e) => set('code', e.target.value)} style={inp} />
        </Field>
        <Field label="品項名稱">
          <input value={r.item_name} onChange={(e) => set('item_name', e.target.value)} style={inp} />
        </Field>
        <div style={{ display: 'flex', gap: 12 }}>
          <Field label="數量"><input type="number" value={r.qty} onChange={(e) => set('qty', e.target.value)} style={inp} /></Field>
          <Field label="排序（主商品0、贈品1,2…）"><input type="number" value={r.sort_order} onChange={(e) => set('sort_order', e.target.value)} style={inp} /></Field>
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
          <button onClick={onClose} style={btnGhost}>取消</button>
          <button onClick={() => onSave(r)} style={btnPrimary}>儲存</button>
        </div>
      </div>
    </div>
  )
}

// ====== 小元件 / 樣式 ======
function Card({ children, style }) {
  return <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 16, marginBottom: 16, ...style }}>{children}</div>
}
function Field({ label, children }) {
  return <div style={{ marginBottom: 10, flex: 1 }}>
    <label style={{ display: 'block', fontSize: 12, color: C.sub, marginBottom: 4 }}>{label}</label>
    {children}
  </div>
}
const inp = { padding: '7px 10px', borderRadius: 8, border: `1px solid ${C.line}`, fontSize: 14, width: '100%', boxSizing: 'border-box' }
const btnPrimary = { padding: '8px 16px', borderRadius: 8, border: 'none', background: C.brand, color: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer' }
const btnGhost = { padding: '8px 16px', borderRadius: 8, border: `1px solid ${C.line}`, background: '#fff', color: C.ink, fontSize: 14, cursor: 'pointer' }
const miniBtn = { padding: '4px 8px', borderRadius: 6, border: 'none', background: 'transparent', color: C.brand, fontSize: 13, cursor: 'pointer' }
const th = { textAlign: 'left', padding: '8px 10px', borderBottom: `2px solid ${C.line}`, color: C.sub, fontWeight: 600, whiteSpace: 'nowrap' }
const td = { padding: '7px 10px', borderBottom: `1px solid ${C.line}`, whiteSpace: 'nowrap' }
const overlay = { position: 'fixed', inset: 0, background: 'rgba(0,0,0,.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, zIndex: 50 }
const modal = { background: '#fff', borderRadius: 14, padding: 24, width: 480, maxWidth: '100%', maxHeight: '90vh', overflowY: 'auto' }
