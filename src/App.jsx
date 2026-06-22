import React, { useState, useEffect, useRef } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import { PARSERS, detectPlatform } from './parsers.js'
import { buildBlocks } from './transform.js'

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
          </nav>
        </div>
      </header>
      <main style={{ maxWidth: 1100, margin: '0 auto', padding: 20 }}>
        {tab === 'convert' ? <ConvertPage /> : <MappingPage />}
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
        tracking_no: String(o.tracking_no || ''), total: o.total || 0, shipping_fee: o.shipping_fee || 0,
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
