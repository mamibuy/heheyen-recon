import React, { useState, useEffect, useRef } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import { PARSERS, detectPlatform, excelDate, parseOfficialBulk } from './parsers.js'
import { buildBlocks } from './transform.js'
import { RECON_PARSERS, parseOfficialLinePayReconDual, parseMegabankRecon, checkReconColumns, checkDualReconColumns, checkMegabankColumns, megabankRateWarnings } from './recon_parsers.js'
import { reconcile, previewInvoice, applyInvoice, loadGatewayOrders } from './reconcile.js'

// ====== Supabase（沿用 Mamibuy 專案）======
const supabase = createClient(
  'https://geirbvjkwsewglvvrfmg.supabase.co',
  'sb_publishable_yDgLU7V2PcL_2QmrQkxo2w_WZGEbP63'
)

// ====== 品牌 Design Token（取樣自「純LGG活菌」主視覺）======
// 寶藍為主色、深藍作標題／大數字、綠與紅僅作對帳語意（相符／差異）
const T = {
  bg: '#eef2f9', surface: '#ffffff', text: '#16233f',
  divider: 'rgba(22,35,63,.14)',
  n100: '#f6f8fc', n200: '#eaeff8', n300: '#d9e1ef', n400: '#bcc7dd',
  n500: '#94a1bd', n600: '#6a7793', n700: '#48546f', n800: '#2b3550', n900: '#16233f',
  a: '#123f96',
  a100: '#e8f0fb', a200: '#cadef6', a300: '#9bbfec', a400: '#5688dc',
  a500: '#2158bf', a600: '#0f3684', a700: '#0c2c6a', a800: '#0a2250', a900: '#071633',
  g: '#1f9d6b',
  g100: '#e4f6ee', g200: '#c4ead9', g300: '#97dabb', g400: '#57c295',
  g500: '#1f9d6b', g600: '#16855a', g700: '#0f6b48',
  navy: '#16265a', gold: '#c8a858',
  danger: '#c0392b', dangerBg: '#fdecea',
  shadowSm: '0 1px 2px rgba(22,35,63,.14)',
  shadowMd: '0 3px 10px rgba(22,35,63,.16)',
  shadowLg: '0 12px 32px rgba(22,35,63,.22)',
  rCard: 26, rPanel: 18, rInner: 14, rPill: 999,
}

// 既有頁面（出貨轉換／商品對照表）沿用 C，改指向新品牌色票
const C = {
  bg: T.bg, card: T.surface, line: T.divider, ink: T.text,
  sub: T.n600, brand: T.a, brandBg: T.a100,
  warn: '#b4541a', warnBg: '#fbeee2', danger: T.danger,
}

const PLATFORMS = ['蝦皮', 'LINE商城', '酷澎', '官網', '兆豐']

const GATEWAY_LIST = [
  { key: 'coupang',        label: '酷澎',               dot: T.n500 },
  { key: 'shopee',         label: '蝦皮',               dot: T.a500 },
  { key: 'payuni_cc',      label: '官網 › 信用卡',      dot: T.g500 },
  { key: 'payuni_linepay', label: '官網 › LINE Pay',    dot: T.g600, twoFile: true },
  { key: 'linepay',        label: 'LINE商城 › LINE Pay', dot: T.a400 },
  { key: 'lanxin',         label: 'LINE商城 › 信用卡',   dot: T.n400 },
  { key: 'megabank',       label: '兆豐福利網',          dot: T.gold, twoFile: true },
]

// 天心 SA 銷貨明細表：檔案最上面有公司名／日期區間等標題列，真正的欄位表頭不在第一列。
// 動態尋找含「客戶訂單」「單號」的表頭列，再逐列組成物件；找不到（表頭本來就在第一列的檔）
// 則退回一般物件模式，行為與原本一致。
function parseTianxinSheet(ws) {
  const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  // 表頭欄位有時帶前後空白，先 trim 再找，避免誤判成「沒有表頭」而退回錯誤的第一列
  const trimmed = arr.map(r => Array.isArray(r) ? r.map(c => String(c).trim()) : [])
  const hi = trimmed.findIndex(r => r.includes('客戶訂單') && r.includes('單號'))
  if (hi < 0) return XLSX.utils.sheet_to_json(ws, { defval: '' })
  const header = trimmed[hi]
  const out = []
  for (let i = hi + 1; i < arr.length; i++) {
    const r = arr[i]
    if (!r.some(c => String(c).trim() !== '')) continue   // 跳過空白列
    const obj = {}
    header.forEach((h, ci) => { if (h) obj[h] = r[ci] })
    out.push(obj)
  }
  return out
}

// 對帳單常見的比對鑰匙欄位，用來判斷某一列是不是真正的表頭
const KEY_COLS = ['訂單編號', '商店訂單編號', '訂單號碼', '交易號碼', '客戶訂單']

// 有些平台報表最上面有賣家帳號／期間／小計等區塊，真正的表頭不在第一列。
// 例：蝦皮「我的進帳」表頭在第 6 列，直接 sheet_to_json 會把「賣家帳號/從/至」
// 當成欄位名，解析結果全是 __EMPTY_N，parser 一筆都讀不到。
// 先用預設方式解析，認得出鑰匙欄位就直接用（既有可正常運作的報表行為不變）；
// 認不出才往下尋找含鑰匙欄位的那一列當表頭。
function sheetToRows(ws) {
  const direct = XLSX.utils.sheet_to_json(ws, { defval: '' })
  if (direct.length && Object.keys(direct[0]).some(k => KEY_COLS.includes(String(k).trim()))) return direct

  const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  const trimmed = arr.map(r => Array.isArray(r) ? r.map(c => String(c).trim()) : [])
  const hi = trimmed.findIndex(r => r.some(c => KEY_COLS.includes(c)))
  if (hi < 0) return direct

  const header = trimmed[hi]
  const out = []
  for (let i = hi + 1; i < arr.length; i++) {
    const r = arr[i]
    if (!r.some(c => String(c).trim() !== '')) continue   // 跳過空白列
    const obj = {}
    header.forEach((h, ci) => { if (h) obj[h] = r[ci] })
    out.push(obj)
  }
  return out
}

// 讀天心銷貨明細（.xls/.xlsx），套用上面的動態表頭解析
function readTianxinFile(e, setRows, setFileName) {
  const f = e.target.files?.[0]
  if (!f) return
  setFileName(f.name)
  const reader = new FileReader()
  reader.onload = ev => {
    const wb = XLSX.read(ev.target.result, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    setRows(parseTianxinSheet(ws))
  }
  reader.readAsArrayBuffer(f)
}

function PasswordGate({ children }) {
  const [authed, setAuthed] = useState(() => localStorage.getItem('hhy_auth') === '1')
  const [pw, setPw] = useState('')
  const [err, setErr] = useState(false)
  if (authed) return children
  return (
    <div style={{ minHeight: '100vh', background: C.bg, display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontFamily: '-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC","Microsoft JhengHei",sans-serif' }}>
      <div style={{ background: '#fff', borderRadius: 14, padding: '40px 36px', boxShadow: '0 4px 24px rgba(0,0,0,0.10)', minWidth: 300, textAlign: 'center' }}>
        <div style={{ fontSize: 22, fontWeight: 700, marginBottom: 6, color: C.ink }}>和和研 · 電商出貨彙整</div>
        <div style={{ fontSize: 13, color: C.sub, marginBottom: 24 }}>請輸入密碼以繼續</div>
        <input
          type="password"
          value={pw}
          onChange={e => { setPw(e.target.value); setErr(false) }}
          onKeyDown={e => {
            if (e.key === 'Enter') {
              if (pw === 'imheheyen') { localStorage.setItem('hhy_auth', '1'); setAuthed(true) }
              else setErr(true)
            }
          }}
          placeholder="密碼"
          autoFocus
          style={{ width: '100%', padding: '10px 14px', fontSize: 15, border: `1.5px solid ${err ? C.danger : C.line}`,
            borderRadius: 8, outline: 'none', boxSizing: 'border-box', marginBottom: 10 }}
        />
        {err && <div style={{ fontSize: 12, color: C.danger, marginBottom: 8 }}>密碼錯誤，請再試一次</div>}
        <button
          onClick={() => {
            if (pw === 'imheheyen') { localStorage.setItem('hhy_auth', '1'); setAuthed(true) }
            else setErr(true)
          }}
          style={{ width: '100%', padding: '10px 0', background: C.brand, color: '#fff', border: 'none',
            borderRadius: 8, fontSize: 15, fontWeight: 600, cursor: 'pointer' }}>
          進入
        </button>
      </div>
    </div>
  )
}

export default function App() {
  const [tab, setTab] = useState('convert')
  return (
    <PasswordGate>
    <div style={{ minHeight: '100vh', background: T.bg, color: T.text,
      fontFamily: '-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang TC","Microsoft JhengHei",sans-serif' }}>
      <style>{`
        .hhy-row:hover { filter: brightness(.97) }
        .hhy-row td { transition: background .12s }
        button:focus-visible { outline: 2px solid ${T.a}; outline-offset: 2px }
        input:focus-visible, select:focus-visible, textarea:focus-visible { outline: 2px solid ${T.a}; outline-offset: 1px }
      `}</style>
      <div style={{ display: 'flex', alignItems: 'center', gap: 26, padding: '14px 32px',
        background: T.n100, boxShadow: T.shadowSm, position: 'sticky', top: 0, zIndex: 20, flexWrap: 'wrap' }}>
        <span style={{ fontSize: 20, fontWeight: 700, color: T.a700, whiteSpace: 'nowrap' }}>和和研 · 對帳系統</span>
        <nav style={{ display: 'flex', gap: 4, marginLeft: 8 }}>
          <TabBtn active={tab === 'convert'} onClick={() => setTab('convert')}>出貨轉換</TabBtn>
          <TabBtn active={tab === 'mapping'} onClick={() => setTab('mapping')}>商品對照表</TabBtn>
          <TabBtn active={tab === 'recon'} onClick={() => setTab('recon')}>金流對帳</TabBtn>
        </nav>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 10, color: T.n600, fontSize: 13 }}>
          <span style={{ width: 30, height: 30, borderRadius: '50%', background: T.g300, display: 'grid',
            placeItems: 'center', color: T.g700, fontSize: 13, fontWeight: 700 }}>和</span>
          營運人員
        </div>
      </div>
      <main style={{ maxWidth: tab === 'recon' ? 1280 : 1100, margin: '0 auto',
        padding: tab === 'recon' ? '22px 32px 60px' : 20 }}>
        {tab === 'convert' && <ConvertPage />}
        {tab === 'mapping' && <MappingPage />}
        {tab === 'recon' && <ReconPage />}
      </main>
    </div>
    </PasswordGate>
  )
}

function TabBtn({ active, onClick, children }) {
  return (
    <button onClick={onClick} style={{
      padding: '7px 15px', borderRadius: T.rPill, border: 'none', cursor: 'pointer',
      fontSize: 14, fontWeight: active ? 700 : 500, whiteSpace: 'nowrap',
      fontFamily: 'inherit', transition: 'background .12s,color .12s',
      background: active ? T.a : 'transparent', color: active ? T.bg : T.n600,
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

      // 含「訂單總表（二）」頁籤 → 官網大筆訂單格式（多行合併）
      if (wb.Sheets['訂單總表（二）']) {
        const ws = wb.Sheets['訂單總表（二）']
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' })
          .filter(r => Object.values(r).some(v => v !== '' && v !== null))
        const parsed = parseOfficialBulk(rows)
        setPlatform('官網')
        setOrders(parsed)
        setBlocks(buildBlocks(parsed, mapping))
        setMsg(`已辨識為「官網」（訂單總表格式），共 ${parsed.length} 筆訂單。`)
        return
      }

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
    const seen = new Set()
    const rows = orders
      .filter(o => { if (seen.has(o.ref_no)) return false; seen.add(o.ref_no); return true })
      .map((o) => ({
        platform: o.platform, ref_no: o.ref_no, order_date: String(o.order_date || ''),
        contact: o.contact, address: o.address, phone: String(o.phone || ''), email: o.email,
        pay_method: o.pay_method, note: o.note, store: o.store, pkg_count: o.pkg_count || 1,
        tracking_no: String(o.tracking_no || ''), total: o.total || 0, shipping_fee: o.shipping_fee || 0, discount: o.discount ?? null,
      }))
    const { error } = await supabase.from('shipping_orders').upsert(rows, { onConflict: 'platform,ref_no' })
    if (error) { setMsg(`存檔失敗：${error.message}`); return }
    // 只對真正新增的訂單（recon_status 還是 null）才設「已出貨」，不蓋既有對帳狀態
    const platform = rows[0]?.platform
    const refNos = rows.map(r => r.ref_no)
    await supabase.from('shipping_orders')
      .update({ recon_status: '已出貨' })
      .eq('platform', platform)
      .in('ref_no', refNos)
      .is('recon_status', null)
    setMsg(`已存入 ${rows.length} 筆到資料庫。`)
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
                      <button onClick={() => setEditing({ ...r, id: undefined })} style={miniBtn}>複製</button>
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
  const [activeGateway, setActiveGateway] = useState('shopee')
  const [txRows, setTxRows] = useState(null)
  const [txFileName, setTxFileName] = useState('')
  const [txMsg, setTxMsg] = useState('')
  const [txResult, setTxResult] = useState(null)
  const [txVersion, setTxVersion] = useState(0)   // 回填完遞增，通知各頁籤重新載入訂單
  const txFileRef = useRef(null)

  function readTxFile(e) {
    readTianxinFile(e, setTxRows, setTxFileName)
  }

  async function handleTianxin() {
    if (!txRows) { setTxMsg('請先上傳檔案'); return }
    setTxMsg('比對中…'); setTxResult(null)

    // 每個客戶訂單只取第一個 SA 單號（一筆訂單可能有多列商品）
    const saMap = {}
    const invMap = {}
    for (const r of txRows) {
      const ref = String(r['客戶訂單'] || '').trim()
      const sa = String(r['單號'] || '').trim()
      const inv = String(r['發票號碼'] || '').trim()
      if (ref && sa.startsWith('SA') && !saMap[ref]) saMap[ref] = sa
      if (ref && inv && !invMap[ref]) invMap[ref] = inv
    }
    const pairs = Object.entries(saMap)
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
      const upd = { sa_no: sa }
      if (invMap[ref]) upd.order_invoice_no = invMap[ref]
      const { error: ue } = await supabase.from('shipping_orders').update(upd).eq('id', id)
      if (!ue) updated++
    }

    setTxResult({ total: pairs.length, updated, unmatched })
    const invCount = Object.keys(invMap).length
    setTxMsg(`${pairs.length} 筆訂單，回填 ${updated} 筆銷貨單號${invCount ? `・${invCount} 筆訂單發票號碼` : ''}，未對應 ${unmatched.length} 筆`)
    if (updated) setTxVersion(v => v + 1)   // 讓目前頁籤的訂單表反映剛回填的結果
  }

  // 天心銷貨單回填 —— 跨通路共用，位置在通路頁籤「之上」：
  // handleTianxin 撈訂單時不帶 platform 條件，一次上傳就比對全部平台的訂單編號，
  // 把 SA 單號與訂單發票號碼寫進對應訂單，不需要每個頁籤各傳一次。
  const tianxinPanel = (
    <div style={{ background: T.surface, borderRadius: T.rCard, boxShadow: T.shadowSm,
      padding: '18px 22px', marginBottom: 18, borderLeft: `3px solid ${T.a}` }}>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 6 }}>
        <span style={{ fontSize: 16, fontWeight: 700, color: T.navy }}>天心銷貨單</span>
        <span style={{ fontSize: 12, color: T.a, fontWeight: 600 }}>跨通路共用 · 一次上傳回填所有頁籤</span>
      </div>
      <div style={{ fontSize: 12, color: T.n600, lineHeight: 1.7, marginBottom: 10 }}>
        比對天心「客戶訂單」與各平台訂單編號，將「單號」寫入銷貨單號、「發票號碼」寫入訂單發票號碼。不分通路，全部平台一起回填。
      </div>
      <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
        <input ref={txFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={readTxFile} style={{ display: 'none' }} />
        <DropButton onClick={() => txFileRef.current.click()} filled={!!txRows}
          label={txFileName || '選擇天心銷貨單'}
          hint={txRows ? `✓ ${txRows.length} 列` : '尚未選擇檔案'} />
        <button onClick={handleTianxin} style={btnPri}>比對回填</button>
      </div>
      <PanelMsg text={txMsg} bad={/錯誤|找不到|請/} />
      {txResult?.unmatched?.length > 0 && (
        <p style={{ marginTop: 4, marginBottom: 0, fontSize: 12, color: C.warn }}>
          未對應：{txResult.unmatched.slice(0, 8).join('、')}
          {txResult.unmatched.length > 8 && `…等 ${txResult.unmatched.length} 筆`}
        </p>
      )}
    </div>
  )

  return (
    <div>
      {/* ── 天心銷貨單：跨通路共用，刻意放在通路切換列之上 ── */}
      {activeGateway !== '__guide__' && tianxinPanel}

      {/* ── 通路切換列 ── */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 20 }}>
        <span style={{ fontSize: 12, letterSpacing: '.08em', color: T.n600, marginRight: 4 }}>金流通路</span>
        {[...GATEWAY_LIST, { key: '__guide__', label: '說明', dot: T.gold }].map(g => {
          const on = activeGateway === g.key
          return (
            <button key={g.key} onClick={() => setActiveGateway(g.key)} aria-pressed={on}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 7, padding: '8px 15px',
                borderRadius: T.rPill, fontSize: 14, cursor: 'pointer', whiteSpace: 'nowrap',
                fontFamily: 'inherit',
                border: `1px solid ${on ? T.a : T.divider}`,
                background: on ? T.a : 'transparent',
                color: on ? T.bg : T.n700,
                transition: 'background .12s,color .12s,border-color .12s',
              }}>
              <span style={{ width: 7, height: 7, borderRadius: '50%', background: on ? T.bg : (g.dot || T.n400) }} />
              {g.label}
            </button>
          )
        })}
      </div>

      {activeGateway === '__guide__' ? (
        <div style={{ background: T.n100, borderRadius: T.rCard, boxShadow: T.shadowSm, padding: '24px 28px' }}>
          <h3 style={{ margin: '0 0 14px', fontSize: 22, color: T.navy }}>金流對帳流程說明</h3>
          <img
            src="recon-guide.png"
            alt="金流對帳說明"
            style={{ maxWidth: '100%', borderRadius: T.rPanel, display: 'block' }}
          />
        </div>
      ) : (
        <GatewayWorkspace
          gateway={activeGateway}
          key={activeGateway}
          txVersion={txVersion}
        />
      )}
    </div>
  )
}

function GatewayWorkspace({ gateway, txVersion }) {
  const gwInfo = GATEWAY_LIST.find(g => g.key === gateway) || {}
  const isTwoFile = !!gwInfo.twoFile
  const isLinePayOfficial = gateway === 'payuni_linepay'
  const isLineMallLinePay = gateway === 'linepay'
  const isLanxin = gateway === 'lanxin'
  const isPayuniCC = gateway === 'payuni_cc'
  const isShopee = gateway === 'shopee'
  const isMegabank = gateway === 'megabank'
  const isCoupang = gateway === 'coupang'
  // 玉山對帳一律採「手動勾選訂單」比對，七條金流都適用
  // （早期依撥款日自動篩選的路徑已無金流使用，相關分支保留但不會走到）
  const isManualSelection = isPayuniCC || isLineMallLinePay || isLanxin || isLinePayOfficial || isShopee || isCoupang || isMegabank
  const STATUSES = ['待出貨', '已出貨', '平台已結算', '已入帳', '已對帳']

  const [rows1, setRows1] = useState(null)
  const [rows2, setRows2] = useState(null)
  const [fileName1, setFileName1] = useState('')
  const [fileName2, setFileName2] = useState('')
  const [reconMsg, setReconMsg] = useState('')
  const [reconResult, setReconResult] = useState(null)
  const [createMissing, setCreateMissing] = useState(true)   // 對帳單比對不到的訂單自動建檔
  const [skipColCheck, setSkipColCheck] = useState(false)    // 略過上傳檔案的欄位格式檢查
  const [reconWarn, setReconWarn] = useState(null)           // 兆豐 6% 費率檢核結果
  const [reconSkipped, setReconSkipped] = useState(null)     // 對帳單中被略過的列（如付款取消）
  const fileRef1 = useRef(null)
  const fileRef2 = useRef(null)

  const [orders, setOrders] = useState([])
  const [filterStatus, setFilterStatus] = useState('')
  const [filterMonth, setFilterMonth] = useState('')
  const [onlyDiff, setOnlyDiff] = useState(false)
  const [sortCol, setSortCol] = useState('order_date')
  const [sortDir, setSortDir] = useState('asc')
  const [selectedIds, setSelectedIds] = useState(new Set())
  const [deleteMsg, setDeleteMsg] = useState('')
  const [editOrder, setEditOrder] = useState(null)
  const [editMsg, setEditMsg] = useState('')
  // 入帳差異的原因註記：點明細表「差異」欄開啟，寫入 diff_note
  // （不可用 note —— 那是出貨報表帶進來的買家訂單備註）
  const [diffNoteOrder, setDiffNoteOrder] = useState(null)
  const [diffNoteText, setDiffNoteText] = useState('')
  const [diffNoteMsg, setDiffNoteMsg] = useState('')
  // 差異視窗內的發票欄位（逐筆訂單，與群組層的手續費／代開發票分開）
  const [diffInvNo, setDiffInvNo] = useState('')
  const [diffInvDate, setDiffInvDate] = useState('')
  const [diffInvAmount, setDiffInvAmount] = useState('')
  const [diffInvUrl, setDiffInvUrl] = useState(null)
  const [diffInvUploading, setDiffInvUploading] = useState(false)
  const [diffInvUploadError, setDiffInvUploadError] = useState('')
  const diffInvFileRef = useRef(null)
  const [viewInvKey, setViewInvKey] = useState(null)
  const [viewTxInvKey, setViewTxInvKey] = useState(null)
  const [viewOrdInvKey, setViewOrdInvKey] = useState(null)
  const [ordInvDeleteConfirm, setOrdInvDeleteConfirm] = useState(false)
  const [ordInvPopupDate, setOrdInvPopupDate] = useState('')
  const [viewBankGroupKey, setViewBankGroupKey] = useState(null)
  const [bankGroupNote, setBankGroupNote] = useState('')
  const [invNote, setInvNote] = useState('')
  const [txInvNote, setTxInvNote] = useState('')
  const [invPdfUrl, setInvPdfUrl] = useState(null)
  const [txInvPdfUrl, setTxInvPdfUrl] = useState(null)
  const [invUploading, setInvUploading] = useState(false)
  const [txInvUploading, setTxInvUploading] = useState(false)
  const [invUploadError, setInvUploadError] = useState('')
  const [txInvUploadError, setTxInvUploadError] = useState('')
  const invPdfRef = useRef(null)
  const txInvPdfRef = useRef(null)
  const [invDeleteConfirm, setInvDeleteConfirm] = useState(false)
  const [txInvDeleteConfirm, setTxInvDeleteConfirm] = useState(false)
  const [invPopupDate, setInvPopupDate] = useState('')
  const [invPopupAmount, setInvPopupAmount] = useState('')
  const [txInvPopupAmount, setTxInvPopupAmount] = useState('')
  const [txInvPopupDate, setTxInvPopupDate] = useState('')

  const [bankRows, setBankRows] = useState([])
  const [bankFileName, setBankFileName] = useState('')
  const [bankSel, setBankSel] = useState({})
  const [bankExpanded, setBankExpanded] = useState({})
  const [bankMsg, setBankMsg] = useState({})
  const [bankEntryChecked, setBankEntryChecked] = useState(new Set())
  const [bankCCOrderSel, setBankCCOrderSel] = useState({})
  const [confirmedGroupExp, setConfirmedGroupExp] = useState({})
  const [releasingKey, setReleasingKey] = useState(null)   // 正在解除的入帳群組
  const bankFileRef = useRef(null)

  const [ordInvRows, setOrdInvRows] = useState(null)
  const [ordInvFileName, setOrdInvFileName] = useState('')
  const [ordInvMsg, setOrdInvMsg] = useState('')
  const ordInvFileRef = useRef(null)

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

  const [inv3Method, setInv3Method] = useState('auto')
  const [inv3No, setInv3No] = useState('')
  const [inv3Date, setInv3Date] = useState('')
  const [inv3Amount, setInv3Amount] = useState('')
  const [inv3From, setInv3From] = useState('')
  const [inv3To, setInv3To] = useState('')
  const [inv3Preview, setInv3Preview] = useState(null)
  const [inv3Msg, setInv3Msg] = useState('')
  const [checked3Ids, setChecked3Ids] = useState(new Set())


  const [shopeeOrdRows, setShopeeOrdRows] = useState(null)
  const [shopeeOrdFileName, setShopeeOrdFileName] = useState('')
  const [shopeeOrdMsg, setShopeeOrdMsg] = useState('')
  const shopeeOrdFileRef = useRef(null)


  const [ordInvEntryNo, setOrdInvEntryNo] = useState('')
  const [ordInvEntryDate, setOrdInvEntryDate] = useState('')
  const [ordInvEntryAmount, setOrdInvEntryAmount] = useState('')
  const [ordInvEntryMethod, setOrdInvEntryMethod] = useState('auto')
  const [ordInvEntryFrom, setOrdInvEntryFrom] = useState('')
  const [ordInvEntryTo, setOrdInvEntryTo] = useState('')
  const [ordInvEntryPreview, setOrdInvEntryPreview] = useState(null)
  const [ordInvEntryChecked, setOrdInvEntryChecked] = useState(new Set())
  const [ordInvEntryMsg, setOrdInvEntryMsg] = useState('')

  // SOP 依「對帳步驟」分卡儲存：{ [stepKey]: { html, img } }
  // 沿用 gateway_sops.html_content 欄位存 JSON（{v:1,steps}），不必動 DB schema
  const [sopSteps, setSopSteps] = useState({})
  const [sopLegacy, setSopLegacy] = useState('')   // 舊版單一 HTML，塞進第一步以免遺失
  const [sopEditing, setSopEditing] = useState(false)
  const [sopSaving, setSopSaving] = useState(false)
  const [sopUploading, setSopUploading] = useState(null)
  const [sopMsg, setSopMsg] = useState('')
  const [sopLinkForm, setSopLinkForm] = useState(false)
  const [sopLinkUrl, setSopLinkUrl] = useState('')
  const [sopLinkText, setSopLinkText] = useState('')
  const sopRefs = useRef({})
  const sopSavedRange = useRef(null)
  // 截圖的 Storage 刪除一律延後到「完成/取消」才執行，避免取消後 DB 仍指向已刪檔案而破圖
  const sopNewPaths = useRef([])   // 本次編輯上傳的：取消才刪
  const sopOldPaths = useRef([])   // 本次編輯換掉/移除的舊檔：存檔才刪

  useEffect(() => { loadOrders() }, [])
  useEffect(() => {
    setInvDeleteConfirm(false)
    if (viewInvKey) {
      const o = orders.find(x => x.fee_invoice_no === viewInvKey)
      setInvNote(o?.fee_invoice_note ?? localStorage.getItem(`inv_note_${viewInvKey}`) ?? '')
      setInvPdfUrl(o?.fee_invoice_pdf_url ?? null)
      setInvPopupDate(o?.fee_invoice_date ?? '')
      setInvPopupAmount(o?.fee_invoice_amount != null ? String(o.fee_invoice_amount) : '')
    }
  }, [viewInvKey])
  useEffect(() => {
    setTxInvDeleteConfirm(false)
    if (viewTxInvKey) {
      const o = orders.find(x => x.tx_fee_invoice_no === viewTxInvKey)
      setTxInvNote(o?.tx_fee_invoice_note ?? localStorage.getItem(`txinv_note_${viewTxInvKey}`) ?? '')
      setTxInvPdfUrl(o?.tx_fee_invoice_pdf_url ?? null)
      setTxInvPopupAmount(localStorage.getItem(`txinv_amount_${viewTxInvKey}`) ?? '')
      setTxInvPopupDate(localStorage.getItem(`txinv_date_${viewTxInvKey}`) ?? '')
    }
  }, [viewTxInvKey])
  useEffect(() => {
    if (viewBankGroupKey) {
      setBankGroupNote(localStorage.getItem(`bankgroup_note_${gateway}_${viewBankGroupKey}`) ?? '')
    }
  }, [viewBankGroupKey])
  useEffect(() => {
    setSopEditing(false); setSopLinkForm(false); setSopMsg('')
    sopNewPaths.current = []; sopOldPaths.current = []
    reloadSop()
  }, [gateway])
  // 元件卸載（例如切換通路）時，把這次編輯中上傳、但沒存檔的截圖清掉
  useEffect(() => () => {
    const orphans = sopNewPaths.current
    if (orphans.length) supabase.storage.from('invoices').remove(orphans)
  }, [])
  // 進入編輯模式時把內容灌進各步驟的 contentEditable（之後不再同步，避免蓋掉打字中的內容）
  useEffect(() => {
    if (!sopEditing) return
    let first = null
    for (const key of Object.keys(sopRefs.current)) {
      const el = sopRefs.current[key]
      if (!el) continue
      el.innerHTML = sopStepData(key).html
      if (!first) first = el
    }
    first?.focus()
  }, [sopEditing])

  async function loadOrders() {
    const data = await loadGatewayOrders(supabase, gateway)
    setOrders(data)
  }

  // 上方全域區的天心回填會動到本頁籤的訂單，回填完重新載入才看得到 SA 單號／發票號碼
  useEffect(() => { if (txVersion) loadOrders() }, [txVersion])

  // 已確認入帳群組：把訂單退回未入帳，讓使用者可以修正勾錯的歸戶
  // 清掉 in_date / actual_in / bank_deposit，狀態退回「平台已結算」（群組只收 已入帳，不會動到 已對帳）
  async function releaseFromGroup(ids, label) {
    if (!ids.length) return
    if (!window.confirm(`確定將 ${label} 退回未入帳？\n入帳日、實際入帳金額會被清空，狀態改回「平台已結算」。`)) return
    setReleasingKey(label)
    const { error } = await supabase.from('shipping_orders')
      .update({ in_date: null, actual_in: null, bank_deposit: null, recon_status: '平台已結算' })
      .in('id', ids)
    setReleasingKey(null)
    if (error) { window.alert('解除失敗：' + error.message); return }
    await loadOrders()
  }

  function readFile(e, setRows, setFileName) {
    const f = e.target.files?.[0]
    if (!f) return
    setFileName(f.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: 'array' })
      let ws
      if (wb.Sheets['Income']) {
        ws = wb.Sheets['Income']
      } else {
        ws = wb.Sheets[wb.SheetNames[0]]
      }
      setRows(sheetToRows(ws))
    }
    reader.readAsArrayBuffer(f)
  }

  function readBankFile(e) {
    const f = e.target.files?.[0]
    if (!f) return
    setBankFileName(f.name)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
      // 動態找欄位索引，相容有無「序號」欄的不同格式
      const headerIdx = all.findIndex(r => r.includes('帳務日期'))
      const hdr = all[headerIdx] || []
      const ci = {
        date:    hdr.indexOf('帳務日期'),
        actual:  hdr.indexOf('實際交易日期'),
        summary: hdr.indexOf('摘要'),
        deposit: hdr.indexOf('存'),
        note:    hdr.indexOf('備註'),
        account: hdr.indexOf('轉出入銀行代號/帳號'),
      }
      const parsed = all.slice(headerIdx + 1)
        .filter(r => {
          if (isShopee) return String(r[ci.note] || '').toUpperCase().includes('SHOPEE') || String(r[ci.account] || '').includes('808/0370979139156')
          if (isPayuniCC) return String(r[ci.summary] || '').includes('ＰＡＹＵ')
          const a = String(r[ci.account] || '')
          if (isLineMallLinePay) return a.includes('387/0000000060558379') || a.includes('808/1229940024585')
          if (isLanxin) return a.includes('008/0000158100035101')
          if (isLinePayOfficial) return a.includes('808/1229940024585')
          return false
        })
        .map(r => ({ date: excelDate(r[ci.date]), actualDate: excelDate(r[ci.actual]), deposit: parseFloat(r[ci.deposit]) || 0, summary: String(r[ci.summary] || ''), account: String(r[ci.account] || '') }))
        .filter(r => r.deposit > 0)
        .sort((a, b) => a.date.localeCompare(b.date))
      setBankRows(parsed)
      setBankSel({})
      setBankExpanded({})
      setBankEntryChecked(new Set())
      setBankCCOrderSel({})
    }
    reader.readAsArrayBuffer(f)
  }

  async function handleShopeeOrdImport() {
    if (!shopeeOrdRows) { setShopeeOrdMsg('請先上傳檔案'); return }
    setShopeeOrdMsg('處理中…')
    const orderMap = {}
    for (const r of shopeeOrdRows) {
      // A欄(0)=平台訂單編號、F欄(5)=訂單日期、K欄(10)=應收
      const ref = String(r[0] || '').trim()
      if (!ref || orderMap[ref]) continue
      const rawDate = String(r[5] || '').trim()
      orderMap[ref] = {
        ref_no: ref,
        order_date: rawDate ? rawDate.slice(0, 10) : null,
        total: parseFloat(r[10]) || null,
        platform: '蝦皮',
        pay_method: String(r[48] || '').trim() || null,
      }
    }
    const toCheck = Object.values(orderMap)
    if (!toCheck.length) { setShopeeOrdMsg('找不到有效訂單資料'); return }
    const { data: existing } = await supabase
      .from('shipping_orders').select('ref_no')
      .in('ref_no', toCheck.map(o => o.ref_no))
    const existingRefs = new Set((existing || []).map(o => o.ref_no))
    const toInsert = toCheck.filter(o => !existingRefs.has(o.ref_no))
    const skipped = toCheck.length - toInsert.length
    if (!toInsert.length) {
      setShopeeOrdMsg(`${toCheck.length} 筆已全部匯入過，無新增`)
      return
    }
    const { error } = await supabase.from('shipping_orders').insert(toInsert)
    if (error) { setShopeeOrdMsg('錯誤：' + error.message); return }
    setShopeeOrdMsg(`新增 ${toInsert.length} 筆${skipped ? `，略過 ${skipped} 筆（已存在）` : ''}`)
    loadOrders()
  }

  function runOrdInvEntryPreview() {
    const from = ordInvEntryFrom
    const to = ordInvEntryTo
    const filtered = orders.filter(o => {
      const d = (o.order_date || '').slice(0, 10)
      if (from && d < from) return false
      if (to && d > to) return false
      return true
    })
    setOrdInvEntryPreview(filtered)
    setOrdInvEntryMsg('')
  }

  async function runApplyOrdInvEntry() {
    const ids = ordInvEntryMethod === 'auto'
      ? (ordInvEntryPreview || []).map(o => o.id)
      : [...ordInvEntryChecked]
    if (!ordInvEntryNo) { setOrdInvEntryMsg('請輸入發票號碼'); return }
    if (!ids.length) { setOrdInvEntryMsg('沒有選取任何訂單'); return }
    const { error } = await supabase.from('shipping_orders').update({
      order_invoice_no: ordInvEntryNo,
      order_invoice_date: ordInvEntryDate || null,
      order_invoice_amount: ordInvEntryAmount ? Number(ordInvEntryAmount) : null,
    }).in('id', ids)
    if (error) { setOrdInvEntryMsg('錯誤：' + error.message); return }
    setOrdInvEntryMsg(`已套用至 ${ids.length} 筆`)
    loadOrders()
  }

  // 單一步驟的 SOP 內容；舊版單一 HTML 併入第一步，避免遺失
  function sopStepData(key) {
    const d = sopSteps[key]
    if (d) return { html: d.html || '', img: d.img || null }
    if (sopLegacy && key === stepDefs[0]?.key) return { html: sopLegacy, img: null }
    return { html: '', img: null }
  }

  // 把各步驟編輯器目前的 DOM 內容收回 state（上傳截圖前後都要，避免打字中的內容被覆蓋）
  function collectSopSteps(base) {
    const next = { ...base }
    for (const s of stepDefs) {
      const el = sopRefs.current[s.key]
      const prev = next[s.key] || sopStepData(s.key)
      next[s.key] = { html: el ? el.innerHTML : (prev.html || ''), img: prev.img || null }
    }
    return next
  }

  // 從 DB 讀回這條金流的 SOP（切換通路、取消編輯都用它）
  async function reloadSop() {
    const { data } = await supabase.from('gateway_sops')
      .select('html_content').eq('gateway', gateway).maybeSingle()
    const raw = data?.html_content || ''
    setSopSteps({}); setSopLegacy('')
    if (!raw) return
    try {
      const p = JSON.parse(raw)
      if (p && p.v === 1 && p.steps) { setSopSteps(p.steps); return }
    } catch { /* 不是 JSON → 當作舊版 HTML */ }
    setSopLegacy(raw)
  }

  // 舊檔標記待刪；若舊檔正好是本次編輯剛傳的，直接刪掉不必等存檔
  function markSopOldPath(oldPath) {
    if (!oldPath) return
    const i = sopNewPaths.current.indexOf(oldPath)
    if (i >= 0) {
      sopNewPaths.current.splice(i, 1)
      supabase.storage.from('invoices').remove([oldPath])
    } else if (!sopOldPaths.current.includes(oldPath)) {
      sopOldPaths.current.push(oldPath)
    }
  }

  async function saveSop() {
    const steps = collectSopSteps(sopSteps)
    setSopSaving(true); setSopMsg('')
    const { error } = await supabase.from('gateway_sops')
      .upsert({ gateway, html_content: JSON.stringify({ v: 1, steps }), updated_at: new Date().toISOString() },
        { onConflict: 'gateway' })
    setSopSaving(false)
    if (error) { setSopMsg('儲存失敗：' + error.message); return }
    // 存檔成功才真的刪掉被換掉／移除的舊截圖
    const stale = sopOldPaths.current
    sopNewPaths.current = []; sopOldPaths.current = []
    if (stale.length) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove(stale)
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    setSopSteps(steps)
    setSopLegacy('')
    setSopEditing(false)
    setSopLinkForm(false)
  }

  async function cancelSop() {
    // 取消：刪掉本次上傳的檔，舊檔原封不動，再從 DB 讀回原內容
    const orphans = sopNewPaths.current
    sopNewPaths.current = []; sopOldPaths.current = []
    setSopEditing(false)
    setSopLinkForm(false)
    setSopMsg('')
    if (orphans.length) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove(orphans)
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    await reloadSop()
  }

  // 每步一張截圖，存 Storage 的 sop/ 夾，URL 記在該步驟
  async function uploadSopStepImage(key, file) {
    if (!file) return
    setSopUploading(key); setSopMsg('')
    const ext = (file.name.split('.').pop() || 'png').toLowerCase()
    const path = `sop/${gateway}_${key}_${Date.now()}.${ext}`
    const { error: upErr } = await supabase.storage.from('invoices').upload(path, file)
    if (upErr) { setSopUploading(null); setSopMsg(`截圖上傳失敗：${upErr.message}`); return }
    sopNewPaths.current.push(path)
    markSopOldPath(storagePath(sopStepData(key).img))
    const { data } = supabase.storage.from('invoices').getPublicUrl(path)
    setSopSteps(p => {
      const next = collectSopSteps(p)
      next[key] = { ...next[key], img: data.publicUrl }
      return next
    })
    setSopUploading(null)
  }

  function removeSopStepImage(key) {
    markSopOldPath(storagePath(sopStepData(key).img))
    setSopSteps(p => {
      const next = collectSopSteps(p)
      next[key] = { ...next[key], img: null }
      return next
    })
  }

  // 關閉抽屜：若還在編輯中，等同按「取消」（清掉未存檔的截圖、還原內容）
  function closeSopDrawer() {
    if (sopEditing) cancelSop()
    else { setSopLinkForm(false); setSopMsg('') }
    setSopOpen(false)
  }

  // 記住游標停在哪一張步驟卡，插入連結時才知道要塞回哪個編輯器
  const sopFocusEl = useRef(null)

  function openSopLinkForm() {
    const sel = window.getSelection()
    if (sel && sel.rangeCount > 0) {
      sopSavedRange.current = sel.getRangeAt(0).cloneRange()
      setSopLinkText(sel.toString())
    } else {
      setSopLinkText('')
    }
    setSopLinkUrl('')
    setSopLinkForm(true)
  }

  function insertSopLink() {
    const url = sopLinkUrl.trim()
    if (!url) { setSopLinkForm(false); return }
    const text = sopLinkText.trim() || url
    const target = sopFocusEl.current || sopRefs.current[stepDefs[0]?.key]
    target?.focus()
    if (sopSavedRange.current) {
      const sel = window.getSelection()
      if (sel) { sel.removeAllRanges(); sel.addRange(sopSavedRange.current) }
      sopSavedRange.current = null
    }
    document.execCommand('insertHTML', false,
      `<a href="${url}" target="_blank" rel="noopener noreferrer" style="color:${T.a700};text-decoration:underline">${text}</a>`)
    setSopLinkForm(false)
    setSopLinkUrl('')
    setSopLinkText('')
  }

  function toggleBankSel(idx, ordId) {
    setBankSel(prev => {
      const s = new Set(prev[idx] || [])
      if (s.has(ordId)) s.delete(ordId); else s.add(ordId)
      return { ...prev, [idx]: s }
    })
  }

  async function confirmBankEntry(idx, br, dateOrders) {
    const ordersToUpdate = isManualSelection
      ? dateOrders.filter(o => bankCCOrderSel[idx]?.has(String(o.id)))
      : dateOrders
    if (ordersToUpdate.length === 0) return
    setBankMsg(p => ({ ...p, [idx]: '寫入中…' }))
    let hasErr = false
    for (const o of ordersToUpdate) {
      const { error } = await supabase
        .from('shipping_orders')
        .update({ in_date: br.date, actual_in: o.payable, bank_deposit: br.deposit, recon_status: '已入帳' })
        .eq('id', o.id)
      if (error) { hasErr = true; break }
    }
    if (!hasErr) {
      setBankMsg(p => ({ ...p, [idx]: `✓ 已回填 ${ordersToUpdate.length} 筆` }))
      await loadOrders()
    } else {
      setBankMsg(p => ({ ...p, [idx]: '❌ 寫入失敗' }))
    }
  }

  async function handleOrdInvImport() {
    if (!ordInvRows) { setOrdInvMsg('請先上傳檔案'); return }
    setOrdInvMsg('比對中…')
    const map = {}
    for (const r of ordInvRows) {
      const inv = String(r['發票號碼'] || '').trim()
      if (!inv) continue
      const key = String(r['客戶訂單'] || r['訂單編號'] || r['參照編號'] || r['平台訂單編號'] || '').trim()
      if (key && !map[key]) map[key] = inv
    }
    const pairs = Object.entries(map)
    if (!pairs.length) { setOrdInvMsg('找不到「發票號碼」欄位或訂單編號，請確認欄位名稱'); return }
    const byRef = {}
    for (const o of orders) byRef[o.ref_no] = o.id
    let updated = 0
    const unmatched = []
    for (const [ref, inv] of pairs) {
      const id = byRef[ref]
      if (!id) { unmatched.push(ref); continue }
      const { error } = await supabase.from('shipping_orders').update({ order_invoice_no: inv }).eq('id', id)
      if (!error) updated++
    }
    const msg = `比對 ${pairs.length} 筆，回填 ${updated} 筆，未對應 ${unmatched.length} 筆`
    setOrdInvMsg(unmatched.length ? msg + `（${unmatched.slice(0, 3).join('、')}…）` : msg)
    loadOrders()
  }

  async function handleReconcile() {
    setReconMsg('比對中…')
    try {
      let parsed
      if (isTwoFile) {
        if (!rows1 || !rows2) {
          setReconMsg(isMegabank ? '請分別上傳訂單明細報表與手續費報表' : '請分別上傳 D-1 和 D-2 兩份對帳單')
          return
        }
        // 欄位檢查：擋掉上傳錯報表（否則會算出一堆 0 並寫進 DB）
        const bad = skipColCheck ? null
          : isMegabank ? checkMegabankColumns(rows1, rows2)
          : checkDualReconColumns(rows1, rows2)
        if (bad) {
          setReconMsg(`錯誤：檔案格式不符 — ${bad.join('、')}。請確認上傳的是正確的對帳單；確定要繼續請勾選下方「略過格式檢查」。`)
          return
        }
        parsed = isMegabank
          ? parseMegabankRecon(rows1, rows2)
          : parseOfficialLinePayReconDual(rows1, rows2)
      } else {
        if (!rows1) { setReconMsg('請先上傳對帳單'); return }
        const missing = skipColCheck ? null : checkReconColumns(gateway, rows1)
        if (missing) {
          setReconMsg(`錯誤：這份檔案不像「${gwInfo.label}」的對帳單，缺少必要欄位 ${missing.join('、')}。請確認上傳的報表是否正確；確定要繼續請勾選下方「略過格式檢查」。`)
          return
        }
        parsed = RECON_PARSERS[gateway](rows1)
      }
      // 兆豐手續費固定為訂單金額的 6%，不符就示警（多半是手續費報表缺列或期間沒對齊）
      setReconWarn(isMegabank ? megabankRateWarnings(parsed) : null)
      // 對帳單中未成立的交易（如付款取消）不寫入，但要列出來讓使用者知道略過了什麼
      setReconSkipped(parsed.skipped?.length ? parsed.skipped : null)
      const result = await reconcile(supabase, gateway, parsed, { createMissing })
      setReconResult(result)
      const parts = [`比對 ${parsed.length} 筆`, `回填 ${result.updated} 筆`]
      if (createMissing) parts.push(`新增 ${result.inserted} 筆`)
      else parts.push(`未對應 ${result.unmatched.length} 筆`)
      setReconMsg(parts.join('、'))
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

  const manualOrders = orders.filter(o => !o.fee_invoice_no).sort((a, b) => (a.order_date || '').localeCompare(b.order_date || ''))
  const checkedOrders = manualOrders.filter(o => checkedIds.has(o.id))
  const manualFeeSum = Math.round(checkedOrders.reduce((s, o) => s + (o.fee_total || 0), 0) * 100) / 100

  const manual3Orders = orders.filter(o => !o.tx_fee_invoice_no).sort((a, b) => (a.order_date || '').localeCompare(b.order_date || ''))
  const checked3Orders = manual3Orders.filter(o => checked3Ids.has(o.id))
  const manual3TxFeeSum = Math.round(checked3Orders.reduce((s, o) => s + (o.tx_fee || 0), 0) * 100) / 100

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

  function switchMethod3(v) { setInv3Method(v); setInv3Preview(null); setChecked3Ids(new Set()); setInv3Msg('') }
  function toggleCheck3(id) {
    setChecked3Ids(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n })
  }
  function toggleAll3() {
    setChecked3Ids(prev => prev.size === manual3Orders.length ? new Set() : new Set(manual3Orders.map(o => o.id)))
  }
  async function runInv3PreviewAuto() {
    if (!inv3From || !inv3To) { setInv3Msg('請填寫涵蓋期間'); return }
    setInv3Msg('查詢中…'); setInv3Preview(null)
    const filtered = orders.filter(o => {
      const d = (o.in_date || o.order_date || '').slice(0, 10)
      return d >= inv3From && d <= inv3To
    })
    const txFeeSum = Math.round(filtered.reduce((s, o) => s + (o.tx_fee || 0), 0) * 100) / 100
    setInv3Preview({ orders: filtered, txFeeSum, method: 'auto' })
    setInv3Msg(filtered.length ? '' : '查無符合期間的訂單')
  }
  async function runApplyTxFeeInvoice() {
    const amount = parseFloat(inv3Amount) || 0
    const isAutoMode = inv3Preview?.method === 'auto'
    const orderIds = isAutoMode ? inv3Preview.orders.map(o => o.id) : [...checked3Ids]
    const txFeeSum = isAutoMode ? inv3Preview.txFeeSum : manual3TxFeeSum
    if (!orderIds.length) { setInv3Msg('請先選取訂單'); return }
    const isMatch = amount > 0 && Math.abs(amount - txFeeSum) < 0.01
    const { error } = await supabase.from('shipping_orders').update({ tx_fee_invoice_no: inv3No || null }).in('id', orderIds)
    if (error) { setInv3Msg('錯誤：' + error.message); return }
    if (inv3No) localStorage.setItem(`txinv_amount_${inv3No}`, String(amount))
    setInv3Msg(`已套用至 ${orderIds.length} 筆（${isMatch ? '相符' : '有差異'}）`)
    setInv3Preview(null); setChecked3Ids(new Set()); loadOrders()
  }

  async function saveInvNote(note) {
    localStorage.setItem(`inv_note_${viewInvKey}`, note)
    await supabase.from('shipping_orders').update({ fee_invoice_note: note || null }).eq('fee_invoice_no', viewInvKey)
  }

  async function saveTxInvNote(note) {
    localStorage.setItem(`txinv_note_${viewTxInvKey}`, note)
    await supabase.from('shipping_orders').update({ tx_fee_invoice_note: note || null }).eq('tx_fee_invoice_no', viewTxInvKey)
  }

  async function saveInvPopupDate(date) {
    await supabase.from('shipping_orders').update({ fee_invoice_date: date || null }).eq('fee_invoice_no', viewInvKey)
    loadOrders()
  }

  async function saveInvPopupAmount(amount) {
    const amt = amount !== '' ? parseFloat(amount) : null
    const feeSum = orders.filter(o => o.fee_invoice_no === viewInvKey).reduce((s, o) => s + (o.fee_total || 0), 0)
    const invoiceCheck = amt != null ? (Math.abs(amt - feeSum) < 0.01 ? '相符' : '有差異') : null
    await supabase.from('shipping_orders').update({ fee_invoice_amount: amt, invoice_check: invoiceCheck }).eq('fee_invoice_no', viewInvKey)
    loadOrders()
  }

  function saveTxInvPopupAmount(amount) {
    localStorage.setItem(`txinv_amount_${viewTxInvKey}`, amount)
  }

  function saveTxInvPopupDate(date) {
    localStorage.setItem(`txinv_date_${viewTxInvKey}`, date)
  }

  function storagePath(url) {
    if (!url) return null
    const marker = '/object/public/invoices/'
    const idx = url.indexOf(marker)
    return idx >= 0 ? url.slice(idx + marker.length) : null
  }

  async function uploadInvPdf(e) {
    const f = e.target.files?.[0]; if (!f) return
    setInvUploading(true); setInvUploadError('')
    const oldPath = storagePath(invPdfUrl)
    if (oldPath) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove([oldPath])
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    const path = `fee/${viewInvKey}_${Date.now()}.pdf`
    const { error: upErr } = await supabase.storage.from('invoices').upload(path, f)
    if (upErr) { setInvUploadError(`Storage 上傳失敗：${upErr.message}`); setInvUploading(false); return }
    const { data } = supabase.storage.from('invoices').getPublicUrl(path)
    const { error: dbErr } = await supabase.from('shipping_orders').update({ fee_invoice_pdf_url: data.publicUrl }).eq('fee_invoice_no', viewInvKey)
    if (dbErr) { setInvUploadError(`DB 更新失敗：${dbErr.message}`); setInvUploading(false); return }
    setInvPdfUrl(data.publicUrl); setInvUploading(false); loadOrders()
  }

  async function uploadTxInvPdf(e) {
    const f = e.target.files?.[0]; if (!f) return
    setTxInvUploading(true); setTxInvUploadError('')
    const oldPath = storagePath(txInvPdfUrl)
    if (oldPath) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove([oldPath])
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    const path = `txfee/${viewTxInvKey}_${Date.now()}.pdf`
    const { error: upErr } = await supabase.storage.from('invoices').upload(path, f)
    if (upErr) { setTxInvUploadError(`Storage 上傳失敗：${upErr.message}`); setTxInvUploading(false); return }
    const { data } = supabase.storage.from('invoices').getPublicUrl(path)
    const { error: dbErr } = await supabase.from('shipping_orders').update({ tx_fee_invoice_pdf_url: data.publicUrl }).eq('tx_fee_invoice_no', viewTxInvKey)
    if (dbErr) { setTxInvUploadError(`DB 更新失敗：${dbErr.message}`); setTxInvUploading(false); return }
    setTxInvPdfUrl(data.publicUrl); setTxInvUploading(false); loadOrders()
  }

  async function deleteInvoice() {
    const ids = orders.filter(o => o.fee_invoice_no === viewInvKey).map(o => o.id)
    const existingUrl = orders.find(o => o.fee_invoice_no === viewInvKey)?.fee_invoice_pdf_url
    const { error } = await supabase.from('shipping_orders')
      .update({ fee_invoice_no: null, fee_invoice_amount: null, fee_invoice_date: null, invoice_check: null, fee_invoice_note: null, fee_invoice_pdf_url: null })
      .in('id', ids)
    if (error) { setInvDeleteConfirm(false); return }
    const oldPath = storagePath(existingUrl)
    if (oldPath) await supabase.storage.from('invoices').remove([oldPath])
    localStorage.removeItem(`inv_note_${viewInvKey}`)
    setViewInvKey(null); loadOrders()
  }

  async function deleteTxInvoice() {
    const ids = orders.filter(o => o.tx_fee_invoice_no === viewTxInvKey).map(o => o.id)
    const existingUrl = orders.find(o => o.tx_fee_invoice_no === viewTxInvKey)?.tx_fee_invoice_pdf_url
    const { error } = await supabase.from('shipping_orders')
      .update({ tx_fee_invoice_no: null, tx_fee_invoice_note: null, tx_fee_invoice_pdf_url: null })
      .in('id', ids)
    if (error) { setTxInvDeleteConfirm(false); return }
    const oldPath = storagePath(existingUrl)
    if (oldPath) await supabase.storage.from('invoices').remove([oldPath])
    localStorage.removeItem(`txinv_note_${viewTxInvKey}`)
    localStorage.removeItem(`txinv_amount_${viewTxInvKey}`)
    localStorage.removeItem(`txinv_date_${viewTxInvKey}`)
    setViewTxInvKey(null); loadOrders()
  }

  function openDiffNote(o) {
    setDiffNoteOrder(o)
    setDiffNoteText(o.diff_note || '')
    setDiffInvNo(o.diff_invoice_no || '')
    setDiffInvDate(o.diff_invoice_date || '')
    setDiffInvAmount(o.diff_invoice_amount ?? '')
    setDiffInvUrl(o.diff_invoice_pdf_url || null)
    setDiffNoteMsg(''); setDiffInvUploadError('')
  }
  async function saveDiffNote() {
    setDiffNoteMsg('儲存中…')
    const { error } = await supabase.from('shipping_orders')
      .update({
        diff_note: diffNoteText.trim() || null,
        diff_invoice_no: diffInvNo.trim() || null,
        diff_invoice_date: diffInvDate || null,
        // 數字欄位不可傳空字串
        diff_invoice_amount: diffInvAmount !== '' && diffInvAmount != null ? parseFloat(diffInvAmount) : null,
      })
      .eq('id', diffNoteOrder.id)
    if (error) { setDiffNoteMsg('錯誤：' + error.message); return }
    setDiffNoteOrder(null); setDiffNoteMsg(''); loadOrders()
  }
  // 附件沿用 fee/txfee 的做法：直接上傳並寫回 DB，不等按「儲存」
  async function uploadDiffInvFile(e) {
    const f = e.target.files?.[0]; if (!f) return
    setDiffInvUploading(true); setDiffInvUploadError('')
    const oldPath = storagePath(diffInvUrl)
    if (oldPath) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove([oldPath])
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    const ext = (f.name.match(/\.[a-z0-9]+$/i) || ['.pdf'])[0]
    const path = `diff/${diffNoteOrder.ref_no}_${Date.now()}${ext}`
    const { error: upErr } = await supabase.storage.from('invoices').upload(path, f)
    if (upErr) { setDiffInvUploadError(`Storage 上傳失敗：${upErr.message}`); setDiffInvUploading(false); return }
    const { data } = supabase.storage.from('invoices').getPublicUrl(path)
    const { error: dbErr } = await supabase.from('shipping_orders')
      .update({ diff_invoice_pdf_url: data.publicUrl }).eq('id', diffNoteOrder.id)
    if (dbErr) { setDiffInvUploadError(`DB 更新失敗：${dbErr.message}`); setDiffInvUploading(false); return }
    setDiffInvUrl(data.publicUrl); setDiffInvUploading(false); loadOrders()
  }
  async function removeDiffInvFile() {
    setDiffInvUploading(true); setDiffInvUploadError('')
    const oldPath = storagePath(diffInvUrl)
    const { error: dbErr } = await supabase.from('shipping_orders')
      .update({ diff_invoice_pdf_url: null }).eq('id', diffNoteOrder.id)
    if (dbErr) { setDiffInvUploadError(`DB 更新失敗：${dbErr.message}`); setDiffInvUploading(false); return }
    if (oldPath) {
      const { error: rmErr } = await supabase.storage.from('invoices').remove([oldPath])
      if (rmErr) console.warn('Storage remove:', rmErr.message)
    }
    setDiffInvUrl(null); setDiffInvUploading(false); loadOrders()
  }

  async function saveEditOrder(updates) {
    setEditMsg('儲存中…')
    const { error } = await supabase
      .from('shipping_orders')
      .update({
        sa_no: updates.sa_no || null,
        recon_status: updates.recon_status || null,
        note: updates.note || null,
        order_invoice_no: updates.order_invoice_no || null,
        fee_invoice_no: updates.fee_invoice_no || null,
        tx_fee_invoice_no: updates.tx_fee_invoice_no || null,
        // 數字欄位：空字串要轉成 null，不能直接送給 Supabase
        total: updates.total !== '' && updates.total != null ? parseFloat(updates.total) : null,
        fee_total: updates.fee_total !== '' && updates.fee_total != null ? parseFloat(updates.fee_total) : null,
        payable: updates.payable !== '' && updates.payable != null ? parseFloat(updates.payable) : null,
        actual_in: updates.actual_in !== '' && updates.actual_in != null ? parseFloat(updates.actual_in) : null,
        in_date: updates.in_date || null,
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
      銷貨單號: o.sa_no ?? '', 訂單發票號碼: o.order_invoice_no ?? '',
      ...(isShopee ? {} : { 對應碼: o.tx_code ?? '' }),
      平台訂單編號: o.ref_no, 訂單日期: o.order_date ?? '', 應收: o.total, 手續費: o.fee_total ?? '',
      ...(isShopee ? {} : { 交易處理費: o.tx_fee ?? '' }),
      應入帳: o.payable ?? '', 實際入帳: o.actual_in ?? '', 入帳日: o.in_date ?? '',
      狀態: o.recon_status,
      手續費發票號碼: o.fee_invoice_no ?? '', 手續費發票備注: o.fee_invoice_note ?? '',
      ...(isLinePayOfficial ? { 交易處理費發票號碼: o.tx_fee_invoice_no ?? '', 交易處理費發票備注: o.tx_fee_invoice_note ?? '' } : {}),
    }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), gwInfo.label || '對帳')
    XLSX.writeFile(wb, `對帳_${gwInfo.label}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  const months = [...new Set(orders.map(o => (o.order_date || '').slice(0, 7)).filter(Boolean))].sort().reverse()

  const SORT_KEY = {
    '銷貨單號': 'sa_no', '訂單發票號碼': 'order_invoice_no', '對應碼': 'tx_code', '平台訂單編號': 'ref_no', '訂單日期': 'order_date',
    '應收': 'total', '手續費': 'fee_total', '交易處理費': 'tx_fee', '應入帳': 'payable',
    '實際入帳': 'actual_in', '入帳日': 'in_date',
    '狀態': 'recon_status', '手續費發票號碼': 'fee_invoice_no', '交易處理費發票號碼': 'tx_fee_invoice_no',
  }

  function handleSort(col) {
    if (sortCol === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc')
    else { setSortCol(col); setSortDir('asc') }
  }

  function getVal(o, key) {
    const v = o[key]
    return v == null ? '' : v
  }

  const shownOrders = orders
    .filter(o => {
      if (filterStatus && o.recon_status !== filterStatus) return false
      if (filterMonth && (o.order_date || '').slice(0, 7) !== filterMonth) return false
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
  // 訂單發票分組
  // 訂單（代開）發票分組：一張月結銷項發票群組多筆訂單。
  // 應開金額 = 進帳報表的「I欄商品原價 − M欄賣家負擔優惠券」，
  // 而該值恰等於「應入帳 + 手續費」（Y欄 + 費用欄加總），故直接由現有欄位回推，
  // 不需另存欄位，既有訂單也能立即套用。
  // 例（2026-07 該份報表 5 筆）：應開 14,320 → 內扣手續費後入帳 12,584。
  const ordInvBase = o => (o.payable || 0) + (o.fee_total || 0)
  const ordInvGroups = {}
  shownOrders.forEach(o => {
    if (!o.order_invoice_no) return
    if (!ordInvGroups[o.order_invoice_no]) ordInvGroups[o.order_invoice_no] = {
      invDate: o.order_invoice_date ?? null,
      invAmount: o.order_invoice_amount ?? null,   // 該組的代開發票金額（每筆重複存同一值）
      baseSum: 0, payableSum: 0, count: 0,
    }
    ordInvGroups[o.order_invoice_no].baseSum += ordInvBase(o)
    ordInvGroups[o.order_invoice_no].payableSum += o.payable || 0
    ordInvGroups[o.order_invoice_no].count++
  })

  const txInvoiceGroups = {}
  shownOrders.forEach(o => {
    if (!o.tx_fee_invoice_no) return
    if (!txInvoiceGroups[o.tx_fee_invoice_no]) txInvoiceGroups[o.tx_fee_invoice_no] = {
      txFeeSum: 0, count: 0,
      invAmount: parseFloat(localStorage.getItem(`txinv_amount_${o.tx_fee_invoice_no}`)) || null,
    }
    txInvoiceGroups[o.tx_fee_invoice_no].txFeeSum += o.tx_fee || 0
    txInvoiceGroups[o.tx_fee_invoice_no].count++
  })

  const invColorIdx = {}
  let _ci = 0
  Object.keys(invoiceGroups).forEach(k => { invColorIdx[k] = _ci++ % 2 })

  const ordInvColorIdx = {}
  let _oci = 0
  Object.keys(ordInvGroups).forEach(k => { ordInvColorIdx[k] = _oci++ % 2 })

  // 入帳群組：依 in_date 彙總已入帳訂單（僅官網>LINE Pay / 官網>信用卡）
  const showBankGroup = isLinePayOfficial || isPayuniCC
  const bankGroups = {}
  if (showBankGroup) shownOrders.forEach(o => {
    if (o.recon_status !== '已入帳' || !o.in_date) return
    const k = o.in_date.slice(0, 10)
    if (!bankGroups[k]) bankGroups[k] = { bankDeposit: o.bank_deposit ?? null, payableSum: 0, count: 0 }
    bankGroups[k].payableSum += o.payable || 0
    bankGroups[k].count++
  })

  const invFeeSum = invPreview?.feeSum ?? (invMethod === 'manual' ? manualFeeSum : null)
  const invAmountNum = parseFloat(invAmount) || 0
  const invDiff = invAmountNum > 0 && invFeeSum != null ? Math.round((invAmountNum - invFeeSum) * 100) / 100 : null
  const invIsMatch = invDiff != null && Math.abs(invDiff) < 0.01
  const hasInvOrders = invPreview?.orders?.length > 0 || (invMethod === 'manual' && checkedIds.size > 0)

  // ── 新增 UI 狀態（版面用，不影響既有邏輯）──
  const [activeStepKey, setActiveStepKey] = useState(null)
  const [query, setQuery] = useState('')
  const [sopOpen, setSopOpen] = useState(false)

  // ── 進度總覽：依 recon_status 統計（跟隨月份篩選）──
  const heroOrders = filterMonth ? orders.filter(o => (o.order_date || '').slice(0, 7) === filterMonth) : orders
  const cShipped = heroOrders.filter(o => !o.recon_status || o.recon_status === '已出貨' || o.recon_status === '待出貨').length
  const cSettled = heroOrders.filter(o => o.recon_status === '平台已結算').length
  const cPaid    = heroOrders.filter(o => o.recon_status === '已入帳').length
  const cDone    = heroOrders.filter(o => o.recon_status === '已對帳').length
  const heroTotal = heroOrders.length
  const donePct = heroTotal ? Math.round(cDone / heroTotal * 100) : 0
  const segW = n => (heroTotal ? (n / heroTotal * 100) : 0)

  // ── 對帳步驟：步驟數依金流而定，完成狀態由訂單資料推導 ──
  // 玉山銀行對帳是每條金流的收尾步驟（含酷澎、兆豐）
  const hasBankStep = true
  const nOrders = orders.length
  const lack = fn => orders.filter(fn).length
  const mkStep = (key, title, missingCount, unit) => ({
    key, title,
    done: nOrders > 0 && missingCount === 0,
    sub: nOrders === 0 ? '尚無訂單' : missingCount === 0 ? '已完成' : `${missingCount} 筆待${unit}`,
  })

  // 步驟順序＝使用者實際作業流程（2026-08-06 依此重排，取代原本「比照線上版區塊順序」的排法）：
  //   ① 上傳該渠道的對帳單／撥款明細／手續費明細
  //   ② 上傳天心，勾稽訂單編號後回填銷貨單號與訂單發票號碼
  //   ③ 手續費發票核對
  //   ④ 上傳玉山對帳單，確認實際入帳日（收尾）
  // 渠道專屬步驟依性質靠邊：費用類（交易處理費）緊接撥款明細，發票類靠發票核對。
  // 要再調整順序前請先跟使用者確認，勿依設計稿或舊版線上順序自行重排。
  const bankStep = () => mkStep('bank', '玉山銀行對帳',
    lack(o => o.recon_status !== '已入帳' && o.recon_status !== '已對帳'), '入帳')

  const stepDefs = []
  if (isShopee) stepDefs.push({
    key: 'shopeeOrd', title: '訂單匯入',
    done: nOrders > 0, sub: nOrders > 0 ? `${nOrders} 筆已匯入` : '待匯入',
  })
  stepDefs.push(mkStep('recon', '上傳撥款明細', lack(o => o.fee_total == null), '回填'))
  // 官網 LINE Pay 原本有獨立的「交易處理費」步驟（上傳 PayUni 帳戶明細）。
  // D-2 已逐筆提供交易處理費（tx_fee），該步驟純屬重複輸入，故移除；
  // 玉山對帳需要的處理費金額改由訂單的 tx_fee 加總得出。
  stepDefs.push(mkStep('sa', '銷貨單號', lack(o => !o.sa_no), '回填'))
  if (isShopee) stepDefs.push(mkStep('ordInv', '訂單發票', lack(o => !o.order_invoice_no), '開立'))
  // 官網LINE Pay 有兩種發票（手續費／交易處理費），這步改叫「手續費發票」以免混淆；其餘金流沿用「發票核對」
  stepDefs.push(mkStep('feeInv', isLinePayOfficial ? '手續費發票' : '發票核對', lack(o => !o.fee_invoice_no), '歸戶'))
  if (isLinePayOfficial) stepDefs.push(mkStep('txFeeInv', '交易處理費發票', lack(o => !o.tx_fee_invoice_no), '歸戶'))
  stepDefs.push(bankStep())

  const firstTodoKey = (stepDefs.find(s => !s.done) || stepDefs[0] || {}).key
  const curStep = stepDefs.some(s => s.key === activeStepKey) ? activeStepKey : firstTodoKey
  const stepCols = stepDefs.length <= 4 ? stepDefs.length : Math.ceil(stepDefs.length / 2)

  // ── 每步的「資料來源」提示（沿用既有後台操作路徑文字）──
  const SOURCE_HINT = {
    recon: isPayuniCC ? '統一金流 → 交易動態 → 入帳表 → 選擇期間 → 全部 → 查詢'
      : isLineMallLinePay ? 'LINE PAY 撥款明細：清算/撥款 → 檢視詳細記錄 → 預計撥款日 → 選擇期間 → Excel → 下載報表'
      : isLanxin ? '藍新金流 → 交易查詢 → 預計撥款日 → 開始查詢 → 下載查詢結果'
      : isLinePayOfficial ? 'Line Pay 撥款明細：存款/撥款 → 預計撥款日 → 選擇期間 → Excel；Payuni：交易動態 → 交易表 → 電子錢包 → 查詢'
      : '',
    bank: '玉山網銀 → 帳戶明細 → 下載對帳單',
  }

  // ── 訂單表：搜尋 + 狀態 chips ──
  const q = query.trim().toUpperCase()
  const searched = q
    ? shownOrders.filter(o => String(o.ref_no || '').toUpperCase().includes(q) || String(o.sa_no || '').toUpperCase().includes(q))
    : shownOrders
  const statusCounts = {}
  orders.forEach(o => { const k = o.recon_status || '待出貨'; statusCounts[k] = (statusCounts[k] || 0) + 1 })
  const chipDefs = [{ label: '全部', value: '', count: orders.length },
    ...STATUSES.map(s => ({ label: s, value: s, count: statusCounts[s] || 0 }))]

  // 上傳 SOP 截圖：存 Storage 後把 <img> 插進 contentEditable
  const stepPanel = {}

  // ── 面板：蝦皮訂單匯入 ──
  stepPanel.shopeeOrd = !isShopee ? null : (
    <div>
      <div style={panelLead}>上傳蝦皮「我的銷售」Excel，自動讀取訂單編號、訂單成立日期、買家總支付金額，重複訂單自動略過。</div>
      <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
        <input ref={shopeeOrdFileRef} type="file" accept=".xlsx,.xls,.csv"
          onChange={e => {
            const f = e.target.files?.[0]; if (!f) return
            setShopeeOrdFileName(f.name)
            const rd = new FileReader()
            rd.onload = ev => {
              const wb = XLSX.read(ev.target.result, { type: 'array' })
              const ws = wb.Sheets[wb.SheetNames[0]]
              const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
              // 跳過第一列（標題），保留 A欄有值的列
              setShopeeOrdRows(all.slice(1).filter(r => String(r[0] || '').trim()))
            }
            rd.readAsArrayBuffer(f)
          }} style={{ display: 'none' }} />
        <DropButton onClick={() => shopeeOrdFileRef.current.click()}
          filled={!!shopeeOrdRows} label={shopeeOrdFileName || '選擇蝦皮訂單 Excel'}
          hint={shopeeOrdRows ? `✓ ${shopeeOrdRows.length} 列` : '尚未選擇檔案'} />
        <button onClick={handleShopeeOrdImport} style={btnPri}>匯入</button>
      </div>
      <PanelMsg text={shopeeOrdMsg} bad={/錯誤|找不到/} />
    </div>
  )

  // ── 面板：銷貨單號（天心）──
  // 上傳欄位已移到頁籤上方的全域區（跨通路一次回填），這裡只留進度指示
  stepPanel.sa = (() => {
    const missing = orders.filter(o => !o.sa_no)
    return (
      <div>
        <div style={panelLead}>
          天心銷貨單是<strong>跨通路共用</strong>的，上傳欄位在本頁最上方的「天心銷貨單」區塊，一次上傳就會回填所有通路的銷貨單號與訂單發票號碼。
        </div>
        <div style={{ marginTop: 12, padding: '14px 18px', borderRadius: T.rInner,
          background: missing.length ? C.warnBg : T.g100 }}>
          {nOrders === 0 ? (
            <span style={{ fontSize: 13, color: T.n600 }}>本通路尚無訂單。</span>
          ) : missing.length === 0 ? (
            <span style={{ fontSize: 13, color: T.g700 }}>✓ 本通路 {nOrders} 筆訂單的銷貨單號都已回填。</span>
          ) : (
            <>
              <span style={{ fontSize: 13, color: C.warn }}>
                本通路尚有 <strong>{missing.length}</strong> 筆訂單沒有銷貨單號（共 {nOrders} 筆）。
              </span>
              <p style={{ margin: '6px 0 0', fontSize: 12, color: C.warn }}>
                待回填：{missing.slice(0, 8).map(o => o.ref_no).join('、')}
                {missing.length > 8 && `…等 ${missing.length} 筆`}
              </p>
            </>
          )}
        </div>
      </div>
    )
  })()

  // ── 面板：上傳撥款明細 ──
  stepPanel.recon = (
    <div>
      <div style={panelLead}>
        {isShopee ? '上傳蝦皮「我的進帳」，系統比對訂單編號並回填應收、手續費與應入帳。'
          : isPayuniCC ? '上傳 PayUni 入帳表，以去槓號的商店訂單編號比對，回填手續費與入帳金額。'
          : isMegabank ? '上傳兆豐「訂單明細報表」與「手續費報表」，系統以訂單編號勾稽兩表。應收取「商品金額＋運費」（非付款總金額，全額福利金折抵的訂單付款總金額為 0），應入帳＝應收－手續費。'
          : isTwoFile ? '上傳兩份撥款報表，系統以支付對應碼勾稽兩表，回填手續費與應入帳金額。'
          : '上傳撥款明細，系統自動比對平台訂單、回填手續費與應入帳金額。'}
      </div>
      {SOURCE_HINT.recon && <SourceHint text={SOURCE_HINT.recon} onSop={() => setSopOpen(true)} />}
      <div style={{ display: 'grid', gridTemplateColumns: isTwoFile ? '1fr 1fr' : '1fr', gap: 14, marginTop: 14 }}>
        {isTwoFile ? (
          <>
            <input ref={fileRef1} type="file" accept=".xlsx,.xls,.csv" onChange={e => readFile(e, setRows1, setFileName1)} style={{ display: 'none' }} />
            <DropButton onClick={() => fileRef1.current.click()} filled={!!rows1}
              label={fileName1 || (isMegabank ? '訂單明細報表' : 'Line Pay撥款明細')} hint={rows1 ? `✓ ${rows1.length} 列` : '尚未選擇檔案'} block />
            <input ref={fileRef2} type="file" accept=".xlsx,.xls,.csv" onChange={e => readFile(e, setRows2, setFileName2)} style={{ display: 'none' }} />
            <DropButton onClick={() => fileRef2.current.click()} filled={!!rows2}
              label={fileName2 || (isMegabank ? '手續費報表' : 'Payuni電子錢包')} hint={rows2 ? `✓ ${rows2.length} 列` : '尚未選擇檔案'} block />
          </>
        ) : (
          <>
            <input ref={fileRef1} type="file" accept=".xlsx,.xls,.csv" onChange={e => readFile(e, setRows1, setFileName1)} style={{ display: 'none' }} />
            <DropButton onClick={() => fileRef1.current.click()} filled={!!rows1}
              label={fileName1 || (isShopee ? '選擇蝦皮「我的進帳」' : isPayuniCC ? '選擇入帳表' : '選擇對帳單')}
              hint={rows1 ? `✓ ${rows1.length} 列` : '尚未選擇檔案'} block />
          </>
        )}
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginTop: 16, padding: '14px 18px',
        background: reconResult ? T.g100 : T.n200, borderRadius: T.rInner, flexWrap: 'wrap' }}>
        {reconResult && (reconResult.updated > 0 || reconResult.inserted > 0) ? (
          <>
            <span style={{ fontSize: 13 }}>手續費合計 <strong style={{ fontSize: 16 }}>NT$ {reconResult.feeTotal.toLocaleString()}</strong>
              <span style={{ fontSize: 11, color: T.n600, marginLeft: 4 }}>（供發票核對）</span></span>
            <span style={{ fontSize: 13 }}>預計撥款 <strong style={{ fontSize: 16 }}>NT$ {reconResult.payableTotal.toLocaleString()}</strong>
              <span style={{ fontSize: 11, color: T.n600, marginLeft: 4 }}>（供玉山對帳單核對）</span></span>
            <span style={{ fontSize: 12, color: T.g700 }}>
              ✓ 已回填 {reconResult.updated} 筆{createMissing ? ` · 新增 ${reconResult.inserted} 筆` : ` · 未對應 ${reconResult.unmatched.length} 筆`}
            </span>
          </>
        ) : (
          <span style={{ fontSize: 13, color: T.n600 }}>選好檔案後按「比對回填」，系統會寫回手續費與應入帳。</span>
        )}
        <button onClick={handleReconcile} style={{ ...btnPri, marginLeft: 'auto' }}>
          {reconResult ? '重新比對回填' : '比對回填'}
        </button>
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 10, fontSize: 13, color: T.n700, cursor: 'pointer' }}>
        <input type="checkbox" checked={createMissing} onChange={e => setCreateMissing(e.target.checked)} />
        未匯入的訂單自動建檔（比對不到時，直接以對帳單資料新增；已存在則略過）
      </label>
      <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 6, fontSize: 13, color: T.n600, cursor: 'pointer' }}>
        <input type="checkbox" checked={skipColCheck} onChange={e => setSkipColCheck(e.target.checked)} />
        略過格式檢查（僅在確定報表正確、只是欄位名稱不同時勾選）
      </label>
      <PanelMsg text={reconMsg} bad={/錯誤|請/} />
      {reconSkipped && (
        <div style={{ marginTop: 8, padding: '10px 14px', background: T.n200, borderRadius: T.rInner, fontSize: 12, color: T.n700 }}>
          <strong>已略過 {reconSkipped.length} 筆未成立的交易</strong>（不寫入資料庫）：
          {reconSkipped.slice(0, 8).map(s => `${s.key}（${s.status}）`).join('、')}
          {reconSkipped.length > 8 && `…等 ${reconSkipped.length} 筆`}
        </div>
      )}
      {reconWarn && (
        <div style={{ marginTop: 8, padding: '10px 14px', background: C.warnBg, borderRadius: T.rInner, fontSize: 12, color: C.warn }}>
          <strong>費率檢核未通過（兆豐應固定 6%）</strong>
          {reconWarn.noFee.length > 0 && (
            <p style={{ margin: '6px 0 0' }}>
              查無手續費（手續費報表可能缺列或期間未涵蓋）：{reconWarn.noFee.slice(0, 8).join('、')}
              {reconWarn.noFee.length > 8 && `…等 ${reconWarn.noFee.length} 筆`}
            </p>
          )}
          {reconWarn.oddRate.length > 0 && (
            <p style={{ margin: '6px 0 0' }}>
              費率異常：{reconWarn.oddRate.slice(0, 6).map(o => `${o.key}（${o.fee}/${o.total}＝${o.rate}%）`).join('、')}
              {reconWarn.oddRate.length > 6 && `…等 ${reconWarn.oddRate.length} 筆`}
            </p>
          )}
        </div>
      )}
      {createMissing && reconResult?.insertedKeys?.length > 0 && (
        <p style={{ marginTop: 6, marginBottom: 0, fontSize: 12, color: T.g700 }}>
          新增訂單：{reconResult.insertedKeys.slice(0, 5).join('、')}
          {reconResult.insertedKeys.length > 5 && `…等 ${reconResult.insertedKeys.length} 筆`}
        </p>
      )}
      {!createMissing && reconResult?.unmatched?.length > 0 && (
        <p style={{ marginTop: 6, marginBottom: 0, fontSize: 12, color: C.warn }}>
          未對應：{reconResult.unmatched.slice(0, 5).join('、')}
          {reconResult.unmatched.length > 5 && `…等 ${reconResult.unmatched.length} 筆`}
        </p>
      )}
    </div>
  )

  // ── 面板：交易處理費（官網 LINE Pay）──
  // ── 面板：玉山銀行對帳（非蝦皮）──
  const bankPanelGeneral = (isShopee || !hasBankStep) ? null : (
    <div>
      <div style={panelLead}>
        {isPayuniCC || isCoupang || isMegabank
          ? '上傳玉山對帳單，勾選對應訂單後確認入帳日。'
          : '上傳玉山對帳單，系統以撥款報表「預計撥款日 + 金額」為錨點比對銀行入帳，自動排除同帳號的其他平台撥款。'}
        {!isPayuniCC && !isCoupang && !isMegabank && <span style={{ color: T.n600 }}>
          （{isLanxin ? '信用卡 008/...35101' : isLinePayOfficial ? 'LINE Pay 808/...24585' : 'LINE Pay 387/...60558379'}）
        </span>}
      </div>
      <SourceHint text={SOURCE_HINT.bank} onSop={() => setSopOpen(true)} />
      <div style={{ marginTop: 14 }}>
        <input type="file" ref={bankFileRef} style={{ display: 'none' }} accept=".xlsx,.xls,.csv" onChange={readBankFile} />
        <DropButton onClick={() => bankFileRef.current.click()} filled={bankRows.length > 0}
          label={bankFileName || '上傳玉山對帳單'}
          hint={bankRows.length > 0 ? `✓ ${bankRows.length} 筆入帳` : '尚未選擇檔案'} />
      </div>

      {/* 已確認入帳群組 — 從 DB 訂單計算，不依賴銀行對帳單是否上傳 */}
      {(() => {
        const confirmed = orders.filter(o => o.recon_status === '已入帳' && o.in_date)
        if (confirmed.length === 0) return null
        const groups = {}
        confirmed.forEach(o => {
          const k = (o.in_date || '').slice(0, 10)
          if (!groups[k]) groups[k] = { date: k, orders: [], payable: 0 }
          groups[k].orders.push(o)
          groups[k].payable += o.payable || 0
        })
        const sorted = Object.values(groups).sort((a, b) => a.date.localeCompare(b.date))
        return <ConfirmedGroups groups={sorted} exp={confirmedGroupExp} setExp={setConfirmedGroupExp} onRelease={releaseFromGroup} releasingKey={releasingKey} />
      })()}

      {bankRows.length > 0 && (() => {
        // 從已上傳的撥款報表（rows1）建立「銀行入帳日 → 預計金額」對照表
        // 蘭新信用卡撥款日隔天才到玉山帳戶，所以用 payoutDate+1 當 bankDate
        const linepayByDate = {}
        // 酷澎／兆豐的報表沒有預計撥款日，不建對照表（否則會拿 linepay parser 去解讀不相干的檔案）
        const payoutRows = (isCoupang || isMegabank) ? []
          : isLinePayOfficial
          ? (rows1 && rows2 ? parseOfficialLinePayReconDual(rows1, rows2) : [])
          : isLanxin
            ? (rows1 ? RECON_PARSERS.lanxin(rows1) : [])
            : (rows1 ? RECON_PARSERS.linepay(rows1) : [])
        payoutRows.forEach(r => {
          const payoutDate = (r.in_date || '').slice(0, 10)
          if (!payoutDate) return
          let bankDate = payoutDate
          if (isLanxin) {
            const d = new Date(payoutDate + 'T00:00:00Z')
            d.setUTCDate(d.getUTCDate() + 1)
            bankDate = d.toISOString().slice(0, 10)
          }
          if (!linepayByDate[bankDate]) linepayByDate[bankDate] = { expected: 0, count: 0, payoutDate }
          linepayByDate[bankDate].expected += r.payable || 0
          linepayByDate[bankDate].count++
        })
        const hasPayoutMap = Object.keys(linepayByDate).length > 0

        // payuniCC / LINE Pay / 蘭新: 手動選取訂單，直接顯示全部銀行入帳
        // isLinePayOfficial: 已篩帳號，直接顯示全部
        // 其他金流（若未來新增）: 以撥款報表日期過濾
        const displayRows = (isManualSelection || isLinePayOfficial)
          ? bankRows
          : hasPayoutMap
            ? bankRows.filter(br => linepayByDate[br.date] !== undefined)
            : bankRows

        async function batchConfirm() {
          const selected = [...bankEntryChecked]
          if (!selected.length) return
          setBankMsg(p => { const n = {...p}; selected.forEach(i => { n[i] = '寫入中…' }); return n })
          let hasErr = false
          for (const idx of selected) {
            const br = displayRows[idx]
            if (!br) continue
            const dateOrders = isLinePayOfficial
              ? orders.filter(o => o.recon_status !== '已入帳' || !o.in_date)
              : (() => {
                  const payoutInfo = linepayByDate[br.date]
                  const matchDate = payoutInfo?.payoutDate || br.date
                  return orders.filter(o => (o.in_date || '').slice(0, 10) === matchDate)
                })()
            for (const o of dateOrders) {
              const { error } = await supabase.from('shipping_orders')
                .update({ in_date: br.date, actual_in: br.deposit, bank_deposit: br.deposit, recon_status: '已入帳' }).eq('id', o.id)
              if (error) { hasErr = true; break }
            }
            if (!hasErr) {
              setBankMsg(p => ({ ...p, [idx]: '✓ 已回填' }))
            } else break
          }
          if (!hasErr) { setBankEntryChecked(new Set()); await loadOrders() }
        }

        const allChecked = displayRows.length > 0 && bankEntryChecked.size === displayRows.length

        return (
          <div style={{ marginTop: 18 }}>
            {!isLinePayOfficial && !isManualSelection && !hasPayoutMap && (
              <p style={{ fontSize: 12, color: C.warn, margin: '0 0 10px' }}>
                ⚠ 尚未上傳撥款報表，無法自動篩選。請先回到「上傳撥款明細」步驟上傳報表。
              </p>
            )}
            {!isLinePayOfficial && !isManualSelection && hasPayoutMap && displayRows.length === 0 && (
              <p style={{ fontSize: 12, color: T.danger, margin: '0 0 10px' }}>
                ⚠ 撥款日期未對到銀行入帳日。撥款報表日期：{Object.keys(linepayByDate).slice(0, 5).join('、')}；銀行對帳單日期：{bankRows.slice(0, 5).map(r => r.date).join('、')}
              </p>
            )}
            {!isManualSelection && (
              <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
                <label style={{ fontSize: 13, display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                  <input type="checkbox" checked={allChecked}
                    onChange={() => setBankEntryChecked(allChecked ? new Set() : new Set(displayRows.map((_, i) => i)))} />
                  全選
                </label>
                {bankEntryChecked.size > 0 && (
                  <button onClick={batchConfirm} style={{ ...btnPri, fontSize: 13 }}>
                    批次確認入帳（{bankEntryChecked.size} 筆）
                  </button>
                )}
              </div>
            )}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              {displayRows.map((br, idx) => {
                const payoutInfo = isLinePayOfficial ? null : linepayByDate[br.date]
                const expected = payoutInfo ? Math.round(payoutInfo.expected * 100) / 100 : null
                const diff = expected != null ? Math.round((br.deposit - expected) * 100) / 100 : null
                const isMatch = diff != null && Math.abs(diff) <= 2
                const expanded = !!bankExpanded[idx]
                const matchDate = payoutInfo?.payoutDate || br.date
                const dateOrders = ((isLinePayOfficial || isManualSelection)
                  ? orders.filter(o => o.recon_status !== '已入帳' || !o.in_date)
                  : orders.filter(o => (o.in_date || '').slice(0, 10) === matchDate)
                ).sort((a, b) => (a.order_date || '').localeCompare(b.order_date || ''))
                const ccSel = isManualSelection ? (bankCCOrderSel[idx] ?? new Set()) : new Set()
                const selectedOrders = isManualSelection ? dateOrders.filter(o => ccSel.has(String(o.id))) : dateOrders
                const ordersPayable = Math.round(selectedOrders.reduce((s, o) => s + (o.payable || 0), 0) * 100) / 100
                const ccDiff = isManualSelection && ccSel.size > 0 ? Math.round((br.deposit - ordersPayable) * 100) / 100 : null
                const ccMatch = ccDiff != null && Math.abs(ccDiff) <= 1
                // 交易處理費改由訂單的 tx_fee（D-2 逐筆帶入）加總，不再依賴另外上傳的帳戶明細
                const txFeeOrders = selectedOrders.filter(o => o.tx_fee)
                const txFeeTotal = Math.round(txFeeOrders.reduce((s, o) => s + (o.tx_fee || 0), 0) * 100) / 100
                const entryChecked = bankEntryChecked.has(idx)
                const isDone = !!bankMsg[idx]?.startsWith('✓')
                const cardBorderColor = isManualSelection
                  ? (ccSel.size > 0 ? (ccMatch ? T.g500 : T.danger) : T.divider)
                  : (entryChecked ? T.a : isMatch ? T.g500 : diff != null ? T.danger : T.divider)
                const cardBg = isManualSelection
                  ? (ccSel.size > 0 && ccMatch ? T.g100 : T.bg)
                  : (entryChecked ? T.a100 : isMatch ? T.g100 : T.bg)
                return (
                  <div key={idx} style={{ border: `1.5px solid ${cardBorderColor}`, borderRadius: T.rPanel, padding: '16px 18px', background: cardBg }}>
                    <div style={{ display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
                      {!isPayuniCC && !isDone && (
                        <input type="checkbox" checked={entryChecked}
                          onChange={() => setBankEntryChecked(p => { const n = new Set(p); n.has(idx) ? n.delete(idx) : n.add(idx); return n })} />
                      )}
                      <span style={{ fontWeight: 700, fontSize: 16, minWidth: 96 }}>{br.date}</span>
                      {!isDone && (isLinePayOfficial || isPayuniCC) && br.summary && (
                        <span style={{ fontSize: 12, color: T.n600, fontFamily: 'monospace' }}>{br.summary}</span>
                      )}
                      {!isDone && expected != null && (
                        <span style={{ fontSize: 13, color: T.n700 }}>
                          {isLanxin ? '蘭新預計撥款：' : 'LINE Pay 預計：'}<strong>NT$ {expected.toLocaleString()}</strong>（{payoutInfo.count} 筆）{payoutInfo.payoutDate && isLanxin && <span style={{ fontSize: 11, marginLeft: 4 }}>撥款日 {payoutInfo.payoutDate}</span>}
                        </span>
                      )}
                      <span style={{ fontSize: 14 }}>
                        銀行入帳 <strong>NT$ {br.deposit.toLocaleString()}</strong>
                      </span>
                      {!isDone && isManualSelection && ccSel.size > 0 && (
                        <span style={{ fontSize: 14 }}>
                          已選 <strong>NT$ {ordersPayable.toLocaleString()}</strong>
                          {ccDiff != null && (
                            <span style={{ marginLeft: 8, fontWeight: 700, color: ccMatch ? T.g700 : T.danger }}>
                              差異 {ccDiff > 0 ? '+' : ''}{ccDiff}{ccMatch && ' ✓ 相符'}
                            </span>
                          )}
                        </span>
                      )}
                      {!isDone && isManualSelection && ccSel.size === 0 && (
                        <span style={{ fontSize: 13, color: T.n600 }}>尚未選取訂單</span>
                      )}
                      {!isDone && isLinePayOfficial && txFeeTotal > 0 && (
                        <span style={{ fontSize: 13, color: T.danger }}>
                          交易處理費 <strong>-NT$ {txFeeTotal.toLocaleString()}</strong>（{txFeeOrders.length} 筆）
                        </span>
                      )}
                      {!isDone && !isManualSelection && <span style={{ fontSize: 11, color: T.n500, fontFamily: 'monospace' }}>{br.account}</span>}
                      {!isDone && diff != null && (
                        <span style={{ fontSize: 14, fontWeight: 700, color: isMatch ? T.g700 : T.danger }}>
                          差異 {diff > 0 ? '+' : ''}{diff}{isMatch && ' ✓ 相符'}
                        </span>
                      )}
                      {!isDone && dateOrders.length > 0 && (
                        <button
                          onClick={() => setBankExpanded(p => ({ ...p, [idx]: !expanded }))}
                          style={{ ...btnSec, fontSize: 13, padding: '5px 12px', marginLeft: 'auto' }}
                        >{expanded ? '收起 ▲' : `選取訂單 ▾ (${dateOrders.length})`}</button>
                      )}
                      {!isDone && (isManualSelection ? ccSel.size > 0 : dateOrders.length > 0) && (
                        <button
                          onClick={() => confirmBankEntry(idx, br, dateOrders)}
                          style={{ ...btnPri, fontSize: 13, padding: '5px 12px' }}
                        >{isManualSelection ? `確認入帳（${ccSel.size} 筆）` : '確認入帳'}</button>
                      )}
                      {bankMsg[idx] && (
                        <span style={{ fontSize: 13, marginLeft: 'auto', fontWeight: 600, color: bankMsg[idx].includes('❌') ? T.danger : T.g700 }}>
                          {bankMsg[idx]}
                        </span>
                      )}
                    </div>

                    {!isDone && expanded && dateOrders.length > 0 && (
                      <div style={{ marginTop: 12, borderTop: `1px solid ${T.divider}`, paddingTop: 12 }}>
                        <div style={{ maxHeight: 320, overflowY: 'auto' }}>
                        <table style={{ width: '100%', fontSize: 12, borderCollapse: 'collapse' }}>
                          <thead>
                            <tr style={{ color: T.n600 }}>
                              {isManualSelection && <th style={{ padding: '4px 6px', width: 24 }}></th>}
                              <th style={subTh}>平台訂單編號</th>
                              <th style={subTh}>訂單日期</th>
                              {isLinePayOfficial && <th style={subTh}>應撥款日</th>}
                              {isLinePayOfficial && <th style={subTh}>實際撥款日</th>}
                              {isManualSelection && !isLinePayOfficial && <th style={subTh}>付款方式</th>}
                              <th style={{ ...subTh, textAlign: 'right' }}>應入帳</th>
                            </tr>
                          </thead>
                          <tbody>
                            {dateOrders.map(o => {
                              const oChecked = isManualSelection && ccSel.has(String(o.id))
                              return (
                              <tr key={o.id} style={{ borderBottom: `1px solid ${T.divider}`, background: oChecked ? T.a100 : 'transparent' }}>
                                {isManualSelection && (
                                  <td style={{ padding: '4px 6px' }}>
                                    <input type="checkbox" checked={oChecked} onChange={() => {
                                      setBankCCOrderSel(p => {
                                        const prev = p[idx] ? new Set(p[idx]) : new Set()
                                        prev.has(String(o.id)) ? prev.delete(String(o.id)) : prev.add(String(o.id))
                                        return { ...p, [idx]: prev }
                                      })
                                    }} />
                                  </td>
                                )}
                                <td style={{ padding: '4px 6px', fontFamily: 'monospace' }}>{o.ref_no}</td>
                                <td style={{ padding: '4px 6px' }}>{o.order_date}</td>
                                {isLinePayOfficial && <td style={{ padding: '4px 6px' }}>{o.in_date || '—'}</td>}
                                {isLinePayOfficial && <td style={{ padding: '4px 6px', color: br.actualDate !== (o.in_date || '').slice(0, 10) ? C.warn : T.text }}>{br.actualDate || '—'}</td>}
                                {isManualSelection && !isLinePayOfficial && <td style={{ padding: '4px 6px', fontSize: 11, color: T.n600 }}>{o.pay_method || '—'}</td>}
                                <td style={{ padding: '4px 6px', textAlign: 'right' }}>{o.payable?.toLocaleString()}</td>
                              </tr>
                              )
                            })}
                            <tr style={{ borderTop: `1px solid ${T.n400}`, fontWeight: 700 }}>
                              <td colSpan={isLinePayOfficial ? 5 : 4} style={{ padding: '6px', color: T.n600, fontSize: 11 }}>
                                {isManualSelection ? `已選 ${ccSel.size} 筆合計` : '訂單合計'}
                              </td>
                              <td style={{ padding: '6px', textAlign: 'right' }}>{ordersPayable.toLocaleString()}</td>
                            </tr>
                          </tbody>
                        </table>
                        </div>
                        {isLinePayOfficial && txFeeOrders.length > 0 && (
                          <table style={{ width: '100%', fontSize: 12, borderCollapse: 'collapse', marginTop: 10, borderTop: `1px solid ${T.divider}` }}>
                            <thead>
                              <tr style={{ color: T.n600 }}>
                                <th style={subTh}>訂單編號</th>
                                <th style={subTh}>訂單日期</th>
                                <th style={{ ...subTh, textAlign: 'right', color: T.danger }}>交易處理費（負）</th>
                              </tr>
                            </thead>
                            <tbody>
                              {txFeeOrders.map(o => (
                                <tr key={o.id} style={{ borderBottom: `1px solid ${T.divider}` }}>
                                  <td style={{ padding: '4px 6px', fontFamily: 'monospace' }}>{o.ref_no}</td>
                                  <td style={{ padding: '4px 6px' }}>{o.order_date || '—'}</td>
                                  <td style={{ padding: '4px 6px', textAlign: 'right', color: T.danger }}>-{o.tx_fee.toLocaleString()}</td>
                                </tr>
                              ))}
                              <tr style={{ fontWeight: 700 }}>
                                <td colSpan={2} style={{ padding: '4px 6px', color: T.n600 }}>合計</td>
                                <td style={{ padding: '4px 6px', textAlign: 'right', color: T.danger }}>-{txFeeTotal.toLocaleString()}</td>
                              </tr>
                            </tbody>
                          </table>
                        )}
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          </div>
        )
      })()}
    </div>
  )

  // ── 面板：玉山銀行對帳（蝦皮）──
  const bankPanelShopee = !isShopee ? null : (
    <div>
      <div style={panelLead}>上傳玉山對帳單，篩出備註含「SHOPEE」或帳號 808/0370979139156 的入帳，手動勾選對應訂單後確認入帳日。</div>
      <SourceHint text={SOURCE_HINT.bank} onSop={() => setSopOpen(true)} />
      <div style={{ marginTop: 14 }}>
        <input type="file" ref={bankFileRef} style={{ display: 'none' }} accept=".xlsx,.xls,.csv" onChange={readBankFile} />
        <DropButton onClick={() => bankFileRef.current.click()} filled={bankRows.length > 0}
          label={bankFileName || '上傳玉山對帳單'}
          hint={bankRows.length > 0 ? `✓ ${bankRows.length} 筆入帳` : '尚未選擇檔案'} />
      </div>

      {/* 已確認入帳群組 */}
      {(() => {
        const confirmed = orders.filter(o => o.recon_status === '已入帳' && o.in_date)
        if (!confirmed.length) return null
        const groups = {}
        confirmed.forEach(o => {
          const k = (o.in_date || '').slice(0, 10)
          if (!groups[k]) groups[k] = { date: k, orders: [], payable: 0 }
          groups[k].orders.push(o)
          groups[k].payable += o.payable || 0
        })
        const sorted = Object.values(groups).sort((a, b) => a.date.localeCompare(b.date))
        return <ConfirmedGroups groups={sorted} exp={confirmedGroupExp} setExp={setConfirmedGroupExp} onRelease={releaseFromGroup} releasingKey={releasingKey} />
      })()}

      {/* 銀行對帳單入帳列表 */}
      {bankRows.length > 0 && (() => {
        const displayRows = bankRows

        async function batchConfirmShopee() {
          const selected = [...bankEntryChecked]
          if (!selected.length) return
          setBankMsg(p => { const n = {...p}; selected.forEach(i => { n[i] = '寫入中…' }); return n })
          let hasErr = false
          for (const idx of selected) {
            const br = displayRows[idx]
            if (!br) continue
            const ccSel = bankCCOrderSel[idx] ?? new Set()
            const dateOrders = orders.filter(o => ccSel.has(String(o.id)))
            for (const o of dateOrders) {
              const { error } = await supabase.from('shipping_orders')
                .update({ in_date: br.date, actual_in: br.deposit, bank_deposit: br.deposit, recon_status: '已入帳' }).eq('id', o.id)
              if (error) { hasErr = true; break }
            }
            if (!hasErr) setBankMsg(p => ({ ...p, [idx]: '✓ 已回填' }))
            else break
          }
          if (!hasErr) { setBankEntryChecked(new Set()); await loadOrders() }
        }

        const allChecked = displayRows.length > 0 && bankEntryChecked.size === displayRows.length

        return (
          <div style={{ marginTop: 18 }}>
            <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
              <label style={{ fontSize: 13, display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }}>
                <input type="checkbox" checked={allChecked}
                  onChange={() => setBankEntryChecked(allChecked ? new Set() : new Set(displayRows.map((_, i) => i)))} />
                全選
              </label>
              {bankEntryChecked.size > 0 && (
                <button onClick={batchConfirmShopee} style={{ ...btnPri, fontSize: 13 }}>
                  批次確認入帳（{bankEntryChecked.size} 筆）
                </button>
              )}
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              {displayRows.map((br, idx) => {
                const ccSel = bankCCOrderSel[idx] ?? new Set()
                const pendingOrders = orders
                  .filter(o => o.recon_status !== '已入帳' || !o.in_date)
                  .sort((a, b) => (a.order_date || '').localeCompare(b.order_date || ''))
                const selectedOrders = pendingOrders.filter(o => ccSel.has(String(o.id)))
                const ordersPayable = Math.round(selectedOrders.reduce((s, o) => s + (o.payable || 0), 0) * 100) / 100
                const ccDiff = ccSel.size > 0 ? Math.round((br.deposit - ordersPayable) * 100) / 100 : null
                const ccMatch = ccDiff != null && Math.abs(ccDiff) <= 1
                const expanded = !!bankExpanded[idx]
                const isDone = !!bankMsg[idx]?.startsWith('✓')
                const cardBorderColor = ccSel.size > 0 ? (ccMatch ? T.g500 : T.danger) : T.divider
                const cardBg = ccSel.size > 0 && ccMatch ? T.g100 : T.bg
                return (
                  <div key={idx} style={{ border: `1.5px solid ${cardBorderColor}`, borderRadius: T.rPanel, padding: '16px 18px', background: cardBg }}>
                    <div style={{ display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
                      {!isDone && (
                        <input type="checkbox" checked={bankEntryChecked.has(idx)}
                          onChange={() => setBankEntryChecked(p => { const n = new Set(p); n.has(idx) ? n.delete(idx) : n.add(idx); return n })} />
                      )}
                      <span style={{ fontWeight: 700, fontSize: 16, minWidth: 96 }}>{br.date}</span>
                      <span style={{ fontSize: 12, color: T.n600 }}>{br.summary}</span>
                      <span style={{ fontSize: 14 }}>銀行入帳 <strong>NT$ {br.deposit.toLocaleString()}</strong></span>
                      {ccSel.size > 0 ? (
                        <span style={{ fontSize: 14 }}>
                          已選 <strong>NT$ {ordersPayable.toLocaleString()}</strong>
                          {ccDiff != null && (
                            <span style={{ marginLeft: 8, fontWeight: 700, color: ccMatch ? T.g700 : T.danger }}>
                              差異 {ccDiff > 0 ? '+' : ''}{ccDiff}{ccMatch && ' ✓ 相符'}
                            </span>
                          )}
                        </span>
                      ) : !isDone && <span style={{ fontSize: 13, color: T.n600 }}>尚未選取訂單</span>}
                      {!isDone && (
                        <button onClick={() => setBankExpanded(p => ({ ...p, [idx]: !expanded }))}
                          style={{ ...btnSec, fontSize: 13, padding: '5px 12px', marginLeft: 'auto' }}>
                          {expanded ? '收起 ▲' : `選取訂單 ▾ (${pendingOrders.length})`}
                        </button>
                      )}
                      {!isDone && ccSel.size > 0 && (
                        <button onClick={async () => {
                          const dateOrders = pendingOrders.filter(o => ccSel.has(String(o.id)))
                          for (const o of dateOrders) {
                            await supabase.from('shipping_orders')
                              .update({ in_date: br.date, actual_in: br.deposit, bank_deposit: br.deposit, recon_status: '已入帳' }).eq('id', o.id)
                          }
                          setBankMsg(p => ({ ...p, [idx]: '✓ 已回填' }))
                          await loadOrders()
                        }} style={{ ...btnPri, fontSize: 13, padding: '5px 12px' }}>
                          確認入帳（{ccSel.size} 筆）
                        </button>
                      )}
                      {bankMsg[idx] && (
                        <span style={{ fontSize: 13, fontWeight: 700, color: bankMsg[idx].startsWith('✓') ? T.g700 : T.danger }}>
                          {bankMsg[idx]}
                        </span>
                      )}
                    </div>
                    {!isDone && expanded && (
                      <div style={{ marginTop: 12, borderTop: `1px solid ${T.divider}`, paddingTop: 12 }}>
                        <table style={{ width: '100%', fontSize: 12, borderCollapse: 'collapse' }}>
                          <thead><tr style={{ color: T.n600 }}>
                            <th style={{ padding: '3px 6px' }}></th>
                            <th style={subTh}>平台訂單編號</th>
                            <th style={subTh}>訂單日期</th>
                            <th style={{ ...subTh, textAlign: 'right' }}>應入帳</th>
                          </tr></thead>
                          <tbody>
                            {pendingOrders.map(o => (
                              <tr key={o.id} style={{ background: ccSel.has(String(o.id)) ? T.a100 : 'transparent', borderBottom: `1px solid ${T.divider}` }}>
                                <td style={{ padding: '3px 6px' }}>
                                  <input type="checkbox" checked={ccSel.has(String(o.id))}
                                    onChange={() => {
                                      const s = new Set(ccSel)
                                      s.has(String(o.id)) ? s.delete(String(o.id)) : s.add(String(o.id))
                                      setBankCCOrderSel(p => ({ ...p, [idx]: s }))
                                    }} />
                                </td>
                                <td style={{ padding: '3px 6px', fontFamily: 'monospace' }}>{o.ref_no}</td>
                                <td style={{ padding: '3px 6px' }}>{o.order_date}</td>
                                <td style={{ padding: '3px 6px', textAlign: 'right' }}>{o.payable?.toLocaleString() ?? '—'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          </div>
        )
      })()}
    </div>
  )

  stepPanel.bank = isShopee ? bankPanelShopee : bankPanelGeneral

  // ── 面板：訂單發票輸入（蝦皮專屬）──
  stepPanel.ordInv = !isShopee ? null : (() => {
    const amtNum = parseFloat(ordInvEntryAmount) || 0
    const previewOrders = ordInvEntryMethod === 'auto'
      ? (ordInvEntryPreview || [])
      : orders.filter(o => !o.order_invoice_no).sort((a, b) => (a.order_date || '').localeCompare(b.order_date || ''))
    const selectedOrders = ordInvEntryMethod === 'auto'
      ? (ordInvEntryPreview || [])
      : orders.filter(o => ordInvEntryChecked.has(o.id))
    // 代開發票金額＝該組訂單的應收合計（商品原價）；蝦皮再內扣手續費與優惠券後撥款
    const totalSum = Math.round(selectedOrders.reduce((s, o) => s + ordInvBase(o), 0) * 100) / 100
    const payableSum = Math.round(selectedOrders.reduce((s, o) => s + (o.payable || 0), 0) * 100) / 100
    const diff = Math.round((totalSum - amtNum) * 100) / 100
    const isMatch = amtNum > 0 && diff === 0
    const hasSel = selectedOrders.length > 0
    return (
      <div>
        <div style={panelLead}>
          依發票號碼分組核對：一個月開立一張代開發票，金額須等於該組訂單的<strong>應開發票金額</strong>
          合計（進帳報表的 I欄商品原價 − M欄賣家負擔優惠券，等同「應入帳 ＋ 手續費」）；
          蝦皮再從中內扣手續費後撥款，剩下的才是應入帳。左側色帶與下方明細表的訂單相互對應。
        </div>
        <InvoiceGroupCards
          groups={ordInvGroups} colorIdx={ordInvColorIdx} kind="ord" storeKey={gateway}
          orders={shownOrders} onOpen={k => { setViewOrdInvKey(k); setOrdInvDeleteConfirm(false); setOrdInvPopupDate(ordInvGroups[k]?.invDate || '') }} />
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
          <Field label="發票號碼">
            <input value={ordInvEntryNo} onChange={e => setOrdInvEntryNo(e.target.value)} placeholder="AB-12345678" style={inpT} />
          </Field>
          <Field label="發票日期">
            <input type="date" value={ordInvEntryDate} onChange={e => setOrdInvEntryDate(e.target.value)} style={inpT} />
          </Field>
          <Field label="發票金額（含稅）">
            <input type="number" value={ordInvEntryAmount} onChange={e => setOrdInvEntryAmount(e.target.value)} placeholder="0" style={inpT} />
          </Field>
        </div>

        <SegBtns value={ordInvEntryMethod}
          onChange={v => { setOrdInvEntryMethod(v); setOrdInvEntryPreview(null); setOrdInvEntryChecked(new Set()); setOrdInvEntryMsg('') }} />

        {ordInvEntryMethod === 'auto' && (
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end', flexWrap: 'wrap' }}>
            <Field label="訂單日期起"><input type="date" value={ordInvEntryFrom} onChange={e => setOrdInvEntryFrom(e.target.value)} style={inpT} /></Field>
            <Field label="訂單日期訖"><input type="date" value={ordInvEntryTo} onChange={e => setOrdInvEntryTo(e.target.value)} style={inpT} /></Field>
            <div style={{ paddingBottom: 10 }}><button onClick={runOrdInvEntryPreview} style={btnPri}>查詢</button></div>
          </div>
        )}

        {ordInvEntryMethod === 'manual' && (
          <div style={pickWrap}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
              <thead>
                <tr>
                  <th style={{ ...thT, position: 'static' }}>
                    <input type="checkbox"
                      checked={previewOrders.length > 0 && ordInvEntryChecked.size === previewOrders.length}
                      onChange={() => {
                        if (ordInvEntryChecked.size === previewOrders.length) setOrdInvEntryChecked(new Set())
                        else setOrdInvEntryChecked(new Set(previewOrders.map(o => o.id)))
                      }} />
                  </th>
                  {['平台訂單編號', '訂單日期', '應開發票金額'].map(c =>
                    <th key={c} style={{ ...thT, position: 'static', textAlign: c === '應開發票金額' ? 'right' : 'left' }}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {previewOrders.length === 0 && (
                  <tr><td colSpan={4} style={{ ...tdT, textAlign: 'center', color: T.n600 }}>所有訂單都已有發票號碼</td></tr>
                )}
                {previewOrders.map((o, i) => (
                  <tr key={i} style={{ background: ordInvEntryChecked.has(o.id) ? T.a100 : 'transparent' }}>
                    <td style={tdT}><input type="checkbox" checked={ordInvEntryChecked.has(o.id)}
                      onChange={() => {
                        const s = new Set(ordInvEntryChecked)
                        s.has(o.id) ? s.delete(o.id) : s.add(o.id)
                        setOrdInvEntryChecked(s)
                      }} /></td>
                    <td style={{ ...tdT, fontFamily: 'monospace' }}>{o.ref_no}</td>
                    <td style={tdT}>{o.order_date || '—'}</td>
                    <td style={{ ...tdT, textAlign: 'right' }}>{ordInvBase(o).toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {(hasSel || (ordInvEntryPreview && ordInvEntryMethod === 'auto')) && (
          <SumBar tone={isMatch ? 'ok' : amtNum > 0 && diff !== 0 ? 'bad' : 'flat'}>
            <span style={{ fontSize: 13 }}>訂單筆數 <strong>{selectedOrders.length}</strong></span>
            <span style={{ fontSize: 13 }}>應開發票金額 <strong>{totalSum.toLocaleString()}</strong></span>
            <span style={{ fontSize: 12, color: T.n600 }}>內扣後入帳 {payableSum.toLocaleString()}</span>
            {amtNum > 0 && <span style={{ fontSize: 13 }}>發票金額 <strong>{amtNum.toLocaleString()}</strong></span>}
            {amtNum > 0 && (
              <span style={{ fontSize: 13, color: isMatch ? T.g700 : T.danger, fontWeight: 700 }}>
                差異 {diff.toLocaleString()}　{isMatch ? '✓ 相符' : '✗ 有差異'}
              </span>
            )}
            <button onClick={runApplyOrdInvEntry} style={{ ...btnPri, marginLeft: 'auto' }}>套用</button>
          </SumBar>
        )}
        <PanelMsg text={ordInvEntryMsg} bad={/錯誤|請/} />
      </div>
    )
  })()

  // ── 面板：手續費發票核對 ──
  stepPanel.feeInv = (
    <div>
      <div style={panelLead}>
        依發票號碼分組核對：一張發票可群組多筆訂單，發票金額須等於該組訂單的手續費合計。左側色帶與下方明細表的訂單相互對應。
      </div>
      <InvoiceGroupCards
        groups={invoiceGroups} colorIdx={invColorIdx} kind="fee" storeKey={gateway}
        orders={shownOrders} onOpen={k => setViewInvKey(k)} />
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginTop: 18, alignItems: 'flex-end' }}>
        <Field label="發票號碼">
          <input value={invNo} onChange={e => setInvNo(e.target.value)} placeholder="AB-12345678" style={inpT} />
        </Field>
        <Field label="發票日期">
          <input type="date" value={invDate} onChange={e => setInvDate(e.target.value)} style={inpT} />
        </Field>
        <Field label="發票金額">
          <input type="number" value={invAmount} onChange={e => setInvAmount(e.target.value)} placeholder="0" style={inpT} />
        </Field>
      </div>

      <SegBtns value={invMethod} onChange={switchMethod} />

      {invMethod === 'auto' && (
        <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <Field label="期間起"><input type="date" value={invFrom} onChange={e => setInvFrom(e.target.value)} style={inpT} /></Field>
          <Field label="期間訖"><input type="date" value={invTo} onChange={e => setInvTo(e.target.value)} style={inpT} /></Field>
          <div style={{ paddingBottom: 10 }}><button onClick={runInvPreviewAuto} style={btnPri}>查詢</button></div>
        </div>
      )}

      {invMethod === 'manual' && (
        <div style={pickWrap}>
          <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ ...thT, position: 'static' }}>
                  <input type="checkbox"
                    checked={manualOrders.length > 0 && checkedIds.size === manualOrders.length}
                    onChange={toggleAll} />
                </th>
                {['銷貨單號', '入帳日', '手續費'].map(c =>
                  <th key={c} style={{ ...thT, position: 'static', textAlign: c === '手續費' ? 'right' : 'left' }}>{c}</th>)}
              </tr>
            </thead>
            <tbody>
              {manualOrders.length === 0 && (
                <tr><td colSpan={4} style={{ ...tdT, textAlign: 'center', color: T.n600 }}>所有訂單都已歸發票</td></tr>
              )}
              {manualOrders.map((o, i) => (
                <tr key={i} style={{ background: checkedIds.has(o.id) ? T.a100 : 'transparent' }}>
                  <td style={tdT}><input type="checkbox" checked={checkedIds.has(o.id)} onChange={() => toggleCheck(o.id)} /></td>
                  <td style={{ ...tdT, fontFamily: 'monospace' }}>{o.ref_no}</td>
                  <td style={tdT}>{o.in_date || o.order_date || '—'}</td>
                  <td style={{ ...tdT, textAlign: 'right' }}>{(o.fee_total ?? 0).toLocaleString()}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {hasInvOrders && (
        <SumBar tone={invIsMatch ? 'ok' : invDiff != null ? 'bad' : 'flat'}>
          {invFeeSum != null && <span style={{ fontSize: 13 }}>手續費加總 <strong>{invFeeSum.toLocaleString()}</strong></span>}
          {invAmountNum > 0 && <span style={{ fontSize: 13 }}>發票金額 <strong>{invAmountNum.toLocaleString()}</strong></span>}
          {invDiff != null && (
            <span style={{ fontSize: 13, color: invIsMatch ? T.g700 : T.danger, fontWeight: 700 }}>
              差異 {invDiff.toLocaleString()}　{invIsMatch ? '✓ 相符' : '✗ 有差異'}
            </span>
          )}
          <button onClick={runApplyInvoice} style={{ ...btnPri, marginLeft: 'auto' }}>套用</button>
        </SumBar>
      )}
      <PanelMsg text={invMsg} bad={/錯誤/} />
    </div>
  )

  // ── 面板：PayUni 交易處理費發票核對（官網 LINE Pay）──
  stepPanel.txFeeInv = !isLinePayOfficial ? null : (() => {
    const inv3FeeSum = inv3Preview?.txFeeSum ?? (inv3Method === 'manual' ? manual3TxFeeSum : null)
    const inv3AmountNum = parseFloat(inv3Amount) || 0
    const inv3Diff = inv3AmountNum > 0 && inv3FeeSum != null ? Math.round((inv3AmountNum - inv3FeeSum) * 100) / 100 : null
    const inv3IsMatch = inv3Diff != null && Math.abs(inv3Diff) < 0.01
    const hasInv3Orders = inv3Preview?.orders?.length > 0 || (inv3Method === 'manual' && checked3Ids.size > 0)
    return (
      <div>
        <div style={panelLead}>
          PayUni 服務費（0.2%）是帳戶層月結總額發票，與 LINE Pay 手續費（2.2%）那張分開核對。
        </div>
        <InvoiceGroupCards
          groups={txInvoiceGroups} colorIdx={{}} kind="tx" storeKey={gateway}
          orders={shownOrders} onOpen={k => setViewTxInvKey(k)} />
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginTop: 18, alignItems: 'flex-end' }}>
          <Field label="發票號碼">
            <input value={inv3No} onChange={e => setInv3No(e.target.value)} placeholder="AB-12345678" style={inpT} />
          </Field>
          <Field label="發票日期">
            <input type="date" value={inv3Date} onChange={e => setInv3Date(e.target.value)} style={inpT} />
          </Field>
          <Field label="發票金額">
            <input type="number" value={inv3Amount} onChange={e => setInv3Amount(e.target.value)} placeholder="0" style={inpT} />
          </Field>
        </div>

        <SegBtns value={inv3Method} onChange={switchMethod3} />

        {inv3Method === 'auto' && (
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end', flexWrap: 'wrap' }}>
            <Field label="期間起"><input type="date" value={inv3From} onChange={e => setInv3From(e.target.value)} style={inpT} /></Field>
            <Field label="期間訖"><input type="date" value={inv3To} onChange={e => setInv3To(e.target.value)} style={inpT} /></Field>
            <div style={{ paddingBottom: 10 }}><button onClick={runInv3PreviewAuto} style={btnPri}>查詢</button></div>
          </div>
        )}

        {inv3Method === 'manual' && (
          <div style={pickWrap}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
              <thead>
                <tr>
                  <th style={{ ...thT, position: 'static' }}>
                    <input type="checkbox"
                      checked={manual3Orders.length > 0 && checked3Ids.size === manual3Orders.length}
                      onChange={toggleAll3} />
                  </th>
                  {['平台訂單編號', '入帳日', '交易處理費'].map(c =>
                    <th key={c} style={{ ...thT, position: 'static', textAlign: c === '交易處理費' ? 'right' : 'left' }}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {manual3Orders.length === 0 && (
                  <tr><td colSpan={4} style={{ ...tdT, textAlign: 'center', color: T.n600 }}>所有訂單都已歸發票</td></tr>
                )}
                {manual3Orders.map((o, i) => (
                  <tr key={i} style={{ background: checked3Ids.has(o.id) ? T.a100 : 'transparent' }}>
                    <td style={tdT}><input type="checkbox" checked={checked3Ids.has(o.id)} onChange={() => toggleCheck3(o.id)} /></td>
                    <td style={{ ...tdT, fontFamily: 'monospace' }}>{o.ref_no}</td>
                    <td style={tdT}>{o.in_date || o.order_date || '—'}</td>
                    <td style={{ ...tdT, textAlign: 'right' }}>{(o.tx_fee ?? 0).toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {hasInv3Orders && (
          <SumBar tone={inv3IsMatch ? 'ok' : inv3Diff != null ? 'bad' : 'flat'}>
            {inv3FeeSum != null && <span style={{ fontSize: 13 }}>交易處理費加總 <strong>{inv3FeeSum.toLocaleString()}</strong></span>}
            {inv3AmountNum > 0 && <span style={{ fontSize: 13 }}>發票金額 <strong>{inv3AmountNum.toLocaleString()}</strong></span>}
            {inv3Diff != null && (
              <span style={{ fontSize: 13, color: inv3IsMatch ? T.g700 : T.danger, fontWeight: 700 }}>
                差異 {inv3Diff.toLocaleString()}　{inv3IsMatch ? '✓ 相符' : '✗ 有差異'}
              </span>
            )}
            <button onClick={runApplyTxFeeInvoice} style={{ ...btnPri, marginLeft: 'auto' }}>套用</button>
          </SumBar>
        )}
        <PanelMsg text={inv3Msg} bad={/錯誤/} />

        {/* 帳戶層費用備註（PayUni 服務費，不逐筆扣） */}
        <div style={{ marginTop: 18, borderTop: `1px solid ${T.divider}`, paddingTop: 16 }}>
          <div style={{ fontSize: 13, color: T.n700, marginBottom: 10 }}>
            帳戶層服務費備註：一筆涵蓋多單、不逐筆扣，僅記錄到訂單的 <code>account_fee_note</code>。
          </div>
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
            <Field label="發票號碼"><input value={inv2No} onChange={e => setInv2No(e.target.value)} placeholder="XC19745594" style={inpT} /></Field>
            <Field label="發票日期"><input type="date" value={inv2Date} onChange={e => setInv2Date(e.target.value)} style={inpT} /></Field>
            <Field label="發票金額"><input type="number" value={inv2Amount} onChange={e => setInv2Amount(e.target.value)} placeholder="0" style={inpT} /></Field>
            <div style={{ paddingBottom: 10 }}><button onClick={applyPayuniAccountFee} style={btnSec}>記錄備註</button></div>
          </div>
          <PanelMsg text={inv2Msg} bad={/錯誤|請/} />
        </div>
      </div>
    )
  })()

  // ── 訂單明細表欄位定義 ──
  const cols = [
    { label: '平台訂單編號', key: 'ref_no' },
    { label: '銷貨單號', key: 'sa_no' },
    { label: '訂單發票號碼', key: 'order_invoice_no' },
    ...(isShopee ? [{ label: '代開發票金額', key: null, align: 'right' }] : []),
    ...(isShopee ? [] : [{ label: '對應碼', key: 'tx_code' }]),
    { label: '訂單日期', key: 'order_date' },
    { label: '應收', key: 'total', align: 'right' },
    { label: '手續費', key: 'fee_total', align: 'right' },
    ...(isShopee ? [] : [{ label: '交易處理費', key: 'tx_fee', align: 'right' }]),
    { label: '應入帳', key: 'payable', align: 'right' },
    { label: '實際入帳', key: 'actual_in', align: 'right' },
    { label: '入帳日', key: 'in_date' },
    { label: '差異', key: null, align: 'right', anchor: true },
    { label: '狀態', key: 'recon_status', anchor: true },
    { label: '手續費發票', key: 'fee_invoice_no', min: 180 },
    ...(isLinePayOfficial ? [{ label: '交易處理費發票', key: 'tx_fee_invoice_no', min: 160 }] : []),
    { label: '', key: null, align: 'right' },
  ]

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>

      {/* ── 進度總覽 ── */}
      <div style={{ background: T.n100, borderRadius: T.rCard, boxShadow: T.shadowSm, padding: '26px 30px',
        display: 'grid', gridTemplateColumns: 'minmax(0,320px) 1fr', gap: 36, alignItems: 'center' }}>
        <div>
          <div style={{ fontSize: 11, letterSpacing: '.1em', color: T.a, marginBottom: 6 }}>
            {(filterMonth || '全部期間').toUpperCase()} · {gwInfo.label}
          </div>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: 12 }}>
            <span style={{ fontSize: 64, lineHeight: .9, fontWeight: 700, color: T.navy }}>{donePct}%</span>
            <span style={{ fontSize: 14, color: T.n700 }}>已完成對帳</span>
          </div>
          <div style={{ marginTop: 14, fontSize: 15 }}>
            <strong style={{ fontSize: 18 }}>{cDone}</strong>
            <span style={{ color: T.n600 }}> / {heroTotal} 筆</span>
            <span style={{ marginLeft: 12, color: T.a700, fontWeight: 600 }}>尚有 {heroTotal - cDone} 筆待對</span>
          </div>
        </div>
        <div>
          <div style={{ fontSize: 12, color: T.n600, marginBottom: 9 }}>訂單進度漏斗 · 依生命週期</div>
          <div style={{ display: 'flex', height: 26, borderRadius: T.rPill, overflow: 'hidden', background: T.n200 }}>
            <div style={{ width: `${segW(cShipped)}%`, background: T.n300 }} />
            <div style={{ width: `${segW(cSettled)}%`, background: T.g300 }} />
            <div style={{ width: `${segW(cPaid)}%`, background: T.a400 }} />
            <div style={{ width: `${segW(cDone)}%`, background: T.g600 }} />
          </div>
          <div style={{ display: 'flex', gap: 22, marginTop: 14, flexWrap: 'wrap' }}>
            {[['已出貨', T.n300, cShipped], ['平台已結算', T.g300, cSettled],
              ['已入帳', T.a400, cPaid], ['已對帳', T.g600, cDone]].map(([lbl, color, n]) => (
              <div key={lbl} style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                <span style={{ width: 11, height: 11, borderRadius: 3, background: color }} />
                <span style={{ fontSize: 13, color: T.n700 }}>{lbl}</span>
                <strong>{n}</strong>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── 對帳步驟工作區 ── */}
      <div style={{ background: T.n100, borderRadius: T.rCard, boxShadow: T.shadowSm, padding: '24px 28px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
          marginBottom: 18, flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, flexWrap: 'wrap' }}>
            <h3 style={{ margin: 0, fontSize: 22, color: T.navy }}>本期對帳步驟</h3>
            <span style={{ fontSize: 13, color: T.n600 }}>依序完成 · 每一步的檔案就在該步驟裡</span>
          </div>
          <button onClick={() => setSopOpen(true)} style={{ ...btnSec, display: 'inline-flex', alignItems: 'center', gap: 7 }}>
            <IconBook />教學 SOP
          </button>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${stepCols},minmax(0,1fr))`, gap: 12 }}>
          {stepDefs.map((s, i) => {
            const active = s.key === curStep
            const state = s.done ? 'done' : (s.key === firstTodoKey ? 'current' : 'todo')
            return (
              <button key={s.key} onClick={() => setActiveStepKey(s.key)} aria-pressed={active}
                style={{
                  border: `1.5px solid ${active ? T.a : state === 'current' ? T.a300 : T.divider}`,
                  background: active ? T.a100 : state === 'done' ? T.g100 : T.bg,
                  borderRadius: T.rPanel, padding: '14px 16px', cursor: 'pointer',
                  transition: 'background .12s,border-color .12s',
                  font: 'inherit', color: 'inherit', textAlign: 'left', display: 'block', width: '100%',
                }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span style={{ width: 28, height: 28, flex: 'none', borderRadius: '50%', display: 'grid',
                    placeItems: 'center', fontSize: 14, fontWeight: 700,
                    background: state === 'done' ? T.g600 : state === 'current' ? T.a : T.n300,
                    color: state === 'todo' ? T.n700 : T.bg }}>
                    {state === 'done' ? '✓' : i + 1}
                  </span>
                  <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{s.title}</span>
                </div>
                <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 6 }}>
                  <span style={{ fontSize: 12, color: T.n600, whiteSpace: 'nowrap' }}>{s.sub}</span>
                  <Tag tone={state === 'done' ? 'accent2' : state === 'current' ? 'accent' : 'neutral'}>
                    {state === 'done' ? '完成' : state === 'current' ? '進行中' : '待處理'}
                  </Tag>
                </div>
              </button>
            )
          })}
        </div>

        <div style={{ marginTop: 20, borderTop: `1px solid ${T.divider}`, paddingTop: 20 }}>
          {stepPanel[curStep] || <div style={{ fontSize: 13, color: T.n600 }}>此金流不需要這個步驟。</div>}
        </div>
      </div>

      {/* ── 訂單明細表 ── */}
      <div style={{ background: T.n100, borderRadius: T.rCard, boxShadow: T.shadowSm, padding: '22px 26px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap', marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 22, color: T.navy, whiteSpace: 'nowrap' }}>訂單明細</h3>
          <div style={{ display: 'flex', gap: 7, flexWrap: 'wrap' }}>
            {chipDefs.map(c => {
              const on = c.value === filterStatus
              return (
                <button key={c.label} onClick={() => setFilterStatus(c.value)} aria-pressed={on}
                  style={{ padding: '6px 13px', borderRadius: T.rPill, fontSize: 13, cursor: 'pointer',
                    fontFamily: 'inherit',
                    border: `1px solid ${on ? T.a : T.divider}`, background: on ? T.a100 : 'transparent',
                    color: on ? T.a700 : T.n700, whiteSpace: 'nowrap', transition: 'background .12s,border-color .12s' }}>
                  {c.label} {c.count}
                </button>
              )
            })}
          </div>
          <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            <select value={filterMonth} onChange={e => setFilterMonth(e.target.value)} style={{ ...inpT, width: 'auto' }}>
              <option value="">全部月份</option>
              {months.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
            <input value={query} onChange={e => setQuery(e.target.value)} placeholder="搜尋訂單編號…"
              style={{ ...inpT, width: 180 }} />
            {selectedIds.size > 0 && (
              <button onClick={deleteSelected} style={{ ...btnSec, color: T.danger, borderColor: T.danger }}>
                刪除 {selectedIds.size} 筆
              </button>
            )}
            {deleteMsg && <span style={{ fontSize: 12, color: deleteMsg.includes('錯誤') ? T.danger : T.n600 }}>{deleteMsg}</span>}
            <button onClick={loadOrders} style={btnSec}>重新整理</button>
            <button onClick={exportOrders} style={btnSec}>匯出</button>
          </div>
        </div>

        <div style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: '62vh',
          borderRadius: T.rInner, border: `1px solid ${T.divider}`, background: T.surface }}>
          <table style={{ borderCollapse: 'collapse', width: '100%', minWidth: 920, fontSize: 13 }}>
            <thead>
              <tr>
                <th style={{ ...thT, width: 34 }}>
                  <input type="checkbox"
                    checked={searched.length > 0 && selectedIds.size === searched.length}
                    onChange={() => setSelectedIds(prev => prev.size === searched.length ? new Set() : new Set(searched.map(o => o.id)))} />
                </th>
                {cols.map(c => {
                  const active = c.key && sortCol === c.key
                  return (
                    <th key={c.label || 'edit'}
                      onClick={() => c.key && handleSort(c.key)}
                      style={{ ...(c.anchor ? thAnchor : thT), textAlign: c.align || 'left',
                        minWidth: c.min, cursor: c.key ? 'pointer' : 'default', userSelect: 'none' }}>
                      {c.label}{active ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {(() => {
                const seenInv = new Set()
                const seenTxInv = new Set()
                const seenOrdInv = new Set()
                const seenBankGroup = new Set()
                return searched.map((o, i) => {
                  const ci = o.fee_invoice_no ? invColorIdx[o.fee_invoice_no] : null
                  const invBg = ci != null ? INV_BG[ci] : undefined
                  const groupBar = ci != null ? INV_BAR[ci] : 'transparent'
                  const rowBg = selectedIds.has(o.id) ? T.a200 : invBg
                  const isFirstInv = o.fee_invoice_no && !seenInv.has(o.fee_invoice_no)
                  if (o.fee_invoice_no) seenInv.add(o.fee_invoice_no)
                  const grp = o.fee_invoice_no ? invoiceGroups[o.fee_invoice_no] : null
                  const isFirstTxInv = o.tx_fee_invoice_no && !seenTxInv.has(o.tx_fee_invoice_no)
                  if (o.tx_fee_invoice_no) seenTxInv.add(o.tx_fee_invoice_no)
                  const txGrp = o.tx_fee_invoice_no ? txInvoiceGroups[o.tx_fee_invoice_no] : null
                  const isFirstOrdInv = o.order_invoice_no && !seenOrdInv.has(o.order_invoice_no)
                  if (o.order_invoice_no) seenOrdInv.add(o.order_invoice_no)
                  const ordInvGrp = o.order_invoice_no ? ordInvGroups[o.order_invoice_no] : null
                  const bankGroupKey = (showBankGroup && o.recon_status === '已入帳' && o.in_date) ? o.in_date.slice(0, 10) : null
                  const isFirstBankGroup = bankGroupKey && !seenBankGroup.has(bankGroupKey)
                  if (bankGroupKey) seenBankGroup.add(bankGroupKey)
                  const bankGrp = bankGroupKey ? bankGroups[bankGroupKey] : null
                  const diff = calcDiff(o)
                  const feeMatch = grp && grp.invAmount != null ? Math.abs(grp.feeSum - grp.invAmount) < 0.01 : null
                  return (
                    <tr key={i} className="hhy-row" style={{ background: rowBg }}>
                      <td style={{ ...tdT, borderLeft: `3px solid ${groupBar}` }}>
                        <input type="checkbox" checked={selectedIds.has(o.id)} onChange={() => toggleSelect(o.id)} />
                      </td>
                      <td style={{ ...tdT, fontFamily: 'monospace' }}>{o.ref_no}</td>
                      <td style={{ ...tdT, fontFamily: 'monospace', fontSize: 12 }}>{o.sa_no || '—'}</td>
                      <td style={{ ...tdT, fontFamily: 'monospace', fontSize: 12, cursor: o.order_invoice_no ? 'pointer' : 'default' }}
                        onClick={o.order_invoice_no ? () => { setViewOrdInvKey(o.order_invoice_no); setOrdInvDeleteConfirm(false); setOrdInvPopupDate(ordInvGroups[o.order_invoice_no]?.invDate || '') } : undefined}>
                        {isFirstOrdInv && ordInvGrp ? (
                          <div>
                            <div style={{ color: T.a700, textDecoration: 'underline' }}>{o.order_invoice_no}</div>
                            <div style={{ fontSize: 11, color: T.n600 }}>共 {ordInvGrp.count} 筆</div>
                          </div>
                        ) : o.order_invoice_no ? <span style={{ fontSize: 12, color: T.n500 }}>↳ 同張發票</span> : <span style={{ color: T.n400 }}>—</span>}
                      </td>
                      {/* 已開立 → 該組實際開立金額；未開立 → 由應入帳＋手續費回推的應開金額（灰字加註）。
                          未開立時刻意不看 order_invoice_amount：舊版 parser 曾把代收金額寫進該欄，
                          留下的殘值不會隨應收／手續費更動而同步。 */}
                      {isShopee && <td style={{ ...tdT, textAlign: 'right' }}>
                        {o.order_invoice_no
                          ? (o.order_invoice_amount != null ? o.order_invoice_amount.toLocaleString() : '—')
                          : <span style={{ color: T.n500 }}>
                              {ordInvBase(o).toLocaleString()}
                              <span style={{ fontSize: 10, color: T.n400, marginLeft: 4 }}>應開</span>
                            </span>}
                      </td>}
                      {!isShopee && <td style={{ ...tdT, fontFamily: 'monospace', fontSize: 11, color: T.n600 }}>{o.tx_code || '—'}</td>}
                      <td style={{ ...tdT, color: T.n700 }}>{o.order_date ? o.order_date.slice(0, 10) : '—'}</td>
                      <td style={{ ...tdT, textAlign: 'right', color: T.n700 }}>{o.total?.toLocaleString()}</td>
                      <td style={{ ...tdT, textAlign: 'right', color: T.n600 }}>{o.fee_total != null ? o.fee_total.toLocaleString() : '—'}</td>
                      {!isShopee && <td style={{ ...tdT, textAlign: 'right', color: T.n600 }}>{o.tx_fee != null ? o.tx_fee.toLocaleString() : '—'}</td>}
                      <td style={{ ...tdT, textAlign: 'right' }}>{o.payable != null ? o.payable.toLocaleString() : '—'}</td>
                      <td style={{ ...tdT, textAlign: 'right' }}>{o.actual_in != null ? o.actual_in.toLocaleString() : '—'}</td>
                      <td style={tdT}>
                        {isFirstBankGroup && bankGrp ? (
                          <div style={{ cursor: 'pointer' }} onClick={() => setViewBankGroupKey(bankGroupKey)}>
                            <div style={{ fontWeight: 600, fontSize: 12, color: T.a700, textDecoration: 'underline' }}>{bankGroupKey}</div>
                            <div style={{ fontSize: 11, color: T.n600, marginTop: 2, whiteSpace: 'nowrap' }}>
                              {bankGrp.count} 筆・合計 {bankGrp.bankDeposit != null ? Math.round(bankGrp.bankDeposit * 100) / 100 : '—'}
                            </div>
                          </div>
                        ) : bankGroupKey ? (
                          <span style={{ fontSize: 12, color: T.n500, cursor: 'pointer' }} onClick={() => setViewBankGroupKey(bankGroupKey)}>↳ 同批入帳</span>
                        ) : o.in_date || '—'}
                      </td>
                      {/* 差異可點擊記錄原因（寫入 diff_note）；已有註記的顯示在數字下方 */}
                      <td style={{ ...tdT, textAlign: 'right', whiteSpace: 'nowrap',
                        fontWeight: diff == null || diff === 0 ? 600 : 700,
                        cursor: diff == null ? 'default' : 'pointer',
                        color: diff == null ? T.n400 : diff === 0 ? T.g700 : T.danger }}
                        title={diff == null ? undefined : (o.diff_note || '點擊記錄差異原因')}
                        onClick={diff == null ? undefined : () => openDiffNote(o)}>
                        <div style={{ textDecoration: diff != null && diff !== 0 ? 'underline' : 'none' }}>
                          {diff == null ? '—' : diff === 0 ? '0 ✓' : (diff > 0 ? '+' + diff : String(diff))}
                        </div>
                        {o.diff_note && (
                          <div style={{ fontSize: 11, fontWeight: 400, color: T.n600, marginTop: 2,
                            maxWidth: 130, whiteSpace: 'normal', lineHeight: 1.4 }}>
                            {o.diff_note.length > 24 ? o.diff_note.slice(0, 24) + '…' : o.diff_note}
                          </div>
                        )}
                        {(o.diff_invoice_no || o.diff_invoice_pdf_url) && (
                          <div style={{ fontSize: 11, fontWeight: 400, color: T.a700, marginTop: 2, whiteSpace: 'nowrap' }}>
                            {o.diff_invoice_pdf_url ? '📎 ' : ''}{o.diff_invoice_no || '已附檔'}
                          </div>
                        )}
                      </td>
                      <td style={tdT}>
                        <Tag tone={statusTone(o.recon_status)}>{o.recon_status || '—'}</Tag>
                      </td>
                      <td style={{ ...tdT, cursor: o.fee_invoice_no ? 'pointer' : 'default' }}
                        onClick={o.fee_invoice_no ? () => setViewInvKey(o.fee_invoice_no) : undefined}>
                        {isFirstInv && grp ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                              <span style={{ fontFamily: 'monospace', fontSize: 13, color: T.text }}>{o.fee_invoice_no}</span>
                              {feeMatch != null && (
                                <Tag tone={feeMatch ? 'accent2' : 'danger'} style={{ fontSize: 11 }}>
                                  {feeMatch ? '✓ 相符' : `✗ 差異 ${Math.round((grp.feeSum - grp.invAmount) * 100) / 100}`}
                                </Tag>
                              )}
                            </div>
                            <div style={{ fontSize: 11, color: T.n600, whiteSpace: 'nowrap' }}>
                              共 {grp.count} 筆 · 費用 {Math.round(grp.feeSum * 100) / 100} / 發票 {grp.invAmount != null ? grp.invAmount.toLocaleString() : '—'}
                            </div>
                          </div>
                        ) : o.fee_invoice_no
                          ? <span style={{ fontSize: 12, color: T.n500 }}>↳ 同張發票</span>
                          : <span style={{ fontSize: 12, color: T.n400 }}>未開立</span>}
                      </td>
                      {isLinePayOfficial && (
                        <td style={{ ...tdT, cursor: o.tx_fee_invoice_no ? 'pointer' : 'default' }}
                          onClick={o.tx_fee_invoice_no ? () => setViewTxInvKey(o.tx_fee_invoice_no) : undefined}>
                          {isFirstTxInv && txGrp ? (
                            <div>
                              <div style={{ fontFamily: 'monospace', fontSize: 13, color: T.a700, textDecoration: 'underline' }}>{o.tx_fee_invoice_no}</div>
                              <div style={{ fontSize: 11, color: T.n600, marginTop: 2, whiteSpace: 'nowrap' }}>
                                共 {txGrp.count} 筆 · 發票 {txGrp.invAmount != null ? txGrp.invAmount.toLocaleString() : '—'}
                              </div>
                            </div>
                          ) : o.tx_fee_invoice_no
                            ? <span style={{ fontSize: 12, color: T.n500 }}>↳ 同張發票</span>
                            : <span style={{ fontSize: 12, color: T.n400 }}>未開立</span>}
                        </td>
                      )}
                      <td style={{ ...tdT, textAlign: 'right' }}>
                        <button onClick={() => { setEditOrder({ ...o }); setEditMsg('') }} style={btnGhostT}>編輯</button>
                      </td>
                    </tr>
                  )
                })
              })()}
              {searched.length === 0 && (
                <tr><td colSpan={cols.length + 1} style={{ ...tdT, textAlign: 'center', color: T.n600, padding: 24 }}>沒有資料</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 12, fontSize: 12, color: T.n600 }}>
          顯示 {searched.length} 筆 · 差異與狀態欄以顏色標示，一眼看出需要處理的訂單。
        </div>
        {/* 蝦皮每月最後一筆入帳會被扣掉上個月的代開發票總額，差異屬正常，提醒別誤判 */}
        {isShopee && (
          <div style={{ marginTop: 18, border: `1.5px solid ${T.divider}`, background: T.n100,
            borderRadius: T.rPanel, padding: '18px 22px' }}>
            <div style={{ fontSize: 20, fontWeight: 700, color: T.navy, lineHeight: 1.5 }}>
              每月會有一筆金額入帳有差異，是前月代開發票總額
            </div>
            <div style={{ marginTop: 8, fontSize: 13, color: T.n600, lineHeight: 1.7 }}>
              例：六月代開費用於七月最後一筆入帳金額中扣除
            </div>
          </div>
        )}
      </div>

      {/* ── 教學 SOP 側欄：依對帳步驟分卡，每步一段文字 + 一張截圖 ── */}
      {sopOpen && (
        <div onClick={() => { closeSopDrawer() }}
          style={{ position: 'fixed', inset: 0, zIndex: 60, background: 'rgba(22,35,63,.42)',
            display: 'flex', justifyContent: 'flex-end' }}>
          <div onClick={e => e.stopPropagation()}
            style={{ width: 'min(460px,92vw)', height: '100%', background: T.n100, boxShadow: T.shadowLg,
              borderRadius: '28px 0 0 28px', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
            <div style={{ padding: '22px 26px', borderBottom: `1px solid ${T.divider}`,
              display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <div>
                <div style={{ fontSize: 11, letterSpacing: '.1em', color: T.a }}>教學 SOP</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: T.navy }}>{gwInfo.label} · 教學 SOP</div>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                {!sopEditing
                  ? <button onClick={() => setSopEditing(true)} style={{ ...btnSec, fontSize: 13 }}>編輯</button>
                  : <>
                      <button onClick={openSopLinkForm} style={{ ...btnSec, fontSize: 13 }}>連結</button>
                      <button onClick={saveSop} disabled={sopSaving} style={{ ...btnPri, fontSize: 13 }}>{sopSaving ? '儲存中…' : '完成'}</button>
                      <button onClick={cancelSop} style={{ ...btnSec, fontSize: 13 }}>取消</button>
                    </>}
                <button onClick={() => { closeSopDrawer() }}
                  aria-label="關閉" style={{ ...btnSec, padding: '8px 12px' }}>✕</button>
              </div>
            </div>

            <div style={{ padding: '22px 26px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div style={{ fontSize: 13, color: T.n600 }}>
                依照對帳步驟編排，每一步要在哪個後台、怎麼下載檔案都寫在這裡。
                {sopEditing ? '可直接編輯文字、插入連結，並為每一步拖入或選擇一張截圖。' : ''}
              </div>

              {sopMsg && <div style={{ fontSize: 13, color: T.danger }}>{sopMsg}</div>}

              {sopEditing && sopLinkForm && (
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap',
                  padding: '10px 14px', background: T.a100, borderRadius: T.rInner }}>
                  <span style={{ fontSize: 12, color: T.n600, flexShrink: 0 }}>插入連結</span>
                  <input value={sopLinkText} onChange={e => setSopLinkText(e.target.value)} placeholder="顯示文字" style={{ ...inpT, width: 120 }} />
                  <input value={sopLinkUrl} onChange={e => setSopLinkUrl(e.target.value)} placeholder="https://..."
                    style={{ ...inpT, width: 180 }} onKeyDown={e => e.key === 'Enter' && insertSopLink()} />
                  <button onClick={insertSopLink} style={{ ...btnPri, fontSize: 13 }}>插入</button>
                  <button onClick={() => setSopLinkForm(false)} style={{ ...btnSec, fontSize: 13 }}>取消</button>
                </div>
              )}

              {stepDefs.map((s, i) => {
                const d = sopStepData(s.key)
                const showSlot = sopEditing || !!d.img   // 定稿模式沒圖就整格隱藏
                return (
                  <div key={s.key} style={{ background: T.surface, borderRadius: T.rPanel,
                    padding: '16px 18px', boxShadow: T.shadowSm }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                      <span style={{ width: 24, height: 24, flex: 'none', borderRadius: '50%', background: T.a,
                        color: '#fff', display: 'grid', placeItems: 'center', fontSize: 13, fontWeight: 700 }}>{i + 1}</span>
                      <span style={{ fontSize: 16, fontWeight: 700 }}>{s.title}</span>
                    </div>

                    {SOURCE_HINT[s.key] && (
                      <div style={{ fontSize: 12, color: T.n600, lineHeight: 1.7, marginBottom: 8 }}>
                        <strong style={{ color: T.navy }}>預設路徑：</strong>{SOURCE_HINT[s.key]}
                      </div>
                    )}

                    {sopEditing ? (
                      <div
                        ref={el => { sopRefs.current[s.key] = el }}
                        contentEditable
                        suppressContentEditableWarning
                        onFocus={() => { sopFocusEl.current = sopRefs.current[s.key] }}
                        onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); document.execCommand('insertLineBreak') } }}
                        style={{
                          minHeight: 70, padding: '10px 12px', borderRadius: T.rInner,
                          border: `1.5px solid ${T.a}`, background: T.surface,
                          fontSize: 13, lineHeight: 1.85, outline: 'none',
                          whiteSpace: 'pre-wrap', wordBreak: 'break-word',
                        }}
                      />
                    ) : (
                      <div
                        style={{ fontSize: 13, lineHeight: 1.85, color: T.n700,
                          whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}
                        dangerouslySetInnerHTML={{
                          __html: d.html ||
                            `<span style="color:${T.n500};font-style:italic">尚未填寫，點右上角「編輯」補上</span>`
                        }}
                      />
                    )}

                    {showSlot && (
                      <div
                        onDragOver={sopEditing ? (e => e.preventDefault()) : undefined}
                        onDrop={sopEditing ? (e => {
                          e.preventDefault()
                          const f = e.dataTransfer.files?.[0]
                          if (f && f.type.startsWith('image/')) uploadSopStepImage(s.key, f)
                        }) : undefined}
                        style={{ marginTop: 12, borderRadius: T.rInner, overflow: 'hidden',
                          border: d.img ? `1px solid ${T.divider}` : `1px dashed ${T.divider}`,
                          background: d.img ? T.surface : T.bg, position: 'relative',
                          minHeight: d.img ? undefined : 140,
                          display: d.img ? 'block' : 'grid', placeItems: 'center' }}>
                        {d.img ? (
                          <>
                            <a href={d.img} target="_blank" rel="noopener noreferrer">
                              <img src={d.img} alt={`${s.title} 截圖`} style={{ width: '100%', display: 'block' }} />
                            </a>
                            {sopEditing && (
                              <div style={{ position: 'absolute', top: 8, right: 8, display: 'flex', gap: 6 }}>
                                <label style={{ ...slotChipBtn, display: 'inline-block' }}>
                                  {sopUploading === s.key ? '上傳中…' : '更換'}
                                  <input type="file" accept="image/*" style={{ display: 'none' }}
                                    onChange={e => { const f = e.target.files?.[0]; if (f) uploadSopStepImage(s.key, f); e.target.value = '' }} />
                                </label>
                                <button onClick={() => removeSopStepImage(s.key)} aria-label="移除截圖" style={slotChipBtn}>
                                  移除
                                </button>
                              </div>
                            )}
                          </>
                        ) : (
                          <label style={{ display: 'grid', placeItems: 'center', gap: 6, padding: 16,
                            cursor: 'pointer', textAlign: 'center', width: '100%' }}>
                            <input type="file" accept="image/*" style={{ display: 'none' }}
                              onChange={e => { const f = e.target.files?.[0]; if (f) uploadSopStepImage(s.key, f); e.target.value = '' }} />
                            <span style={{ fontSize: 20, color: T.n400 }}>＋</span>
                            <span style={{ fontSize: 12, color: T.n600 }}>
                              {sopUploading === s.key ? '上傳中…' : `拖入或點擊上傳截圖：${s.title}（選填）`}
                            </span>
                          </label>
                        )}
                      </div>
                    )}
                  </div>
                )
              })}

              <div style={{ border: `1px dashed ${T.divider}`, borderRadius: T.rPanel, padding: '14px 18px',
                fontSize: 12, color: T.n600 }}>
                💡 這份 SOP 會隨通路切換各自獨立，每個平台（酷澎／蝦皮／官網／LINE商城）都有自己的一份。
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 入帳差異原因 modal */}
      {diffNoteOrder && (() => {
        const d = calcDiff(diffNoteOrder)
        return (
          <div style={overlay} onClick={() => setDiffNoteOrder(null)}>
            <div style={modal} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>入帳差異原因</h3>
              <p style={{ fontSize: 13, color: T.n600, margin: '0 0 12px', fontFamily: 'monospace' }}>{diffNoteOrder.ref_no}</p>
              <div style={{ display: 'flex', gap: 18, fontSize: 13, background: T.n100,
                borderRadius: T.rInner, padding: '10px 14px', marginBottom: 14 }}>
                <span style={{ color: T.n700 }}>應入帳 <strong style={{ color: T.text }}>{diffNoteOrder.payable != null ? diffNoteOrder.payable.toLocaleString() : '—'}</strong></span>
                <span style={{ color: T.n700 }}>實際入帳 <strong style={{ color: T.text }}>{diffNoteOrder.actual_in != null ? diffNoteOrder.actual_in.toLocaleString() : '—'}</strong></span>
                <span style={{ color: T.n700, marginLeft: 'auto' }}>差異 <strong style={{ color: d === 0 ? T.g700 : T.danger }}>
                  {d == null ? '—' : d === 0 ? '0' : (d > 0 ? '+' + d : String(d))}</strong></span>
              </div>
              <Field label="原因／備註">
                <textarea value={diffNoteText} onChange={e => setDiffNoteText(e.target.value)} rows={3}
                  placeholder="例：扣除六月代開發票總額"
                  style={{ ...inpT, borderRadius: T.rInner, resize: 'vertical', lineHeight: 1.6 }} />
              </Field>
              <p style={{ fontSize: 12, color: T.n600, margin: '4px 0 0' }}>
                存成訂單的「差異原因」，與出貨報表帶進來的買家備註分開。留空即清除。
              </p>

              <div style={{ borderTop: `1px solid ${T.divider}`, margin: '16px 0 12px', paddingTop: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: T.n800, marginBottom: 8 }}>對應發票</div>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  <Field label="發票號碼">
                    <input value={diffInvNo} onChange={e => setDiffInvNo(e.target.value)}
                      placeholder="AB-12345678" style={inpT} />
                  </Field>
                  <Field label="發票日期">
                    <input type="date" value={diffInvDate} onChange={e => setDiffInvDate(e.target.value)} style={inpT} />
                  </Field>
                  <Field label="金額">
                    <input type="number" value={diffInvAmount} onChange={e => setDiffInvAmount(e.target.value)}
                      placeholder="0" style={inpT} />
                  </Field>
                </div>
                <Field label="附件">
                  <input ref={diffInvFileRef} type="file" accept="application/pdf,image/*"
                    style={{ display: 'none' }} onChange={uploadDiffInvFile} />
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                    <button onClick={() => diffInvFileRef.current.click()} style={{ ...btnSec, fontSize: 12 }} disabled={diffInvUploading}>
                      {diffInvUploading ? '上傳中…' : diffInvUrl ? '重新上傳' : '上傳檔案'}
                    </button>
                    {diffInvUrl && (
                      <>
                        <a href={diffInvUrl} target="_blank" rel="noopener noreferrer"
                          style={{ fontSize: 12, color: T.a700, textDecoration: 'underline' }}>查看附件</a>
                        <button onClick={removeDiffInvFile} disabled={diffInvUploading}
                          style={{ ...btnSec, fontSize: 12, color: T.danger, borderColor: T.danger }}>移除</button>
                      </>
                    )}
                  </div>
                  <div style={{ fontSize: 12, color: T.n600, marginTop: 6 }}>
                    支援 PDF 與圖片；選檔後立即上傳，不需按儲存。
                  </div>
                  {diffInvUploadError && <div style={{ fontSize: 12, color: T.danger, marginTop: 4 }}>{diffInvUploadError}</div>}
                </Field>
              </div>
              {diffNoteMsg && <p style={{ fontSize: 13, color: diffNoteMsg.includes('錯誤') ? T.danger : T.n600, margin: '4px 0' }}>{diffNoteMsg}</p>}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
                <button onClick={() => setDiffNoteOrder(null)} style={btnSec}>取消</button>
                <button onClick={saveDiffNote} style={btnPri}>儲存</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* 編輯 modal */}
      {editOrder && (
        <div style={overlay} onClick={() => setEditOrder(null)}>
          <div style={modal} onClick={e => e.stopPropagation()}>
            <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>編輯訂單</h3>
            <p style={{ fontSize: 13, color: T.n600, margin: '0 0 12px', fontFamily: 'monospace' }}>{editOrder.ref_no}</p>
            <Field label="銷貨單號（ERP SA 單號）">
              <input value={editOrder.sa_no || ''} onChange={e => setEditOrder(p => ({ ...p, sa_no: e.target.value }))}
                placeholder="SA-XXXXXXXX" style={inpT} />
            </Field>
            <Field label="狀態">
              <select value={editOrder.recon_status || ''} onChange={e => setEditOrder(p => ({ ...p, recon_status: e.target.value }))} style={inpT}>
                {['待出貨', '已出貨', '平台已結算', '已入帳', '已對帳'].map(s => <option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="訂單發票號碼">
              <input value={editOrder.order_invoice_no || ''} onChange={e => setEditOrder(p => ({ ...p, order_invoice_no: e.target.value }))}
                placeholder="AB-12345678" style={inpT} />
            </Field>
            <Field label="手續費發票號碼">
              <input value={editOrder.fee_invoice_no || ''} onChange={e => setEditOrder(p => ({ ...p, fee_invoice_no: e.target.value }))}
                placeholder="AB-12345678" style={inpT} />
            </Field>
            {isLinePayOfficial && (
              <Field label="交易處理費發票號碼">
                <input value={editOrder.tx_fee_invoice_no || ''} onChange={e => setEditOrder(p => ({ ...p, tx_fee_invoice_no: e.target.value }))}
                  placeholder="AB-12345678" style={inpT} />
              </Field>
            )}
            <div style={{ display: 'flex', gap: 10 }}>
              <Field label="應收">
                <input type="number" value={editOrder.total ?? ''} onChange={e => setEditOrder(p => ({ ...p, total: e.target.value }))} style={inpT} />
              </Field>
              <Field label="手續費">
                <input type="number" value={editOrder.fee_total ?? ''} onChange={e => setEditOrder(p => ({ ...p, fee_total: e.target.value }))} style={inpT} />
              </Field>
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <Field label="應入帳">
                <input type="number" value={editOrder.payable ?? ''} onChange={e => setEditOrder(p => ({ ...p, payable: e.target.value }))} style={inpT} />
              </Field>
              <Field label="實際入帳">
                <input type="number" value={editOrder.actual_in ?? ''} onChange={e => setEditOrder(p => ({ ...p, actual_in: e.target.value }))} style={inpT} />
              </Field>
              <Field label="入帳日">
                <input type="date" value={editOrder.in_date || ''} onChange={e => setEditOrder(p => ({ ...p, in_date: e.target.value }))} style={inpT} />
              </Field>
            </div>
            <Field label="備註">
              <textarea value={editOrder.note || ''} onChange={e => setEditOrder(p => ({ ...p, note: e.target.value }))} rows={4}
                style={{ ...inpT, borderRadius: T.rInner, resize: 'vertical', lineHeight: 1.6 }} />
            </Field>
            {editMsg && <p style={{ fontSize: 13, color: editMsg.includes('錯誤') ? T.danger : T.n600, margin: '4px 0' }}>{editMsg}</p>}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
              <button onClick={() => setEditOrder(null)} style={btnSec}>取消</button>
              <button onClick={() => saveEditOrder(editOrder)} style={btnPri}>儲存</button>
            </div>
          </div>
        </div>
      )}

      {/* 發票資訊 modal */}
      {viewInvKey && (() => {
        const grp = invoiceGroups[viewInvKey]
        if (!grp) return null
        const checkColor = grp.invoiceCheck === '相符' ? T.g700 : grp.invoiceCheck === '有差異' ? T.danger : T.n600
        const staticRows = [
          ['手續費發票號碼', viewInvKey, 'monospace'],
          ['手續費合計', `NT$ ${Math.round(grp.feeSum * 100) / 100}`, 'inherit'],
          ['包含訂單', `${grp.count} 筆`, 'inherit'],
          ['核對結果', grp.invoiceCheck || '—', 'inherit'],
        ]
        return (
          <div style={overlay} onClick={() => setViewInvKey(null)}>
            <div style={{ ...modal, width: 400 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>發票資訊</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <tbody>
                  <tr style={mRow}>
                    <td style={mLabel}>手續費發票號碼</td>
                    <td style={{ ...mVal, fontFamily: 'monospace' }}>{viewInvKey}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>發票日期</td>
                    <td style={{ padding: '6px 0' }}>
                      <input type="date" value={invPopupDate} onChange={e => setInvPopupDate(e.target.value)}
                        onBlur={e => saveInvPopupDate(e.target.value)} style={popupInp} />
                    </td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>發票金額</td>
                    <td style={{ padding: '6px 0' }}>
                      <input type="number" value={invPopupAmount} onChange={e => setInvPopupAmount(e.target.value)}
                        onBlur={e => saveInvPopupAmount(e.target.value)} placeholder="0" style={popupInp} />
                    </td>
                  </tr>
                  {staticRows.map(([label, val, ff]) => (
                    <tr key={label} style={mRow}>
                      <td style={mLabel}>{label}</td>
                      <td style={{ ...mVal, fontWeight: label === '核對結果' ? 700 : 400, color: label === '核對結果' ? checkColor : T.text, fontFamily: ff }}>{val}</td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ ...mLabel, verticalAlign: 'top' }}>備注</td>
                    <td style={mVal}>
                      <textarea value={invNote} onChange={e => setInvNote(e.target.value)}
                        onBlur={e => saveInvNote(e.target.value)} placeholder="輸入備注…" rows={3} style={popupArea} />
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...mLabel, verticalAlign: 'top' }}>發票 PDF</td>
                    <td style={mVal}>
                      <input ref={invPdfRef} type="file" accept="application/pdf" style={{ display: 'none' }} onChange={uploadInvPdf} />
                      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                        <button onClick={() => invPdfRef.current.click()} style={{ ...btnSec, fontSize: 12 }} disabled={invUploading}>
                          {invUploading ? '上傳中…' : invPdfUrl ? '重新上傳' : '上傳 PDF'}
                        </button>
                        {invPdfUrl && (
                          <a href={invPdfUrl} target="_blank" rel="noopener noreferrer"
                            style={{ fontSize: 12, color: T.a700, textDecoration: 'underline' }}>查看 PDF</a>
                        )}
                      </div>
                      {invUploadError && <div style={{ fontSize: 12, color: T.danger, marginTop: 4 }}>{invUploadError}</div>}
                    </td>
                  </tr>
                </tbody>
              </table>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 20 }}>
                {!invDeleteConfirm ? (
                  <button onClick={() => setInvDeleteConfirm(true)}
                    style={{ ...btnSec, color: T.danger, borderColor: T.danger, fontSize: 12 }}>
                    刪除發票資訊
                  </button>
                ) : (
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <span style={{ fontSize: 12, color: T.danger }}>確定刪除？</span>
                    <button onClick={deleteInvoice}
                      style={{ ...btnPri, background: T.danger, borderColor: T.danger, fontSize: 12 }}>
                      確認
                    </button>
                    <button onClick={() => setInvDeleteConfirm(false)} style={{ ...btnSec, fontSize: 12 }}>取消</button>
                  </div>
                )}
                <button onClick={() => setViewInvKey(null)} style={btnSec}>關閉</button>
              </div>
            </div>
          </div>
        )
      })()}

      {viewTxInvKey && (() => {
        const txGrpDetail = txInvoiceGroups[viewTxInvKey]
        if (!txGrpDetail) return null
        const txAmt = parseFloat(txInvPopupAmount) || null
        const txCheck = txAmt != null ? (Math.abs(txAmt - txGrpDetail.txFeeSum) < 0.01 ? '相符' : '有差異') : null
        const txCheckColor = txCheck === '相符' ? T.g700 : txCheck === '有差異' ? T.danger : T.n600
        const txStaticRows = [
          ['交易處理費發票號碼', viewTxInvKey, 'monospace'],
          ['交易處理費合計', `NT$ ${Math.round(txGrpDetail.txFeeSum * 100) / 100}`, 'inherit'],
          ['包含訂單', `${txGrpDetail.count} 筆`, 'inherit'],
          ['核對結果', txCheck || '—', 'inherit'],
        ]
        return (
          <div style={overlay} onClick={() => setViewTxInvKey(null)}>
            <div style={{ ...modal, width: 400 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>交易處理費發票資訊</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <tbody>
                  <tr style={mRow}>
                    <td style={mLabel}>交易處理費發票號碼</td>
                    <td style={{ ...mVal, fontFamily: 'monospace' }}>{viewTxInvKey}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>發票日期</td>
                    <td style={{ padding: '6px 0' }}>
                      <input type="date" value={txInvPopupDate} onChange={e => { setTxInvPopupDate(e.target.value); saveTxInvPopupDate(e.target.value) }}
                        style={popupInp} />
                    </td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>發票金額</td>
                    <td style={{ padding: '6px 0' }}>
                      <input type="number" value={txInvPopupAmount} onChange={e => setTxInvPopupAmount(e.target.value)}
                        onBlur={e => saveTxInvPopupAmount(e.target.value)} placeholder="0" style={popupInp} />
                    </td>
                  </tr>
                  {txStaticRows.map(([label, val, ff]) => (
                    <tr key={label} style={mRow}>
                      <td style={mLabel}>{label}</td>
                      <td style={{ ...mVal, fontWeight: label === '核對結果' ? 700 : 400, color: label === '核對結果' ? txCheckColor : T.text, fontFamily: ff }}>{val}</td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ ...mLabel, verticalAlign: 'top' }}>備注</td>
                    <td style={mVal}>
                      <textarea value={txInvNote} onChange={e => setTxInvNote(e.target.value)}
                        onBlur={e => saveTxInvNote(e.target.value)} placeholder="輸入備注…" rows={3} style={popupArea} />
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...mLabel, verticalAlign: 'top' }}>發票 PDF</td>
                    <td style={mVal}>
                      <input ref={txInvPdfRef} type="file" accept="application/pdf" style={{ display: 'none' }} onChange={uploadTxInvPdf} />
                      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                        <button onClick={() => txInvPdfRef.current.click()} style={{ ...btnSec, fontSize: 12 }} disabled={txInvUploading}>
                          {txInvUploading ? '上傳中…' : txInvPdfUrl ? '重新上傳' : '上傳 PDF'}
                        </button>
                        {txInvPdfUrl && (
                          <a href={txInvPdfUrl} target="_blank" rel="noopener noreferrer"
                            style={{ fontSize: 12, color: T.a700, textDecoration: 'underline' }}>查看 PDF</a>
                        )}
                      </div>
                      {txInvUploadError && <div style={{ fontSize: 12, color: T.danger, marginTop: 4 }}>{txInvUploadError}</div>}
                    </td>
                  </tr>
                </tbody>
              </table>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 20 }}>
                {!txInvDeleteConfirm ? (
                  <button onClick={() => setTxInvDeleteConfirm(true)}
                    style={{ ...btnSec, color: T.danger, borderColor: T.danger, fontSize: 12 }}>
                    刪除發票資訊
                  </button>
                ) : (
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <span style={{ fontSize: 12, color: T.danger }}>確定刪除？</span>
                    <button onClick={deleteTxInvoice}
                      style={{ ...btnPri, background: T.danger, borderColor: T.danger, fontSize: 12 }}>
                      確認
                    </button>
                    <button onClick={() => setTxInvDeleteConfirm(false)} style={{ ...btnSec, fontSize: 12 }}>取消</button>
                  </div>
                )}
                <button onClick={() => setViewTxInvKey(null)} style={btnSec}>關閉</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* 訂單發票資訊 modal */}
      {viewOrdInvKey && (() => {
        const grp = ordInvGroups[viewOrdInvKey]
        if (!grp) return null
        const saveOrdInvDate = async (date) => {
          await supabase.from('shipping_orders').update({ order_invoice_date: date || null }).eq('order_invoice_no', viewOrdInvKey)
          await loadOrders()
        }
        const deleteOrdInvoice = async () => {
          await supabase.from('shipping_orders')
            .update({ order_invoice_no: null, order_invoice_date: null, order_invoice_amount: null })
            .eq('order_invoice_no', viewOrdInvKey)
          setViewOrdInvKey(null)
          setOrdInvDeleteConfirm(false)
          await loadOrders()
        }
        return (
          <div style={overlay} onClick={() => setViewOrdInvKey(null)}>
            <div style={{ ...modal, width: 400 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>訂單發票資訊</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <tbody>
                  <tr style={mRow}>
                    <td style={mLabel}>訂單發票號碼</td>
                    <td style={{ ...mVal, fontFamily: 'monospace' }}>{viewOrdInvKey}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>發票日期</td>
                    <td style={{ padding: '6px 0' }}>
                      <input type="date" value={ordInvPopupDate}
                        onChange={e => setOrdInvPopupDate(e.target.value)}
                        onBlur={e => saveOrdInvDate(e.target.value)}
                        style={popupInp} />
                    </td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>應開發票金額</td>
                    <td style={mVal}>{(Math.round(grp.baseSum * 100) / 100).toLocaleString()}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>代開發票金額</td>
                    <td style={mVal}>{grp.invAmount != null ? grp.invAmount.toLocaleString() : '未填'}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>核對結果</td>
                    <td style={{ ...mVal, color: grp.invAmount == null ? T.n600 : Math.abs(grp.baseSum - grp.invAmount) < 0.01 ? T.g700 : T.danger }}>
                      {grp.invAmount == null ? '—'
                        : Math.abs(grp.baseSum - grp.invAmount) < 0.01 ? '✓ 相符'
                        : `✗ 差異 ${Math.round((grp.baseSum - grp.invAmount) * 100) / 100}`}
                    </td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>內扣手續費後入帳</td>
                    <td style={mVal}>{(Math.round(grp.payableSum * 100) / 100).toLocaleString()}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>包含訂單</td>
                    <td style={mVal}>{grp.count} 筆</td>
                  </tr>
                </tbody>
              </table>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 20 }}>
                {!ordInvDeleteConfirm ? (
                  <button onClick={() => setOrdInvDeleteConfirm(true)}
                    style={{ ...btnSec, color: T.danger, borderColor: T.danger, fontSize: 12 }}>
                    刪除發票資訊
                  </button>
                ) : (
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <span style={{ fontSize: 12, color: T.danger }}>確定刪除？</span>
                    <button onClick={deleteOrdInvoice}
                      style={{ ...btnPri, background: T.danger, borderColor: T.danger, fontSize: 12 }}>
                      確認
                    </button>
                    <button onClick={() => setOrdInvDeleteConfirm(false)} style={{ ...btnSec, fontSize: 12 }}>取消</button>
                  </div>
                )}
                <button onClick={() => setViewOrdInvKey(null)} style={btnSec}>關閉</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* 入帳群組細節 modal */}
      {viewBankGroupKey && showBankGroup && (() => {
        const bgOrders = orders.filter(o => o.recon_status === '已入帳' && (o.in_date || '').slice(0, 10) === viewBankGroupKey)
        const payableSum = Math.round(bgOrders.reduce((s, o) => s + (o.payable || 0), 0) * 100) / 100
        const bankDeposit = bgOrders[0]?.bank_deposit ?? null
        const diff = bankDeposit != null ? Math.round((bankDeposit - payableSum) * 100) / 100 : null
        const noteKey = `bankgroup_note_${gateway}_${viewBankGroupKey}`
        const saveBankGroupNote = (val) => {
          setBankGroupNote(val)
          localStorage.setItem(noteKey, val)
        }
        return (
          <div style={overlay} onClick={() => setViewBankGroupKey(null)}>
            <div style={{ ...modal, width: 440 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ marginTop: 0, fontSize: 20, color: T.navy }}>入帳細節</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <tbody>
                  <tr style={mRow}>
                    <td style={mLabel}>入帳日</td>
                    <td style={{ ...mVal, fontWeight: 700 }}>{viewBankGroupKey}</td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>玉山實際入帳</td>
                    <td style={{ ...mVal, fontWeight: 700 }}>
                      {bankDeposit != null ? `NT$ ${bankDeposit.toLocaleString()}` : '—'}
                    </td>
                  </tr>
                  <tr style={mRow}>
                    <td style={mLabel}>應入帳合計</td>
                    <td style={mVal}>NT$ {payableSum.toLocaleString()}</td>
                  </tr>
                  {diff != null && (
                    <tr style={mRow}>
                      <td style={mLabel}>差額</td>
                      <td style={{ ...mVal, fontWeight: 700, color: Math.abs(diff) < 0.01 ? T.g700 : T.danger }}>
                        {diff === 0 ? '相符' : `${diff > 0 ? '+' : ''}${diff}`}
                      </td>
                    </tr>
                  )}
                  <tr style={mRow}>
                    <td style={{ ...mLabel, verticalAlign: 'top' }}>備注</td>
                    <td style={mVal}>
                      <textarea value={bankGroupNote} onChange={e => saveBankGroupNote(e.target.value)}
                        placeholder="輸入備注（例如差額說明）…" rows={3} style={popupArea} />
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...mLabel, borderBottom: 'none' }}>包含訂單</td>
                    <td style={{ ...mVal, borderBottom: 'none' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, marginTop: 2 }}>
                        <thead>
                          <tr style={{ background: T.n200 }}>
                            <th style={subTh}>訂單編號</th>
                            <th style={{ ...subTh, textAlign: 'right' }}>應入帳</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bgOrders.map(o => (
                            <tr key={o.id} style={{ borderTop: `1px solid ${T.divider}` }}>
                              <td style={{ padding: '4px 6px', fontFamily: 'monospace', fontSize: 11 }}>{o.ref_no}</td>
                              <td style={{ padding: '4px 6px', textAlign: 'right' }}>{o.payable?.toLocaleString() ?? '—'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </td>
                  </tr>
                </tbody>
              </table>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 20 }}>
                <button onClick={() => setViewBankGroupKey(null)} style={btnSec}>關閉</button>
              </div>
            </div>
          </div>
        )
      })()}
    </div>
  )
}

// 對帳狀態 → tag 色調（已對帳=綠、已入帳=寶藍、其餘中性／描邊）
function statusTone(s) {
  if (s === '已對帳') return 'accent2'
  if (s === '已入帳') return 'accent'
  if (s === '已出貨') return 'outline'
  return 'neutral'
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
const th = { textAlign: 'left', padding: '8px 10px', borderBottom: `2px solid ${C.line}`, color: C.sub, fontWeight: 600, whiteSpace: 'nowrap', position: 'sticky', top: 0, background: '#fff', zIndex: 1 }
const td = { padding: '7px 10px', borderBottom: `1px solid ${C.line}`, whiteSpace: 'nowrap' }
const overlay = { position: 'fixed', inset: 0, background: 'rgba(22,35,63,.42)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, zIndex: 50 }
const modal = { background: T.surface, borderRadius: T.rPanel, padding: 26, width: 480, maxWidth: '100%', maxHeight: '90vh', overflowY: 'auto', boxShadow: T.shadowLg }

// ====== 金流對帳：品牌 token 樣式 ======
const btnPri = { padding: '8px 16px', borderRadius: T.rPill, border: `1px solid ${T.a}`, background: T.a, color: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit', whiteSpace: 'nowrap' }
const btnSec = { padding: '8px 16px', borderRadius: T.rPill, border: `1px solid ${T.divider}`, background: T.surface, color: T.n700, fontSize: 14, cursor: 'pointer', fontFamily: 'inherit', whiteSpace: 'nowrap' }
const btnGhostT = { padding: '5px 12px', borderRadius: T.rPill, border: 'none', background: 'transparent', color: T.a700, fontSize: 13, cursor: 'pointer', fontFamily: 'inherit', whiteSpace: 'nowrap' }
const inpT = { padding: '7px 12px', borderRadius: T.rPill, border: `1px solid ${T.divider}`, fontSize: 14, width: '100%', boxSizing: 'border-box', fontFamily: 'inherit', background: T.surface, color: T.text, outline: 'none' }
const thT = { textAlign: 'left', padding: '11px 14px', background: T.n200, color: T.n600, fontWeight: 600, fontSize: 11, letterSpacing: '.06em', whiteSpace: 'nowrap', position: 'sticky', top: 0, zIndex: 1 }
const thAnchor = { ...thT, background: T.n300, color: T.n800 }
const tdT = { padding: '11px 14px', borderBottom: `1px solid ${T.divider}`, whiteSpace: 'nowrap', fontSize: 13 }
const subTh = { padding: '4px 6px', textAlign: 'left', fontWeight: 500, color: T.n600, fontSize: 11, whiteSpace: 'nowrap' }
const panelLead = { fontSize: 14, color: T.n700, marginBottom: 12, lineHeight: 1.7 }
const pickWrap = { overflowX: 'auto', maxHeight: 260, overflowY: 'auto', margin: '8px 0', border: `1px solid ${T.divider}`, borderRadius: T.rInner, background: T.surface }
const popupInp = { fontSize: 13, padding: '5px 10px', borderRadius: T.rPill, border: `1px solid ${T.divider}`, outline: 'none', fontFamily: 'inherit', width: '100%', boxSizing: 'border-box' }
const popupArea = { width: '100%', fontSize: 13, padding: '8px 10px', borderRadius: T.rInner, border: `1px solid ${T.divider}`, resize: 'vertical', outline: 'none', fontFamily: 'inherit', boxSizing: 'border-box' }
const slotChipBtn = { borderRadius: T.rPill, border: 'none', background: 'rgba(22,35,63,.72)', color: '#fff', fontSize: 12, padding: '4px 10px', cursor: 'pointer', fontFamily: 'inherit', lineHeight: 1.5 }
const mRow = { borderBottom: `1px solid ${T.divider}` }
const mLabel = { padding: '9px 0', color: T.n600, width: 116, verticalAlign: 'top' }
const mVal = { padding: '9px 0' }

// 手續費發票群組色帶：同一張發票的訂單共用一組（底色 + 左直條），兩組交替
const INV_BG = [T.g100, T.a100]    // 沙綠 / 淡藍
const INV_BAR = [T.g500, T.a500]

const TAG_TONE = {
  accent:  { bg: T.a100, fg: T.a700 },
  accent2: { bg: T.g100, fg: T.g700 },
  neutral: { bg: T.n200, fg: T.n700 },
  danger:  { bg: T.dangerBg, fg: T.danger },
  outline: { bg: 'transparent', fg: T.n700, border: T.divider },
}
function Tag({ tone = 'neutral', children, style }) {
  const t = TAG_TONE[tone] || TAG_TONE.neutral
  return (
    <span style={{
      display: 'inline-block', padding: '3px 10px', borderRadius: T.rPill, fontSize: 12, fontWeight: 600,
      background: t.bg, color: t.fg, border: `1px solid ${t.border || 'transparent'}`, whiteSpace: 'nowrap', ...style,
    }}>{children}</span>
  )
}

// 上傳區塊：未上傳＝寶藍虛線；已上傳＝綠實底 + ✓
function DropButton({ onClick, filled, label, hint, block }) {
  return (
    <button onClick={onClick} style={{
      border: `1.5px dashed ${filled ? T.g500 : T.a}`,
      background: filled ? T.g100 : T.a100,
      borderRadius: T.rPanel, padding: 18, cursor: 'pointer',
      display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
      font: 'inherit', color: 'inherit', width: block ? '100%' : undefined,
      minWidth: block ? undefined : 260, flex: block ? undefined : '0 1 auto',
      transition: 'border-color .12s,background .12s',
    }}>
      <span style={{ width: 44, height: 44, flex: 'none', borderRadius: '50%',
        background: filled ? T.g600 : T.a, display: 'grid', placeItems: 'center', color: '#fff', fontSize: 18 }}>
        {filled ? '✓' : '↑'}
      </span>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 15, fontWeight: 700, overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</div>
        <div style={{ fontSize: 12, color: filled ? T.g700 : T.a700 }}>{hint}</div>
      </div>
    </button>
  )
}

// 「資料來源」小膠囊：就地說明這一步的檔案去哪抓
function SourceHint({ text, onSop }) {
  return (
    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 7, padding: '5px 12px',
      borderRadius: T.rPill, background: T.a100, fontSize: 12, color: T.a700, lineHeight: 1.7, flexWrap: 'wrap' }}>
      <span style={{ fontWeight: 700 }}>ⓘ</span>
      <span>資料來源：{text}</span>
      <button onClick={onSop} style={{ border: 'none', background: 'none', padding: 0, font: 'inherit',
        color: T.a700, textDecoration: 'underline', cursor: 'pointer' }}>完整 SOP</button>
    </div>
  )
}

function PanelMsg({ text, bad }) {
  if (!text) return null
  const isBad = bad ? bad.test(text) : false
  return <p style={{ marginTop: 10, marginBottom: 0, fontSize: 13, color: isBad ? T.danger : T.g700 }}>{text}</p>
}

// 方式 A（期間篩選）／方式 B（手動勾選）切換
function SegBtns({ value, onChange }) {
  return (
    <div style={{ display: 'inline-flex', margin: '14px 0', borderRadius: T.rPill, overflow: 'hidden', border: `1px solid ${T.divider}` }}>
      {[['auto', '方式 A — 期間篩選'], ['manual', '方式 B — 手動勾選']].map(([v, lbl]) => (
        <button key={v} onClick={() => onChange(v)} style={{
          padding: '7px 16px', border: 'none', cursor: 'pointer', fontSize: 13, fontFamily: 'inherit',
          background: value === v ? T.a : T.surface, color: value === v ? '#fff' : T.n700, whiteSpace: 'nowrap',
        }}>{lbl}</button>
      ))}
    </div>
  )
}

function SumBar({ tone, children }) {
  const bg = tone === 'ok' ? T.g100 : tone === 'bad' ? T.dangerBg : T.n200
  return (
    <div style={{ padding: '14px 18px', borderRadius: T.rInner, marginTop: 12, background: bg,
      display: 'flex', gap: 20, alignItems: 'center', flexWrap: 'wrap' }}>{children}</div>
  )
}

// 已確認入帳群組（依 in_date 彙總）
// onRelease(ids, label)：把整組或單筆訂單退回未入帳，供修正勾錯的歸戶
function ConfirmedGroups({ groups, exp, setExp, onRelease, releasingKey }) {
  return (
    <div style={{ marginTop: 18 }}>
      <p style={{ fontSize: 11, letterSpacing: '.06em', color: T.n600, margin: '0 0 8px', fontWeight: 600 }}>已確認入帳</p>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {groups.map(g => {
          const open = !!exp[g.date]
          return (
            <div key={g.date} style={{ border: `1.5px solid ${T.g500}`, borderRadius: T.rPanel, padding: '14px 18px', background: T.g100 }}>
              <div style={{ display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
                <span style={{ fontWeight: 700, fontSize: 15 }}>入帳日 {g.date}</span>
                <span style={{ fontSize: 14 }}>應入帳合計 <strong style={{ color: T.g700 }}>NT$ {(Math.round(g.payable * 100) / 100).toLocaleString()}</strong></span>
                <Tag tone="accent2">{g.orders.length} 筆</Tag>
                <button onClick={() => setExp(p => ({ ...p, [g.date]: !open }))}
                  style={{ ...btnSec, fontSize: 12, padding: '4px 12px', marginLeft: 'auto' }}>
                  {open ? '收起 ▲' : '展開 ▼'}
                </button>
                {onRelease && (
                  <button
                    disabled={releasingKey != null}
                    onClick={() => onRelease(g.orders.map(o => o.id), `入帳日 ${g.date} 這組（${g.orders.length} 筆）`)}
                    style={{ ...btnSec, fontSize: 12, padding: '4px 12px', color: T.danger, borderColor: T.danger,
                      opacity: releasingKey != null ? .5 : 1, cursor: releasingKey != null ? 'default' : 'pointer' }}>
                    解除整組
                  </button>
                )}
              </div>
              {open && (
                <div style={{ marginTop: 10, borderTop: `1px solid ${T.g300}`, paddingTop: 10 }}>
                  <table style={{ width: '100%', fontSize: 12, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={subTh}>平台訂單編號</th>
                        <th style={subTh}>訂單日期</th>
                        <th style={{ ...subTh, textAlign: 'right' }}>應入帳</th>
                        {onRelease && <th style={{ ...subTh, textAlign: 'right' }} />}
                      </tr>
                    </thead>
                    <tbody>
                      {g.orders.slice().sort((a, b) => (a.order_date || '').localeCompare(b.order_date || '')).map(o => (
                        <tr key={o.id} style={{ borderBottom: `1px solid ${T.g200}` }}>
                          <td style={{ padding: '4px 6px', fontFamily: 'monospace' }}>{o.ref_no}</td>
                          <td style={{ padding: '4px 6px' }}>{o.order_date}</td>
                          <td style={{ padding: '4px 6px', textAlign: 'right' }}>{o.payable?.toLocaleString()}</td>
                          {onRelease && (
                            <td style={{ padding: '4px 6px', textAlign: 'right' }}>
                              <button
                                disabled={releasingKey != null}
                                onClick={() => onRelease([o.id], `訂單 ${o.ref_no}`)}
                                style={{ background: 'none', border: 'none', padding: '2px 4px', fontSize: 12,
                                  color: T.danger, textDecoration: 'underline', fontFamily: 'inherit',
                                  opacity: releasingKey != null ? .5 : 1,
                                  cursor: releasingKey != null ? 'default' : 'pointer' }}>
                                移除
                              </button>
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}

// 群組型發票卡片：一張發票涵蓋哪幾筆訂單、費用合計 vs 發票金額
function InvoiceGroupCards({ groups, colorIdx, kind, orders, onOpen, storeKey }) {
  const keys = Object.keys(groups)
  const kindName = kind === 'tx' ? '交易處理費發票' : kind === 'ord' ? '代開發票' : '手續費發票'
  const field = kind === 'tx' ? 'tx_fee_invoice_no' : kind === 'ord' ? 'order_invoice_no' : 'fee_invoice_no'
  // 收合狀態記在 localStorage：發票張數一多這區塊會很長，把下方輸入欄位擠出畫面。
  // 依「通路＋發票種類」各記各的，切換頁籤不會被別的頁籤蓋掉。
  const lsKey = `invcards_open_${storeKey || ''}_${kind}`
  const [open, setOpen] = useState(() => localStorage.getItem(lsKey) !== '0')
  useEffect(() => { setOpen(localStorage.getItem(lsKey) !== '0') }, [lsKey])
  function toggle() {
    setOpen(v => { localStorage.setItem(lsKey, v ? '0' : '1'); return !v })
  }

  if (!keys.length) {
    return (
      <div style={{ border: `1.5px dashed ${T.divider}`, background: T.bg, borderRadius: T.rPanel,
        padding: '16px 18px', fontSize: 13, color: T.n600 }}>
        尚未開立{kind === 'tx' ? '交易處理費' : kind === 'ord' ? '代開' : '手續費'}發票 — 用下方欄位輸入發票號碼、勾選訂單即可群組成一張發票。
      </div>
    )
  }

  const cards = keys.map((k, i) => {
    const g = groups[k]
    const sum = Math.round((g.feeSum ?? g.txFeeSum ?? g.baseSum ?? 0) * 100) / 100
    const amt = g.invAmount
    return { k, g, sum, amt, ok: amt != null ? Math.abs(sum - amt) < 0.01 : null, i }
  })
  const badCount = cards.filter(c => c.ok === false).length
  const pendingCount = cards.filter(c => c.ok == null).length

  return (
    <div>
      <div onClick={toggle} style={{
        display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', userSelect: 'none',
        padding: '8px 2px', marginBottom: open ? 10 : 0,
      }}>
        <span style={{ fontSize: 11, color: T.n600, transform: open ? 'none' : 'rotate(-90deg)', transition: 'transform .15s' }}>▼</span>
        <span style={{ fontSize: 13, fontWeight: 600, color: T.n800 }}>已開立{kindName} · {keys.length} 張</span>
        {badCount > 0 && <Tag tone="danger" style={{ fontSize: 11 }}>{badCount} 張差異</Tag>}
        {pendingCount > 0 && <Tag tone="neutral" style={{ fontSize: 11, fontWeight: 400 }}>{pendingCount} 張未填金額</Tag>}
        <span style={{ marginLeft: 'auto', fontSize: 12, color: T.a700 }}>{open ? '收合' : '展開'}</span>
      </div>
      {open && (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(280px,1fr))', gap: 14 }}>
      {cards.map(({ k, g, sum, amt, ok, i }) => {
        const ci = colorIdx[k] != null ? colorIdx[k] : i % 2
        const refs = orders.filter(o => o[field] === k).map(o => String(o.ref_no || ''))
        const border = ok === false ? T.danger : ok === true ? T.g500 : T.divider
        const bg = ok === false ? T.dangerBg : ok === true ? T.g100 : T.bg
        return (
          <div key={k} onClick={() => onOpen(k)} style={{
            border: `1.5px solid ${border}`, borderLeft: `6px solid ${ok === false ? T.danger : INV_BAR[ci]}`,
            background: bg, borderRadius: T.rPanel, padding: '16px 18px', cursor: 'pointer',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
              <div style={{ fontSize: 11, letterSpacing: '.06em', color: ok === false ? T.danger : T.n600 }}>
                {kind === 'tx' ? '交易處理費發票' : kind === 'ord' ? '代開發票' : '手續費發票'} · 共 {g.count} 筆
              </div>
              {ok != null && (
                <Tag tone={ok ? 'accent2' : 'danger'} style={{ fontSize: 11 }}>
                  {ok ? '✓ 相符' : `✗ 差異 ${Math.round((sum - amt) * 100) / 100}`}
                </Tag>
              )}
            </div>
            <div style={{ fontSize: 17, fontWeight: 700, marginTop: 4, fontFamily: 'monospace' }}>{k}</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5, marginTop: 10 }}>
              {refs.slice(0, 6).map(r => (
                <Tag key={r} tone="neutral" style={{ fontFamily: 'monospace', fontSize: 11, fontWeight: 400 }}>
                  …{r.slice(-4)}
                </Tag>
              ))}
              {refs.length > 6 && <Tag tone="neutral" style={{ fontSize: 11, fontWeight: 400 }}>+{refs.length - 6}</Tag>}
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12, fontSize: 13 }}>
              <span style={{ color: T.n700 }}>{kind === 'ord' ? '應開發票金額' : '費用合計'}</span><strong>NT$ {sum.toLocaleString()}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 4, fontSize: 13 }}>
              <span style={{ color: T.n700 }}>發票金額</span>
              <strong style={{ color: ok === false ? T.danger : T.text }}>
                {amt != null ? `NT$ ${amt.toLocaleString()}` : '未填'}
              </strong>
            </div>
            {kind === 'ord' && (
              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 4, fontSize: 12, color: T.n600 }}>
                <span>內扣手續費後入帳</span>
                <span>NT$ {(Math.round((g.payableSum || 0) * 100) / 100).toLocaleString()}</span>
              </div>
            )}
          </div>
        )
      })}
    </div>
      )}
    </div>
  )
}

function IconBook() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2.75" strokeLinecap="round" strokeLinejoin="round">
      <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
      <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
    </svg>
  )
}
