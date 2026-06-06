import { useState, useMemo, useCallback, useEffect } from 'react'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts'
import {
  TrendingDown, Package, Upload, Download,
  Search, Edit2, Check, X, FlaskConical, Clock,
  HardDriveDownload, Loader2, Wifi, WifiOff, PackagePlus,
} from 'lucide-react'
import XLSX from 'xlsx-js-style'
import {
  supabase,
  toAppReagent, toAppLog, toAppInbound,
  toDbReagent, toDbLog, toDbInbound,
} from './lib/supabase'
import { generateMonthlyReport } from './lib/exportReport'

// ── 날짜 유틸 ──────────────────────────────────────────────────────────────
const fmtDate = (d) => {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

const parseExcelDate = (val) => {
  if (!val) return ''
  if (val instanceof Date) return fmtDate(val)
  if (typeof val === 'string') {
    if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 10)
    return val
  }
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000))
    return fmtDate(d)
  }
  return String(val)
}

const fmtDateTime = (d) => {
  const h = String(d.getHours()).padStart(2, '0')
  const min = String(d.getMinutes()).padStart(2, '0')
  return `${fmtDate(d)} ${h}:${min}`
}

const diffDays = (dateStr) => {
  const exp = new Date(dateStr)
  const now = new Date()
  now.setHours(0, 0, 0, 0)
  exp.setHours(0, 0, 0, 0)
  return Math.ceil((exp - now) / (1000 * 60 * 60 * 24))
}

// ══════════════════════════════════════════════════════════════════════════
// App (root)
// ══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [reagents, setReagents] = useState([])
  const [logs, setLogs] = useState([])
  const [inboundLogs, setInboundLogs] = useState([])
  const [loading, setLoading] = useState(true)
  const [syncing, setSyncing] = useState(false)
  const [online, setOnline] = useState(true)
  const [showInboundModal, setShowInboundModal] = useState(false)

  const [activeTab, setActiveTab] = useState('dashboard')
  const [search, setSearch] = useState('')
  const [activeFilter, setActiveFilter] = useState(null)
  const [editingId, setEditingId] = useState(null)
  const [editValue, setEditValue] = useState('')

  // ── 데이터 로드 + 실시간 구독 ──────────────────────────────────────────
  const loadReagents = useCallback(async () => {
    const { data, error } = await supabase
      .from('reagents').select('*, created_at').order('id')
    if (error) { console.error(error); return }
    setReagents(data.map(toAppReagent))
  }, [])

  const loadLogs = useCallback(async () => {
    const { data, error } = await supabase
      .from('dispatch_logs').select('*').order('datetime', { ascending: false })
    if (error) { console.error(error); return }
    setLogs(data.map(toAppLog))
  }, [])

  const loadInboundLogs = useCallback(async () => {
    const { data, error } = await supabase
      .from('inbound_logs').select('*').order('datetime', { ascending: false })
    if (error) { console.error(error); return }
    setInboundLogs(data.map(toAppInbound))
  }, [])

  useEffect(() => {
    ;(async () => {
      setLoading(true)
      await Promise.all([loadReagents(), loadLogs(), loadInboundLogs()])
      setLoading(false)
    })()

    const reagentChannel = supabase
      .channel('reagents-rt')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'reagents' }, loadReagents)
      .subscribe((status) => setOnline(status === 'SUBSCRIBED'))

    const logChannel = supabase
      .channel('logs-rt')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'dispatch_logs' }, loadLogs)
      .subscribe()

    const inboundChannel = supabase
      .channel('inbound-rt')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'inbound_logs' }, loadInboundLogs)
      .subscribe()

    return () => {
      supabase.removeChannel(reagentChannel)
      supabase.removeChannel(logChannel)
      supabase.removeChannel(inboundChannel)
    }
  }, [loadReagents, loadLogs, loadInboundLogs])

  // ── 알림 집계 ──────────────────────────────────────────────────────────
  const expiring = useMemo(
    () => reagents.filter((r) => { const d = diffDays(r.expiryDate); return d >= 0 && d < 15 }),
    [reagents],
  )
  const lowStock = useMemo(
    () => reagents.filter((r) => r.currentStock < r.minStock),
    [reagents],
  )

  // ── 최근 7일 차트 데이터 ───────────────────────────────────────────────
  const chartData = useMemo(() => {
    const now = new Date()
    return Array.from({ length: 7 }, (_, i) => {
      const d = new Date(now)
      d.setDate(d.getDate() - (6 - i))
      const dateKey = fmtDate(d)
      const total = logs
        .filter((l) => l.datetime.startsWith(dateKey))
        .reduce((s, l) => s + l.qty, 0)
      return { date: `${d.getMonth() + 1}/${d.getDate()}`, 출고량: total }
    })
  }, [logs])

  // ── 테이블 필터링 ──────────────────────────────────────────────────────
  const filteredReagents = useMemo(() => {
    let list = reagents
    if (activeFilter === 'expiring') list = expiring
    else if (activeFilter === 'lowStock') list = lowStock
    if (search.trim()) {
      const q = search.trim().toLowerCase()
      list = list.filter(
        (r) =>
          r.name.toLowerCase().includes(q) ||
          r.lotNo.toLowerCase().includes(q) ||
          r.manufacturer.toLowerCase().includes(q),
      )
    }
    return list
  }, [reagents, activeFilter, search, expiring, lowStock])

  // ── 출고 ───────────────────────────────────────────────────────────────
  const handleDispatch = useCallback(async (r) => {
    if (r.currentStock <= 0) { alert(`${r.name} 재고가 없습니다.`); return }
    const now = new Date()
    const newStock = r.currentStock - 1
    const newDispatched = r.totalDispatched + 1
    const logEntry = {
      id: Date.now(),
      reagentId: r.id,
      reagentName: r.name,
      lotNo: r.lotNo,
      qty: 1,
      datetime: fmtDateTime(now),
    }

    setReagents((prev) =>
      prev.map((x) =>
        x.id === r.id ? { ...x, currentStock: newStock, totalDispatched: newDispatched } : x,
      ),
    )
    setLogs((prev) => [logEntry, ...prev])

    setSyncing(true)
    const [{ error: e1 }, { error: e2 }] = await Promise.all([
      supabase.from('reagents').update({
        current_stock: newStock,
        total_dispatched: newDispatched,
        updated_at: now.toISOString(),
      }).eq('id', r.id),
      supabase.from('dispatch_logs').insert(toDbLog(logEntry)),
    ])
    setSyncing(false)
    if (e1 || e2) {
      alert('저장 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.')
      await Promise.all([loadReagents(), loadLogs()])
    }
  }, [loadReagents, loadLogs])

  // ── 입고 ───────────────────────────────────────────────────────────────
  const handleInbound = useCallback(async ({ mode, form }) => {
    const now = new Date()
    const datetime = fmtDateTime(now)
    setSyncing(true)

    if (mode === 'new') {
      const newId = Math.max(0, ...reagents.map((r) => r.id)) + 1
      const newReagent = {
        id: newId,
        name: form.name.trim(),
        reagentType: form.reagentType || 'Reagent',
        manufacturer: form.manufacturer.trim(),
        lotNo: form.lotNo.trim(),
        receivedQty: Number(form.qty),
        currentStock: Number(form.qty),
        minStock: Number(form.minStock) || 0,
        expiryDate: form.expiryDate,
        totalDispatched: 0,
        createdAt: now.toISOString(),
      }
      const logEntry = {
        id: Date.now(),
        reagentId: newId,
        reagentName: newReagent.name,
        lotNo: newReagent.lotNo,
        qty: Number(form.qty),
        datetime,
        notes: form.notes || '',
      }

      setReagents((prev) => [...prev, newReagent])
      setInboundLogs((prev) => [logEntry, ...prev])

      const [{ error: e1 }, { error: e2 }] = await Promise.all([
        supabase.from('reagents').insert(toDbReagent(newReagent)),
        supabase.from('inbound_logs').insert(toDbInbound(logEntry)),
      ])
      setSyncing(false)
      if (e1 || e2) {
        alert('저장 오류가 발생했습니다.')
        await Promise.all([loadReagents(), loadInboundLogs()])
        return
      }
    } else {
      const target = reagents.find((r) => r.id === Number(form.reagentId))
      if (!target) { setSyncing(false); return }

      const addQty = Number(form.qty)
      const newStock = target.currentStock + addQty
      const newReceived = target.receivedQty + addQty
      const updatedLotNo = form.newLotNo.trim() || target.lotNo
      const logEntry = {
        id: Date.now(),
        reagentId: target.id,
        reagentName: target.name,
        lotNo: updatedLotNo,
        qty: addQty,
        datetime,
        notes: form.notes || '',
      }

      setReagents((prev) =>
        prev.map((r) =>
          r.id === target.id
            ? { ...r, currentStock: newStock, receivedQty: newReceived, lotNo: updatedLotNo }
            : r,
        ),
      )
      setInboundLogs((prev) => [logEntry, ...prev])

      const dbUpdate = {
        current_stock: newStock,
        received_qty: newReceived,
        updated_at: now.toISOString(),
      }
      if (form.newLotNo.trim()) dbUpdate.lot_no = form.newLotNo.trim()

      const [{ error: e1 }, { error: e2 }] = await Promise.all([
        supabase.from('reagents').update(dbUpdate).eq('id', target.id),
        supabase.from('inbound_logs').insert(toDbInbound(logEntry)),
      ])
      setSyncing(false)
      if (e1 || e2) {
        alert('저장 오류가 발생했습니다.')
        await Promise.all([loadReagents(), loadInboundLogs()])
        return
      }
    }

    setShowInboundModal(false)
  }, [reagents, loadReagents, loadInboundLogs])

  // ── 최소 재고 편집 ─────────────────────────────────────────────────────
  const startEdit = useCallback((r) => { setEditingId(r.id); setEditValue(String(r.minStock)) }, [])
  const confirmEdit = useCallback(async (id) => {
    const val = parseInt(editValue, 10)
    if (!isNaN(val) && val >= 0) {
      setReagents((prev) => prev.map((r) => (r.id === id ? { ...r, minStock: val } : r)))
      await supabase.from('reagents').update({ min_stock: val }).eq('id', id)
    }
    setEditingId(null)
  }, [editValue])
  const cancelEdit = useCallback(() => setEditingId(null), [])

  // ── 엑셀 업로드 ────────────────────────────────────────────────────────
  const handleUpload = useCallback(async (e) => {
    const file = e.target.files[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true })
        const isBackup = wb.SheetNames.includes('재고')

        setSyncing(true)
        if (isBackup) {
          const reagentRows = XLSX.utils.sheet_to_json(wb.Sheets['재고'])
          const logRows = wb.Sheets['출고이력_전체']
            ? XLSX.utils.sheet_to_json(wb.Sheets['출고이력_전체'])
            : []

          const restoredReagents = reagentRows.map((row) => ({
            id: Number(row['id']),
            name: String(row['시약명'] ?? ''),
            manufacturer: String(row['제조사'] ?? ''),
            lotNo: String(row['Lot No'] ?? ''),
            receivedQty: Number(row['입고량'] ?? 0),
            currentStock: Number(row['현재재고'] ?? 0),
            minStock: Number(row['최소유지재고'] ?? 0),
            expiryDate: parseExcelDate(row['유효기간']),
            totalDispatched: Number(row['누적출고량'] ?? 0),
          }))
          const restoredLogs = logRows.map((row) => ({
            id: Number(row['id']),
            reagentId: Number(row['reagentId'] ?? 0),
            reagentName: String(row['시약명'] ?? ''),
            lotNo: String(row['Lot No'] ?? ''),
            qty: Number(row['출고수량'] ?? 1),
            datetime: String(row['출고일시'] ?? ''),
          }))

          await supabase.from('dispatch_logs').delete().neq('id', 0)
          await supabase.from('reagents').delete().neq('id', 0)
          if (restoredReagents.length > 0)
            await supabase.from('reagents').insert(restoredReagents.map(toDbReagent))
          if (restoredLogs.length > 0)
            await supabase.from('dispatch_logs').insert(restoredLogs.map(toDbLog))

          setReagents(restoredReagents)
          setLogs(restoredLogs)
          alert(`✅ 백업 복구 완료\n시약 ${restoredReagents.length}종 · 출고이력 ${restoredLogs.length}건`)
        } else {
          const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
          const mapped = data.map((row, idx) => ({
            id: idx + 1,
            name: row['시약명'] ?? row['name'] ?? '',
            manufacturer: row['제조사'] ?? row['manufacturer'] ?? '',
            lotNo: row['Lot No'] ?? row['lotNo'] ?? '',
            receivedQty: Number(row['입고량'] ?? row['receivedQty'] ?? 0),
            currentStock: Number(row['현재재고'] ?? row['currentStock'] ?? 0),
            minStock: Number(row['최소유지재고'] ?? row['minStock'] ?? 0),
            expiryDate: parseExcelDate(row['유효기간'] ?? row['expiryDate']),
            totalDispatched: Number(row['누적출고량'] ?? row['totalDispatched'] ?? 0),
          }))
          await supabase.from('reagents').delete().neq('id', 0)
          await supabase.from('reagents').insert(mapped.map(toDbReagent))
          setReagents(mapped)
        }
        setSyncing(false)
      } catch (err) {
        setSyncing(false)
        console.error(err)
        alert('파일을 읽는 중 오류가 발생했습니다.')
      }
    }
    reader.readAsArrayBuffer(file)
    e.target.value = ''
  }, [])

  // ── 전체 상태 백업 ─────────────────────────────────────────────────────
  const handleBackup = useCallback(() => {
    const now = new Date()
    const stamp = `${fmtDate(now)}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`
    const reagentRows = reagents.map((r) => ({
      'id': r.id, '시약명': r.name, '제조사': r.manufacturer, 'Lot No': r.lotNo,
      '입고량': r.receivedQty, '현재재고': r.currentStock,
      '최소유지재고': r.minStock, '유효기간': r.expiryDate, '누적출고량': r.totalDispatched,
    }))
    const logRows = logs.map((l) => ({
      'id': l.id, 'reagentId': l.reagentId, '시약명': l.reagentName,
      'Lot No': l.lotNo, '출고수량': l.qty, '출고일시': l.datetime,
    }))
    const wsR = XLSX.utils.json_to_sheet(reagentRows)
    wsR['!cols'] = [{ wch: 6 }, { wch: 32 }, { wch: 14 }, { wch: 16 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 10 }]
    const wsL = XLSX.utils.json_to_sheet(logRows)
    wsL['!cols'] = [{ wch: 14 }, { wch: 8 }, { wch: 32 }, { wch: 16 }, { wch: 8 }, { wch: 18 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, wsR, '재고')
    XLSX.utils.book_append_sheet(wb, wsL, '출고이력_전체')
    XLSX.writeFile(wb, `재고백업_${stamp}.xlsx`)
  }, [reagents, logs])

  // ── 입고 엑셀 템플릿 다운로드 ────────────────────────────────────────
  const downloadInboundTemplate = useCallback(() => {
    const ws = XLSX.utils.json_to_sheet([
      { 시약명: 'CBC 희석액', 구분: 'Reagent', 제조사: 'Sysmex', 'Lot No': 'SX2025-001', 입고량: 10, 최소유지재고: 5, 유효기간: '2026-12-31', 비고: '' },
    ])
    ws['!cols'] = [{ wch: 28 }, { wch: 10 }, { wch: 14 }, { wch: 16 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 20 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, '입고등록')
    XLSX.writeFile(wb, '입고등록_템플릿.xlsx')
  }, [])

  // ── 엑셀 일괄 입고 ────────────────────────────────────────────────────
  const handleBulkInbound = useCallback(async (rows) => {
    if (rows.length === 0) return
    const now = new Date()
    const datetime = fmtDateTime(now)
    setSyncing(true)

    // Lot No 기준으로 기존 시약 맵 생성
    const lotMap = new Map(reagents.map((r) => [r.lotNo.trim(), r]))
    let nextId = Math.max(0, ...reagents.map((r) => r.id)) + 1

    const newReagentsList = []
    const updatesList = []      // { id, newStock, newReceived }
    const inboundEntries = []

    rows.forEach((row, idx) => {
      const lotNo = String(row['Lot No'] ?? row['Lot No.'] ?? '').trim()
      const qty   = Number(row['입고량'] ?? 0)
      const name  = String(row['시약명'] ?? '').trim()
      if (!name || qty < 1) return

      const logBase = {
        id: Date.now() + idx,
        qty,
        datetime,
        notes: String(row['비고'] ?? ''),
      }

      if (lotMap.has(lotNo)) {
        const existing = lotMap.get(lotNo)
        const newStock    = existing.currentStock + qty
        const newReceived = existing.receivedQty   + qty
        updatesList.push({ id: existing.id, newStock, newReceived })
        inboundEntries.push({ ...logBase, reagentId: existing.id, reagentName: existing.name, lotNo: existing.lotNo })
        // 맵 업데이트 (같은 Lot 여러 행 처리)
        lotMap.set(lotNo, { ...existing, currentStock: newStock, receivedQty: newReceived })
      } else {
        const id = nextId++
        const rawType = String(row['구분'] ?? '').trim()
        const typeNorm = rawType.charAt(0).toUpperCase() + rawType.slice(1).toLowerCase()
        const reagentType = ['Reagent', 'Cal', 'Con'].includes(typeNorm) ? typeNorm : 'Reagent'
        const newR = {
          id,
          name,
          reagentType,
          manufacturer: String(row['제조사'] ?? '').trim(),
          lotNo,
          receivedQty: qty,
          currentStock: qty,
          minStock: Number(row['최소유지재고'] ?? 0),
          expiryDate: parseExcelDate(row['유효기간']),
          totalDispatched: 0,
          createdAt: now.toISOString(),
        }
        newReagentsList.push(newR)
        lotMap.set(lotNo, newR)
        inboundEntries.push({ ...logBase, reagentId: id, reagentName: name, lotNo })
      }
    })

    // DB 저장
    const ops = []
    if (newReagentsList.length > 0)
      ops.push(supabase.from('reagents').insert(newReagentsList.map(toDbReagent)))
    for (const u of updatesList)
      ops.push(supabase.from('reagents').update({ current_stock: u.newStock, received_qty: u.newReceived, updated_at: now.toISOString() }).eq('id', u.id))
    if (inboundEntries.length > 0)
      ops.push(supabase.from('inbound_logs').insert(inboundEntries.map(toDbInbound)))

    const results = await Promise.all(ops)
    const hasError = results.some((r) => r.error)
    setSyncing(false)

    if (hasError) {
      alert('일부 저장 중 오류가 발생했습니다.')
      await Promise.all([loadReagents(), loadInboundLogs()])
    } else {
      await Promise.all([loadReagents(), loadInboundLogs()])
      alert(`✅ 일괄 입고 완료\n신규 ${newReagentsList.length}종 · 재고 추가 ${updatesList.length}종 · 총 ${inboundEntries.length}건`)
      setShowInboundModal(false)
    }
  }, [reagents, loadReagents, loadInboundLogs])

  // ── 월간 입출고 대장 ───────────────────────────────────────────────────
  const handleExport = useCallback(() => {
    const now = new Date()
    const ym = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
    const monthly = logs.filter((l) => l.datetime.startsWith(ym))
    if (monthly.length === 0) { alert('이번 달 출고 이력이 없습니다.'); return }
    generateMonthlyReport(logs, reagents, now.getFullYear(), now.getMonth() + 1, inboundLogs)
  }, [logs, reagents, inboundLogs])

  const toggleFilter = useCallback(
    (type) => setActiveFilter((prev) => (prev === type ? null : type)), [],
  )

  // ── 로딩 화면 ──────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-center space-y-3">
          <Loader2 size={36} className="text-blue-500 animate-spin mx-auto" />
          <p className="text-slate-500 text-sm">데이터를 불러오는 중...</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-slate-50">
      {/* ── Header ── */}
      <header className="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6">
          <div className="flex items-center justify-between h-14 gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <div className="bg-blue-600 p-1.5 rounded-lg shrink-0">
                <FlaskConical size={20} className="text-white" />
              </div>
              <div className="min-w-0">
                <h1 className="text-base font-bold text-slate-800 leading-tight truncate">
                  시약 재고 관리 시스템
                </h1>
                <p className="text-[11px] text-slate-400">진단검사의학과 · Reagent Inventory</p>
              </div>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              {/* 동기화 상태 */}
              <div className="flex items-center gap-1 text-xs mr-1">
                {syncing ? (
                  <><Loader2 size={12} className="text-blue-400 animate-spin" /><span className="text-blue-400">저장 중</span></>
                ) : online ? (
                  <><Wifi size={12} className="text-emerald-500" /><span className="text-emerald-500">연결됨</span></>
                ) : (
                  <><WifiOff size={12} className="text-slate-400" /><span className="text-slate-400">연결 끊김</span></>
                )}
              </div>
              <button
                onClick={() => setShowInboundModal(true)}
                className="flex items-center gap-1.5 bg-teal-600 hover:bg-teal-700 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors whitespace-nowrap"
                title="시약 입고 등록"
              >
                <PackagePlus size={14} />
                입고 등록
              </button>
              <button
                onClick={handleBackup}
                className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors whitespace-nowrap"
                title="재고 + 출고이력 전체 백업"
              >
                <HardDriveDownload size={14} />
                현재 재고상태 백업
              </button>
              <label className="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium px-3 py-1.5 rounded-lg cursor-pointer transition-colors whitespace-nowrap">
                <Upload size={14} />
                엑셀 업로드
                <input type="file" accept=".xlsx,.csv" className="hidden" onChange={handleUpload} />
              </label>
            </div>
          </div>
          {/* Tabs */}
          <div className="flex gap-0.5">
            {[
              { key: 'dashboard', label: '대시보드' },
              { key: 'history',   label: '출고 이력' },
              { key: 'inbound',   label: '입고 이력' },
            ].map(({ key, label }) => (
              <button
                key={key}
                onClick={() => setActiveTab(key)}
                className={`px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
                  activeTab === key
                    ? 'border-blue-600 text-blue-600'
                    : 'border-transparent text-slate-500 hover:text-slate-700'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        </div>
      </header>

      {/* ── Main ── */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-5">
        {activeTab === 'dashboard' ? (
          <DashboardTab
            reagents={filteredReagents}
            expiring={expiring}
            lowStock={lowStock}
            chartData={chartData}
            activeFilter={activeFilter}
            search={search}
            editingId={editingId}
            editValue={editValue}
            onToggleFilter={toggleFilter}
            onClearFilter={() => setActiveFilter(null)}
            onSearch={setSearch}
            onDispatch={handleDispatch}
            onStartEdit={startEdit}
            onConfirmEdit={confirmEdit}
            onCancelEdit={cancelEdit}
            onEditValue={setEditValue}
          />
        ) : activeTab === 'history' ? (
          <HistoryTab logs={logs} onExport={handleExport} />
        ) : (
          <InboundHistoryTab inboundLogs={inboundLogs} />
        )}
      </main>

      {showInboundModal && (
        <InboundModal
          reagents={reagents}
          syncing={syncing}
          onClose={() => setShowInboundModal(false)}
          onSubmit={handleInbound}
          onBulkInbound={handleBulkInbound}
          onDownloadTemplate={downloadInboundTemplate}
        />
      )}
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// DashboardTab
// ══════════════════════════════════════════════════════════════════════════
function DashboardTab({
  reagents, expiring, lowStock, chartData,
  activeFilter, search, editingId, editValue,
  onToggleFilter, onClearFilter, onSearch,
  onDispatch, onStartEdit, onConfirmEdit, onCancelEdit, onEditValue,
}) {
  return (
    <>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <AlertCard
          icon={<Clock size={20} />} color="amber"
          title="유효기간 임박 시약" subtitle="유효기간 15일 미만"
          count={expiring.length} active={activeFilter === 'expiring'}
          onClick={() => onToggleFilter('expiring')}
        />
        <AlertCard
          icon={<TrendingDown size={20} />} color="red"
          title="재고 부족 시약" subtitle="최소 유지 재고 미만"
          count={lowStock.length} active={activeFilter === 'lowStock'}
          onClick={() => onToggleFilter('lowStock')}
        />
      </div>

      <div className="bg-white rounded-xl border border-slate-100 shadow-sm p-5">
        <div className="flex items-center gap-2 mb-4">
          <Package size={15} className="text-blue-500" />
          <h2 className="text-sm font-semibold text-slate-600">최근 7일 일별 출고량 추이</h2>
        </div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={chartData} margin={{ top: 4, right: 8, bottom: 0, left: -20 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
            <XAxis dataKey="date" tick={{ fontSize: 11, fill: '#94a3b8' }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: '#94a3b8' }} axisLine={false} tickLine={false} allowDecimals={false} />
            <Tooltip
              contentStyle={{ fontSize: 12, borderRadius: 8, border: 'none', boxShadow: '0 4px 16px rgba(0,0,0,0.1)' }}
              formatter={(v) => [`${v}개`, '출고량']}
              cursor={{ fill: '#eff6ff' }}
            />
            <Bar dataKey="출고량" fill="#3b82f6" radius={[4, 4, 0, 0]} maxBarSize={40} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="flex flex-col sm:flex-row gap-3 items-start sm:items-center">
        <div className="relative flex-1 w-full">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={search}
            onChange={(e) => onSearch(e.target.value)}
            placeholder="시약명, Lot No, 제조사로 검색..."
            className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 bg-white"
          />
        </div>
        {activeFilter && (
          <div className="flex items-center gap-2 shrink-0">
            <span className="text-xs text-slate-500">필터:</span>
            <span className={`flex items-center gap-1 text-xs px-2.5 py-1 rounded-full font-medium ${
              activeFilter === 'expiring' ? 'bg-amber-100 text-amber-700' : 'bg-red-100 text-red-700'
            }`}>
              {activeFilter === 'expiring' ? '유효기간 임박' : '재고 부족'}
              <button onClick={onClearFilter} className="ml-0.5 hover:opacity-70"><X size={12} /></button>
            </span>
          </div>
        )}
      </div>

      <ReagentTable
        reagents={reagents}
        editingId={editingId} editValue={editValue}
        onEditValue={onEditValue} onStartEdit={onStartEdit}
        onConfirmEdit={onConfirmEdit} onCancelEdit={onCancelEdit}
        onDispatch={onDispatch}
      />
    </>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// AlertCard
// ══════════════════════════════════════════════════════════════════════════
function AlertCard({ icon, color, title, subtitle, count, active, onClick }) {
  const cfg = {
    amber: {
      border: active ? 'border-amber-400 bg-amber-50' : 'border-amber-200 bg-white hover:border-amber-400',
      iconBg: 'bg-amber-100 text-amber-600', count: 'text-amber-600',
    },
    red: {
      border: active ? 'border-red-400 bg-red-50' : 'border-red-200 bg-white hover:border-red-400',
      iconBg: 'bg-red-100 text-red-600', count: 'text-red-600',
    },
  }[color]

  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-4 p-4 rounded-xl border-2 transition-all w-full text-left shadow-sm ${cfg.border}`}
    >
      <div className={`p-2.5 rounded-xl ${cfg.iconBg}`}>{icon}</div>
      <div className="flex-1 min-w-0">
        <p className="text-sm font-semibold text-slate-700 truncate">{title}</p>
        <p className="text-xs text-slate-400">{subtitle}</p>
      </div>
      <div className="text-right">
        <div className={`text-3xl font-bold tabular-nums ${cfg.count}`}>{count}</div>
        <div className="text-xs text-slate-400">건</div>
      </div>
    </button>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// ReagentTable
// ══════════════════════════════════════════════════════════════════════════
function ReagentTable({
  reagents, editingId, editValue,
  onEditValue, onStartEdit, onConfirmEdit, onCancelEdit, onDispatch,
}) {
  return (
    <div className="bg-white rounded-xl border border-slate-100 shadow-sm overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-blue-600 text-white">
              {['시약명', '제조사', '구분', 'Lot No', '현재재고', '최소재고 (편집)', '유효기간', '누적출고', '출고'].map((h) => (
                <th key={h} className="px-4 py-2.5 text-left text-xs font-semibold whitespace-nowrap tracking-wide">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {reagents.length === 0 ? (
              <tr><td colSpan={9} className="text-center py-16 text-slate-400 text-sm">표시할 데이터가 없습니다.</td></tr>
            ) : reagents.map((r, i) => {
              const days = diffDays(r.expiryDate)
              const isExpiring = days >= 0 && days < 15
              const isExpired = days < 0
              const isLow = r.currentStock < r.minStock
              const typeCfg = {
                Reagent: 'bg-blue-100 text-blue-700',
                Cal:     'bg-purple-100 text-purple-700',
                Con:     'bg-orange-100 text-orange-700',
              }
              return (
                <tr key={r.id} className={`border-t border-slate-50 hover:bg-blue-50/30 transition-colors ${i % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'}`}>
                  <td className="px-4 py-3 font-medium text-slate-800 max-w-[200px]">
                    <div className="truncate">{r.name}</div>
                  </td>
                  <td className="px-4 py-3 text-slate-500 whitespace-nowrap">{r.manufacturer}</td>
                  <td className="px-4 py-3 whitespace-nowrap">
                    <span className={`text-[11px] font-semibold px-2 py-0.5 rounded-full ${typeCfg[r.reagentType] ?? typeCfg.Reagent}`}>
                      {r.reagentType ?? 'Reagent'}
                    </span>
                  </td>
                  <td className="px-4 py-3 font-mono text-xs text-slate-500 whitespace-nowrap">{r.lotNo}</td>
                  <td className="px-4 py-3 whitespace-nowrap">
                    <span className={`font-bold text-base ${isLow ? 'text-red-600' : 'text-slate-800'}`}>{r.currentStock}</span>
                    {isLow && <span className="ml-1.5 text-[10px] font-semibold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full">부족</span>}
                  </td>
                  <td className="px-4 py-3 whitespace-nowrap">
                    {editingId === r.id ? (
                      <div className="flex items-center gap-1">
                        <input
                          type="number" min="0" value={editValue}
                          onChange={(e) => onEditValue(e.target.value)}
                          onKeyDown={(e) => { if (e.key === 'Enter') onConfirmEdit(r.id); if (e.key === 'Escape') onCancelEdit() }}
                          className="w-16 border border-blue-300 rounded px-2 py-0.5 text-sm focus:outline-none focus:ring-1 focus:ring-blue-400"
                          autoFocus
                        />
                        <button onClick={() => onConfirmEdit(r.id)} className="text-blue-600 hover:text-blue-800 p-0.5"><Check size={14} /></button>
                        <button onClick={onCancelEdit} className="text-slate-400 hover:text-slate-600 p-0.5"><X size={14} /></button>
                      </div>
                    ) : (
                      <button onClick={() => onStartEdit(r)} className="flex items-center gap-1.5 text-slate-600 hover:text-blue-600 group transition-colors" title="클릭하여 수정">
                        <span className="font-medium">{r.minStock}</span>
                        <Edit2 size={11} className="text-slate-300 group-hover:text-blue-500 transition-colors" />
                      </button>
                    )}
                  </td>
                  <td className="px-4 py-3 whitespace-nowrap">
                    {isExpired ? (
                      <span className="text-xs px-2 py-0.5 rounded-full bg-slate-200 text-slate-500 font-medium">만료됨</span>
                    ) : isExpiring ? (
                      <span className="text-xs px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 font-bold">D-{days} ({r.expiryDate})</span>
                    ) : (
                      <span className="text-slate-500 text-xs">{r.expiryDate}</span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-slate-500 tabular-nums">{r.totalDispatched}</td>
                  <td className="px-4 py-3">
                    <button
                      onClick={() => onDispatch(r)}
                      disabled={r.currentStock <= 0}
                      className="bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed text-white text-xs font-semibold px-3 py-1.5 rounded-lg transition-colors whitespace-nowrap"
                    >
                      출고
                    </button>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
      <div className="px-4 py-2 bg-slate-50 border-t border-slate-100 text-xs text-slate-400">
        총 {reagents.length}개 시약
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// HistoryTab (출고)
// ══════════════════════════════════════════════════════════════════════════
function HistoryTab({ logs, onExport }) {
  const [search, setSearch] = useState('')
  const filtered = useMemo(() => {
    if (!search.trim()) return logs
    const q = search.trim().toLowerCase()
    return logs.filter((l) => l.reagentName.toLowerCase().includes(q) || l.lotNo.toLowerCase().includes(q))
  }, [logs, search])

  const now = new Date()
  const thisMonth = `${now.getFullYear()}년 ${now.getMonth() + 1}월`

  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h2 className="text-base font-semibold text-slate-700">출고 이력 목록</h2>
          <p className="text-xs text-slate-400 mt-0.5">총 {logs.length}건의 출고 기록</p>
        </div>
        <button
          onClick={onExport}
          className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors shrink-0"
        >
          <Download size={15} />
          {thisMonth} 입출고 대장 다운로드
        </button>
      </div>
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={search} onChange={(e) => setSearch(e.target.value)}
          placeholder="시약명, Lot No 검색..."
          className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 bg-white"
        />
      </div>
      <div className="bg-white rounded-xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-blue-600 text-white">
                {['출고일시', '시약명', 'Lot No', '출고수량'].map((h) => (
                  <th key={h} className="px-4 py-2.5 text-left text-xs font-semibold whitespace-nowrap tracking-wide">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={4} className="text-center py-16 text-slate-400 text-sm">출고 이력이 없습니다.</td></tr>
              ) : filtered.map((l, i) => (
                <tr key={l.id} className={`border-t border-slate-50 hover:bg-blue-50/30 transition-colors ${i % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'}`}>
                  <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">{l.datetime}</td>
                  <td className="px-4 py-2.5 font-medium text-slate-800">{l.reagentName}</td>
                  <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">{l.lotNo}</td>
                  <td className="px-4 py-2.5">
                    <span className="bg-blue-100 text-blue-700 text-xs font-bold px-2 py-0.5 rounded-full">{l.qty}개</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="px-4 py-2 bg-slate-50 border-t border-slate-100 text-xs text-slate-400">
          {filtered.length}건 표시 중
        </div>
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// InboundHistoryTab (입고)
// ══════════════════════════════════════════════════════════════════════════
function InboundHistoryTab({ inboundLogs }) {
  const [search, setSearch] = useState('')
  const filtered = useMemo(() => {
    if (!search.trim()) return inboundLogs
    const q = search.trim().toLowerCase()
    return inboundLogs.filter(
      (l) => l.reagentName.toLowerCase().includes(q) || l.lotNo.toLowerCase().includes(q),
    )
  }, [inboundLogs, search])

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-base font-semibold text-slate-700">입고 이력 목록</h2>
        <p className="text-xs text-slate-400 mt-0.5">총 {inboundLogs.length}건의 입고 기록</p>
      </div>
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={search} onChange={(e) => setSearch(e.target.value)}
          placeholder="시약명, Lot No 검색..."
          className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-300 bg-white"
        />
      </div>
      <div className="bg-white rounded-xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-teal-600 text-white">
                {['입고일시', '시약명', 'Lot No', '입고수량', '비고'].map((h) => (
                  <th key={h} className="px-4 py-2.5 text-left text-xs font-semibold whitespace-nowrap tracking-wide">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={5} className="text-center py-16 text-slate-400 text-sm">입고 이력이 없습니다.</td></tr>
              ) : filtered.map((l, i) => (
                <tr key={l.id} className={`border-t border-slate-50 hover:bg-teal-50/30 transition-colors ${i % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'}`}>
                  <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">{l.datetime}</td>
                  <td className="px-4 py-2.5 font-medium text-slate-800">{l.reagentName}</td>
                  <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">{l.lotNo}</td>
                  <td className="px-4 py-2.5">
                    <span className="bg-teal-100 text-teal-700 text-xs font-bold px-2 py-0.5 rounded-full">{l.qty}개</span>
                  </td>
                  <td className="px-4 py-2.5 text-xs text-slate-400">{l.notes || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="px-4 py-2 bg-slate-50 border-t border-slate-100 text-xs text-slate-400">
          {filtered.length}건 표시 중
        </div>
      </div>
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// InboundModal
// ══════════════════════════════════════════════════════════════════════════
function InboundModal({ reagents, syncing, onClose, onSubmit, onBulkInbound, onDownloadTemplate }) {
  const [mode, setMode] = useState('new')
  const [form, setForm] = useState({
    name: '', reagentType: 'Reagent', manufacturer: '', lotNo: '', qty: 1, minStock: 0,
    expiryDate: '', notes: '', reagentId: reagents[0]?.id ?? '', newLotNo: '',
  })
  const setF = (k, v) => setForm((prev) => ({ ...prev, [k]: v }))

  const handleSubmit = () => {
    if (mode === 'new') {
      if (!form.name.trim())  { alert('시약명을 입력해주세요.'); return }
      if (!form.lotNo.trim()) { alert('Lot No를 입력해주세요.'); return }
      if (Number(form.qty) < 1) { alert('입고량을 1 이상 입력해주세요.'); return }
    } else {
      if (!form.reagentId)    { alert('시약을 선택해주세요.'); return }
      if (Number(form.qty) < 1) { alert('추가 수량을 1 이상 입력해주세요.'); return }
    }
    onSubmit({ mode, form })
  }

  const handleBulkFile = (e) => {
    const file = e.target.files[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true })
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
      onBulkInbound(rows)
    }
    reader.readAsArrayBuffer(file)
    e.target.value = ''
  }

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md max-h-[90vh] flex flex-col">
        {/* 헤더 */}
        <div className="flex items-center justify-between px-6 py-4 border-b shrink-0">
          <div className="flex items-center gap-2">
            <PackagePlus size={18} className="text-teal-600" />
            <h2 className="text-base font-bold text-slate-800">입고 등록</h2>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 transition-colors">
            <X size={18} />
          </button>
        </div>
        {/* 모드 탭 */}
        <div className="flex border-b shrink-0">
          {[['new', '신규 Lot'], ['existing', '기존 재고 추가'], ['bulk', '엑셀 일괄']].map(([m, label]) => (
            <button
              key={m}
              onClick={() => setMode(m)}
              className={`flex-1 py-2.5 text-sm font-medium border-b-2 transition-colors ${
                mode === m
                  ? 'border-teal-500 text-teal-600'
                  : 'border-transparent text-slate-400 hover:text-slate-600'
              }`}
            >
              {label}
            </button>
          ))}
        </div>
        {/* 입력 폼 */}
        <div className="p-6 space-y-4 overflow-y-auto">
          {mode === 'new' ? (
            <>
              <div className="grid grid-cols-2 gap-3">
                <ModalField label="시약명 *" value={form.name} onChange={(v) => setF('name', v)} placeholder="예) CBC 희석액" />
                <div>
                  <label className="block text-xs font-semibold text-slate-500 mb-1.5">구분 *</label>
                  <select
                    value={form.reagentType}
                    onChange={(e) => setF('reagentType', e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-300 bg-white"
                  >
                    <option value="Reagent">Reagent</option>
                    <option value="Cal">Cal</option>
                    <option value="Con">Con</option>
                  </select>
                </div>
              </div>
              <ModalField label="제조사" value={form.manufacturer} onChange={(v) => setF('manufacturer', v)} placeholder="예) Sysmex" />
              <ModalField label="Lot No. *" value={form.lotNo} onChange={(v) => setF('lotNo', v)} placeholder="예) SX2025-001" />
              <div className="grid grid-cols-2 gap-3">
                <ModalField label="입고량 *" type="number" value={form.qty} onChange={(v) => setF('qty', v)} />
                <ModalField label="최소유지재고" type="number" value={form.minStock} onChange={(v) => setF('minStock', v)} />
              </div>
              <ModalField label="유효기간" type="date" value={form.expiryDate} onChange={(v) => setF('expiryDate', v)} />
              <ModalField label="비고" value={form.notes} onChange={(v) => setF('notes', v)} placeholder="(선택 사항)" />
            </>
          ) : mode === 'existing' ? (
            <>
              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">시약 선택 *</label>
                <select
                  value={form.reagentId}
                  onChange={(e) => setF('reagentId', e.target.value)}
                  className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-300 bg-white"
                >
                  {reagents.map((r) => (
                    <option key={r.id} value={r.id}>
                      {r.name} — {r.lotNo} (현재 {r.currentStock}개)
                    </option>
                  ))}
                </select>
              </div>
              <ModalField
                label="변경된 Lot No (없으면 빈칸)"
                value={form.newLotNo}
                onChange={(v) => setF('newLotNo', v)}
                placeholder="예) SX2025-002 (Lot 변경 시만 입력)"
              />
              <ModalField label="추가 수량 *" type="number" value={form.qty} onChange={(v) => setF('qty', v)} />
              <ModalField label="비고" value={form.notes} onChange={(v) => setF('notes', v)} placeholder="(선택 사항)" />
            </>
          ) : (
            <div className="space-y-4">
              <div className="bg-slate-50 rounded-xl p-4 space-y-1 text-xs text-slate-500">
                <p className="font-semibold text-slate-700 mb-2">엑셀 형식 안내</p>
                <p>필수: <span className="font-mono text-slate-700">시약명, Lot No, 입고량</span></p>
                <p>선택: <span className="font-mono text-slate-700">구분, 제조사, 최소유지재고, 유효기간, 비고</span></p>
                <p className="mt-2 text-teal-600">• Lot No가 기존과 일치하면 재고 추가</p>
                <p className="text-teal-600">• 새 Lot No면 신규 시약으로 자동 등록</p>
              </div>
              <button
                onClick={onDownloadTemplate}
                className="flex items-center gap-2 w-full justify-center border border-teal-300 text-teal-700 hover:bg-teal-50 text-sm font-medium py-2.5 rounded-lg transition-colors"
              >
                <Download size={14} />
                템플릿 다운로드
              </button>
              <label className={`flex flex-col items-center gap-2 w-full border-2 border-dashed rounded-xl py-8 cursor-pointer transition-colors ${syncing ? 'border-slate-200 bg-slate-50' : 'border-teal-200 hover:border-teal-400 hover:bg-teal-50'}`}>
                <Upload size={22} className={syncing ? 'text-slate-300' : 'text-teal-400'} />
                <span className="text-sm font-medium text-slate-600">
                  {syncing ? '처리 중...' : '엑셀 파일 선택 (.xlsx)'}
                </span>
                <input type="file" accept=".xlsx" className="hidden" disabled={syncing} onChange={handleBulkFile} />
              </label>
            </div>
          )}
        </div>
        {/* 버튼 */}
        <div className="flex gap-3 px-6 pb-5 shrink-0">
          <button
            onClick={onClose}
            className="flex-1 py-2.5 text-sm font-medium border border-slate-200 rounded-lg text-slate-600 hover:bg-slate-50 transition-colors"
          >
            취소
          </button>
          {mode !== 'bulk' && (
            <button
              onClick={handleSubmit}
              disabled={syncing}
              className="flex-1 py-2.5 text-sm font-semibold bg-teal-600 hover:bg-teal-700 disabled:bg-slate-200 disabled:text-slate-400 text-white rounded-lg transition-colors"
            >
              {syncing ? '저장 중...' : '등록하기'}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}

function ModalField({ label, value, onChange, type = 'text', placeholder = '' }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-slate-500 mb-1.5">{label}</label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        min={type === 'number' ? 0 : undefined}
        className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-300"
      />
    </div>
  )
}
