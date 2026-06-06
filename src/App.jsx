import { useState, useMemo, useCallback } from 'react'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts'
import {
  AlertTriangle, TrendingDown, Package, Upload, Download,
  Search, Edit2, Check, X, FlaskConical, Clock, HardDriveDownload,
} from 'lucide-react'
import * as XLSX from 'xlsx'

// ── 날짜 유틸 ──────────────────────────────────────────────────────────────
const fmtDate = (d) => {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
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

// ── 초기 시약 Mock 데이터 ────────────────────────────────────────────────
const INIT_REAGENTS = [
  {
    id: 1, name: '혈액형 검사 시약 (ABO/Rh)', manufacturer: 'Bio-Rad',
    lotNo: 'BR2024-001', receivedQty: 100, currentStock: 8,
    minStock: 10, expiryDate: '2026-06-15', totalDispatched: 92,
  },
  {
    id: 2, name: 'CBC 희석액', manufacturer: 'Sysmex',
    lotNo: 'SX2024-102', receivedQty: 200, currentStock: 45,
    minStock: 30, expiryDate: '2026-08-20', totalDispatched: 155,
  },
  {
    id: 3, name: 'PT/APTT 응고 시약', manufacturer: 'Stago',
    lotNo: 'ST2024-055', receivedQty: 50, currentStock: 12,
    minStock: 15, expiryDate: '2026-06-10', totalDispatched: 38,
  },
  {
    id: 4, name: 'HbA1c 측정 시약', manufacturer: 'Tosoh',
    lotNo: 'TS2024-210', receivedQty: 80, currentStock: 30,
    minStock: 20, expiryDate: '2026-09-05', totalDispatched: 50,
  },
  {
    id: 5, name: 'Troponin I 키트', manufacturer: 'Abbott',
    lotNo: 'AB2024-334', receivedQty: 60, currentStock: 5,
    minStock: 8, expiryDate: '2026-07-30', totalDispatched: 55,
  },
  {
    id: 6, name: '요검사 스트립', manufacturer: 'Roche',
    lotNo: 'RC2024-412', receivedQty: 300, currentStock: 90,
    minStock: 50, expiryDate: '2026-10-15', totalDispatched: 210,
  },
  {
    id: 7, name: 'CRP 정량 시약', manufacturer: 'Beckman',
    lotNo: 'BK2024-078', receivedQty: 40, currentStock: 18,
    minStock: 10, expiryDate: '2026-06-18', totalDispatched: 22,
  },
]

// ── 최근 7일 가상 출고 로그 생성 ─────────────────────────────────────────
function genMockLogs() {
  const logs = []
  let id = 1
  const now = new Date()
  for (let d = 6; d >= 0; d--) {
    const base = new Date(now)
    base.setDate(base.getDate() - d)
    const count = 2 + (d % 3)
    for (let i = 0; i < count; i++) {
      const r = INIT_REAGENTS[i % INIT_REAGENTS.length]
      const hour = new Date(base)
      hour.setHours(8 + ((i * 3) % 10), (i * 13) % 60, 0, 0)
      logs.push({
        id: id++,
        reagentId: r.id,
        reagentName: r.name,
        lotNo: r.lotNo,
        qty: 1 + (i % 3),
        datetime: fmtDateTime(hour),
      })
    }
  }
  return logs.reverse()
}

// ══════════════════════════════════════════════════════════════════════════
// App (root)
// ══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [reagents, setReagents] = useState(INIT_REAGENTS)
  const [logs, setLogs] = useState(genMockLogs)
  const [activeTab, setActiveTab] = useState('dashboard')
  const [search, setSearch] = useState('')
  const [activeFilter, setActiveFilter] = useState(null) // 'expiring' | 'lowStock'
  const [editingId, setEditingId] = useState(null)
  const [editValue, setEditValue] = useState('')

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
      return {
        date: `${d.getMonth() + 1}/${d.getDate()}`,
        출고량: total,
      }
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
  const handleDispatch = useCallback((r) => {
    if (r.currentStock <= 0) {
      alert(`${r.name} 재고가 없습니다.`)
      return
    }
    const now = new Date()
    setReagents((prev) =>
      prev.map((x) =>
        x.id === r.id
          ? { ...x, currentStock: x.currentStock - 1, totalDispatched: x.totalDispatched + 1 }
          : x,
      ),
    )
    setLogs((prev) => [
      {
        id: Date.now(),
        reagentId: r.id,
        reagentName: r.name,
        lotNo: r.lotNo,
        qty: 1,
        datetime: fmtDateTime(now),
      },
      ...prev,
    ])
  }, [])

  // ── 최소 재고 편집 ─────────────────────────────────────────────────────
  const startEdit = useCallback((r) => {
    setEditingId(r.id)
    setEditValue(String(r.minStock))
  }, [])
  const confirmEdit = useCallback(
    (id) => {
      const val = parseInt(editValue, 10)
      if (!isNaN(val) && val >= 0) {
        setReagents((prev) => prev.map((r) => (r.id === id ? { ...r, minStock: val } : r)))
      }
      setEditingId(null)
    },
    [editValue],
  )
  const cancelEdit = useCallback(() => setEditingId(null), [])

  // ── 엑셀 업로드 (일반 재고파일 + 백업 복구 통합) ──────────────────────
  const handleUpload = useCallback((e) => {
    const file = e.target.files[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' })
        const isBackup = wb.SheetNames.includes('재고')

        if (isBackup) {
          // ── 백업 복구: 재고 + 출고이력 모두 복원 ──
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
            expiryDate: String(row['유효기간'] ?? ''),
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

          setReagents(restoredReagents)
          setLogs(restoredLogs)
          alert(`✅ 백업 복구 완료\n시약 ${restoredReagents.length}종 · 출고이력 ${restoredLogs.length}건이 복원되었습니다.`)
        } else {
          // ── 일반 재고 파일 업로드 ──
          const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
          const mapped = data.map((row, idx) => ({
            id: idx + 1,
            name: row['시약명'] ?? row['name'] ?? '',
            manufacturer: row['제조사'] ?? row['manufacturer'] ?? '',
            lotNo: row['Lot No'] ?? row['lotNo'] ?? '',
            receivedQty: Number(row['입고량'] ?? row['receivedQty'] ?? 0),
            currentStock: Number(row['현재재고'] ?? row['currentStock'] ?? 0),
            minStock: Number(row['최소유지재고'] ?? row['minStock'] ?? 0),
            expiryDate: String(row['유효기간'] ?? row['expiryDate'] ?? ''),
            totalDispatched: Number(row['누적출고량'] ?? row['totalDispatched'] ?? 0),
          }))
          setReagents(mapped)
        }
      } catch {
        alert('파일을 읽는 중 오류가 발생했습니다. 형식을 확인해 주세요.')
      }
    }
    reader.readAsArrayBuffer(file)
    e.target.value = ''
  }, [])

  // ── 전체 상태 백업 다운로드 ────────────────────────────────────────────
  const handleBackup = useCallback(() => {
    const now = new Date()
    const stamp = `${fmtDate(now)}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`

    // Sheet 1: 재고
    const reagentRows = reagents.map((r) => ({
      'id': r.id,
      '시약명': r.name,
      '제조사': r.manufacturer,
      'Lot No': r.lotNo,
      '입고량': r.receivedQty,
      '현재재고': r.currentStock,
      '최소유지재고': r.minStock,
      '유효기간': r.expiryDate,
      '누적출고량': r.totalDispatched,
    }))
    const wsReagents = XLSX.utils.json_to_sheet(reagentRows)
    wsReagents['!cols'] = [
      { wch: 6 }, { wch: 32 }, { wch: 14 }, { wch: 16 },
      { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 10 },
    ]

    // Sheet 2: 출고이력_전체
    const logRows = logs.map((l) => ({
      'id': l.id,
      'reagentId': l.reagentId,
      '시약명': l.reagentName,
      'Lot No': l.lotNo,
      '출고수량': l.qty,
      '출고일시': l.datetime,
    }))
    const wsLogs = XLSX.utils.json_to_sheet(logRows)
    wsLogs['!cols'] = [
      { wch: 14 }, { wch: 8 }, { wch: 32 }, { wch: 16 }, { wch: 8 }, { wch: 18 },
    ]

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, wsReagents, '재고')
    XLSX.utils.book_append_sheet(wb, wsLogs, '출고이력_전체')
    XLSX.writeFile(wb, `재고백업_${stamp}.xlsx`)
  }, [reagents, logs])

  // ── 이번 달 엑셀 다운로드 ──────────────────────────────────────────────
  const handleExport = useCallback(() => {
    const now = new Date()
    const ym = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
    const monthly = logs.filter((l) => l.datetime.startsWith(ym))
    if (monthly.length === 0) {
      alert('이번 달 출고 이력이 없습니다.')
      return
    }
    const rows = monthly.map((l) => ({
      '출고일시': l.datetime,
      '시약명': l.reagentName,
      'Lot No': l.lotNo,
      '출고수량': l.qty,
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [{ wch: 18 }, { wch: 30 }, { wch: 16 }, { wch: 10 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, '출고이력')
    XLSX.writeFile(wb, `출고이력_${ym}.xlsx`)
  }, [logs])

  const toggleFilter = useCallback(
    (type) => setActiveFilter((prev) => (prev === type ? null : type)),
    [],
  )

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
              <button
                onClick={handleBackup}
                className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors whitespace-nowrap"
                title="현재 재고 + 출고이력 전체를 엑셀로 저장합니다"
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
              { key: 'history', label: '출고 이력' },
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
        ) : (
          <HistoryTab logs={logs} onExport={handleExport} />
        )}
      </main>
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
      {/* Alert Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <AlertCard
          icon={<Clock size={20} />}
          color="amber"
          title="유효기간 임박 시약"
          subtitle="유효기간 15일 미만"
          count={expiring.length}
          active={activeFilter === 'expiring'}
          onClick={() => onToggleFilter('expiring')}
        />
        <AlertCard
          icon={<TrendingDown size={20} />}
          color="red"
          title="재고 부족 시약"
          subtitle="최소 유지 재고 미만"
          count={lowStock.length}
          active={activeFilter === 'lowStock'}
          onClick={() => onToggleFilter('lowStock')}
        />
      </div>

      {/* Chart */}
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

      {/* Search + Filter badge */}
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
            <span
              className={`flex items-center gap-1 text-xs px-2.5 py-1 rounded-full font-medium ${
                activeFilter === 'expiring'
                  ? 'bg-amber-100 text-amber-700'
                  : 'bg-red-100 text-red-700'
              }`}
            >
              {activeFilter === 'expiring' ? '유효기간 임박' : '재고 부족'}
              <button onClick={onClearFilter} className="ml-0.5 hover:opacity-70">
                <X size={12} />
              </button>
            </span>
          </div>
        )}
      </div>

      {/* Reagent Table */}
      <ReagentTable
        reagents={reagents}
        editingId={editingId}
        editValue={editValue}
        onEditValue={onEditValue}
        onStartEdit={onStartEdit}
        onConfirmEdit={onConfirmEdit}
        onCancelEdit={onCancelEdit}
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
      iconBg: 'bg-amber-100 text-amber-600',
      count: 'text-amber-600',
      badge: count > 0 ? 'bg-amber-500' : 'bg-slate-300',
    },
    red: {
      border: active ? 'border-red-400 bg-red-50' : 'border-red-200 bg-white hover:border-red-400',
      iconBg: 'bg-red-100 text-red-600',
      count: 'text-red-600',
      badge: count > 0 ? 'bg-red-500' : 'bg-slate-300',
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
              {['시약명', '제조사', 'Lot No', '현재재고', '최소재고 (편집)', '유효기간', '누적출고', '출고'].map(
                (h) => (
                  <th
                    key={h}
                    className="px-4 py-2.5 text-left text-xs font-semibold whitespace-nowrap tracking-wide"
                  >
                    {h}
                  </th>
                ),
              )}
            </tr>
          </thead>
          <tbody>
            {reagents.length === 0 ? (
              <tr>
                <td colSpan={8} className="text-center py-16 text-slate-400 text-sm">
                  표시할 데이터가 없습니다.
                </td>
              </tr>
            ) : (
              reagents.map((r, i) => {
                const days = diffDays(r.expiryDate)
                const isExpiring = days >= 0 && days < 15
                const isExpired = days < 0
                const isLow = r.currentStock < r.minStock

                return (
                  <tr
                    key={r.id}
                    className={`border-t border-slate-50 hover:bg-blue-50/30 transition-colors ${
                      i % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                    }`}
                  >
                    {/* 시약명 */}
                    <td className="px-4 py-3 font-medium text-slate-800 max-w-[200px]">
                      <div className="truncate">{r.name}</div>
                    </td>
                    {/* 제조사 */}
                    <td className="px-4 py-3 text-slate-500 whitespace-nowrap">{r.manufacturer}</td>
                    {/* Lot No */}
                    <td className="px-4 py-3 font-mono text-xs text-slate-500 whitespace-nowrap">
                      {r.lotNo}
                    </td>
                    {/* 현재재고 */}
                    <td className="px-4 py-3 whitespace-nowrap">
                      <span
                        className={`font-bold text-base ${isLow ? 'text-red-600' : 'text-slate-800'}`}
                      >
                        {r.currentStock}
                      </span>
                      {isLow && (
                        <span className="ml-1.5 text-[10px] font-semibold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full">
                          부족
                        </span>
                      )}
                    </td>
                    {/* 최소재고 (편집) */}
                    <td className="px-4 py-3 whitespace-nowrap">
                      {editingId === r.id ? (
                        <div className="flex items-center gap-1">
                          <input
                            type="number"
                            min="0"
                            value={editValue}
                            onChange={(e) => onEditValue(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') onConfirmEdit(r.id)
                              if (e.key === 'Escape') onCancelEdit()
                            }}
                            className="w-16 border border-blue-300 rounded px-2 py-0.5 text-sm focus:outline-none focus:ring-1 focus:ring-blue-400"
                            autoFocus
                          />
                          <button
                            onClick={() => onConfirmEdit(r.id)}
                            className="text-blue-600 hover:text-blue-800 p-0.5"
                          >
                            <Check size={14} />
                          </button>
                          <button
                            onClick={onCancelEdit}
                            className="text-slate-400 hover:text-slate-600 p-0.5"
                          >
                            <X size={14} />
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => onStartEdit(r)}
                          className="flex items-center gap-1.5 text-slate-600 hover:text-blue-600 group transition-colors"
                          title="클릭하여 수정"
                        >
                          <span className="font-medium">{r.minStock}</span>
                          <Edit2
                            size={11}
                            className="text-slate-300 group-hover:text-blue-500 transition-colors"
                          />
                        </button>
                      )}
                    </td>
                    {/* 유효기간 */}
                    <td className="px-4 py-3 whitespace-nowrap">
                      {isExpired ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-slate-200 text-slate-500 font-medium">
                          만료됨
                        </span>
                      ) : isExpiring ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 font-bold">
                          D-{days} ({r.expiryDate})
                        </span>
                      ) : (
                        <span className="text-slate-500 text-xs">{r.expiryDate}</span>
                      )}
                    </td>
                    {/* 누적출고 */}
                    <td className="px-4 py-3 text-slate-500 tabular-nums">{r.totalDispatched}</td>
                    {/* 출고 버튼 */}
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
              })
            )}
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
// HistoryTab
// ══════════════════════════════════════════════════════════════════════════
function HistoryTab({ logs, onExport }) {
  const [search, setSearch] = useState('')

  const filtered = useMemo(() => {
    if (!search.trim()) return logs
    const q = search.trim().toLowerCase()
    return logs.filter(
      (l) => l.reagentName.toLowerCase().includes(q) || l.lotNo.toLowerCase().includes(q),
    )
  }, [logs, search])

  const now = new Date()
  const thisMonth = `${now.getFullYear()}년 ${now.getMonth() + 1}월`

  return (
    <div className="space-y-4">
      {/* Header row */}
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
          {thisMonth} 엑셀 다운로드
        </button>
      </div>

      {/* Search */}
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="시약명, Lot No 검색..."
          className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 bg-white"
        />
      </div>

      {/* Log Table */}
      <div className="bg-white rounded-xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-blue-600 text-white">
                {['출고일시', '시약명', 'Lot No', '출고수량'].map((h) => (
                  <th
                    key={h}
                    className="px-4 py-2.5 text-left text-xs font-semibold whitespace-nowrap tracking-wide"
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={4} className="text-center py-16 text-slate-400 text-sm">
                    출고 이력이 없습니다.
                  </td>
                </tr>
              ) : (
                filtered.map((l, i) => (
                  <tr
                    key={l.id}
                    className={`border-t border-slate-50 hover:bg-blue-50/30 transition-colors ${
                      i % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                    }`}
                  >
                    <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">
                      {l.datetime}
                    </td>
                    <td className="px-4 py-2.5 font-medium text-slate-800">{l.reagentName}</td>
                    <td className="px-4 py-2.5 font-mono text-xs text-slate-500 whitespace-nowrap">
                      {l.lotNo}
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="bg-blue-100 text-blue-700 text-xs font-bold px-2 py-0.5 rounded-full">
                        {l.qty}개
                      </span>
                    </td>
                  </tr>
                ))
              )}
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
