import XLSX from 'xlsx-js-style'

// ── 색상 팔레트 ───────────────────────────────────────────────────────────
const C = {
  // Sheet1 헤더: 딥 블루
  h1Bg: '1E3A8A', h1Fg: 'FFFFFF',
  // Sheet2 헤더: 딥 틸
  h2Bg: '0F766E', h2Fg: 'FFFFFF',
  // 타이틀 배경
  titleBg: 'DBEAFE', titleFg: '1E3A8A',
  // 데이터 행
  rowOdd: 'FFFFFF', rowEven: 'EFF6FF',
  // 테두리
  bThin: 'CBD5E1', bMedium: '64748B',
  // 상태 색
  danger: 'B91C1C', safe: '166534', muted: '64748B',
}

const FONT = '맑은 고딕'

// ── 스타일 빌더 ────────────────────────────────────────────────────────────
const border = (topBot = 'thin') => ({
  top:    { style: topBot,  color: { rgb: topBot === 'medium' ? C.bMedium : C.bThin } },
  bottom: { style: topBot,  color: { rgb: topBot === 'medium' ? C.bMedium : C.bThin } },
  left:   { style: 'thin',  color: { rgb: C.bThin } },
  right:  { style: 'thin',  color: { rgb: C.bThin } },
})

const hdrStyle = (bgRgb, fgRgb, align = 'center') => ({
  font: { name: FONT, sz: 10, bold: true, color: { rgb: fgRgb } },
  fill: { patternType: 'solid', fgColor: { rgb: bgRgb } },
  alignment: { horizontal: align, vertical: 'center', wrapText: false },
  border: border('medium'),
})

const dataStyle = (rowIdx, align = 'left', fontOverride = {}) => ({
  font: { name: FONT, sz: 10, color: { rgb: '1E293B' }, ...fontOverride },
  fill: { patternType: 'solid', fgColor: { rgb: rowIdx % 2 === 0 ? C.rowOdd : C.rowEven } },
  alignment: { horizontal: align, vertical: 'center' },
  border: border('thin'),
})

const titleStyle = (sz = 14) => ({
  font: { name: FONT, sz, bold: true, color: { rgb: C.titleFg } },
  fill: { patternType: 'solid', fgColor: { rgb: C.titleBg } },
  alignment: { horizontal: 'center', vertical: 'center' },
})

const metaStyle = () => ({
  font: { name: FONT, sz: 9, italic: true, color: { rgb: C.muted } },
  alignment: { horizontal: 'right', vertical: 'center' },
})

// ── 셀 쓰기 헬퍼 ──────────────────────────────────────────────────────────
const setCell = (ws, r, c, value, style) => {
  ws[XLSX.utils.encode_cell({ r, c })] = { v: value, s: style }
}

// ── 날짜/시간 분리 ─────────────────────────────────────────────────────────
const splitDatetime = (dt) => {
  if (!dt) return ['', '']
  const [d, t] = dt.split(' ')
  return [d ?? '', t ?? '']
}

const fmtCreatedAt = (createdAt) => {
  if (!createdAt) return '-'
  return String(createdAt).slice(0, 10)
}

// ══════════════════════════════════════════════════════════════════════════
// 메인: 월간 보고서 생성
// ══════════════════════════════════════════════════════════════════════════
export function generateMonthlyReport(logs, reagents, year, month) {
  const ym = `${year}-${String(month).padStart(2, '0')}`

  const monthlyLogs = [...logs]
    .filter((l) => l.datetime.startsWith(ym))
    .sort((a, b) => a.datetime.localeCompare(b.datetime))

  const wb = XLSX.utils.book_new()

  buildSheet1(wb, monthlyLogs, reagents, year, month)
  buildSheet2(wb, monthlyLogs, reagents, year, month)

  XLSX.writeFile(wb, `${year}년_${month}월_시약입출고대장.xlsx`)
}

// ══════════════════════════════════════════════════════════════════════════
// Sheet 1 — 월간 출고 상세내역
// ══════════════════════════════════════════════════════════════════════════
function buildSheet1(wb, monthlyLogs, reagents, year, month) {
  const ws = {}

  const COLS = [
    { label: '시약명',      wch: 30, align: 'left'   },
    { label: '제조사',      wch: 14, align: 'center' },
    { label: 'Lot No',      wch: 16, align: 'center' },
    { label: '유효기간',    wch: 12, align: 'center' },
    { label: '입고일자',    wch: 12, align: 'center' },
    { label: '초기 입고량', wch: 11, align: 'center' },
    { label: '출고 일자',   wch: 12, align: 'center' },
    { label: '출고 시간',   wch: 10, align: 'center' },
    { label: '출고 수량',   wch: 10, align: 'center' },
  ]
  const NC = COLS.length

  let R = 0

  // 제목 행
  setCell(ws, R, 0, `${year}년 ${month}월 시약 입출고 대장`, titleStyle(14))
  R++

  // 메타 행
  setCell(ws, R, NC - 1,
    `출력일: ${new Date().toLocaleDateString('ko-KR')}  |  총 ${monthlyLogs.length}건`,
    metaStyle())
  R++

  // 빈 행
  R++

  // 헤더 행
  COLS.forEach((col, c) =>
    setCell(ws, R, c, col.label, hdrStyle(C.h1Bg, C.h1Fg, col.align)),
  )
  R++

  // 데이터 행
  if (monthlyLogs.length === 0) {
    setCell(ws, R, 0, '이번 달 출고 이력이 없습니다.',
      { font: { name: FONT, sz: 10, italic: true, color: { rgb: C.muted } },
        alignment: { horizontal: 'center' }, border: border('thin') })
    R++
  } else {
    monthlyLogs.forEach((log, idx) => {
      const rg = reagents.find((r) => r.id === log.reagentId) ?? {}
      const [date, time] = splitDatetime(log.datetime)
      const row = [
        { v: rg.name ?? log.reagentName,        align: 'left'   },
        { v: rg.manufacturer ?? '',              align: 'center' },
        { v: log.lotNo,                          align: 'center' },
        { v: rg.expiryDate ?? '',                align: 'center' },
        { v: fmtCreatedAt(rg.createdAt),         align: 'center' },
        { v: rg.receivedQty ?? 0,                align: 'center' },
        { v: date,                               align: 'center' },
        { v: time,                               align: 'center' },
        { v: log.qty,                            align: 'center' },
      ]
      row.forEach(({ v, align }, c) =>
        setCell(ws, R, c, v, dataStyle(idx, align)),
      )
      R++
    })
  }

  // 합계 행
  if (monthlyLogs.length > 0) {
    const total = monthlyLogs.reduce((s, l) => s + l.qty, 0)
    const sumStyle = {
      font: { name: FONT, sz: 10, bold: true, color: { rgb: C.h1Fg } },
      fill: { patternType: 'solid', fgColor: { rgb: C.h1Bg } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: border('medium'),
    }
    for (let c = 0; c < NC - 1; c++) setCell(ws, R, c, c === 0 ? '합  계' : '', sumStyle)
    setCell(ws, R, NC - 1, total, sumStyle)
    R++
  }

  // 병합: 제목 행 전체 병합
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
  ]
  ws['!ref']  = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: R - 1, c: NC - 1 })
  ws['!cols'] = COLS.map((col) => ({ wch: col.wch }))
  ws['!rows'] = [
    { hpt: 30 }, // 제목
    { hpt: 16 }, // 메타
    { hpt: 6  }, // 빈줄
    { hpt: 22 }, // 헤더
  ]

  XLSX.utils.book_append_sheet(wb, ws, '월간 출고 상세내역')
}

// ══════════════════════════════════════════════════════════════════════════
// Sheet 2 — 시약별 월간 요약
// ══════════════════════════════════════════════════════════════════════════
function buildSheet2(wb, monthlyLogs, reagents, year, month) {
  const ws = {}

  const COLS = [
    { label: '시약명',           wch: 30, align: 'left'   },
    { label: '제조사',           wch: 14, align: 'center' },
    { label: 'Lot No',           wch: 16, align: 'center' },
    { label: '이번 달 총 출고량', wch: 16, align: 'center' },
    { label: '현재 잔여재고',    wch: 14, align: 'center' },
    { label: '최소 유지재고',    wch: 14, align: 'center' },
    { label: '재고 상태',        wch: 12, align: 'center' },
  ]
  const NC = COLS.length

  let R = 0

  // 제목
  setCell(ws, R, 0, `${year}년 ${month}월 시약별 출고 요약`, titleStyle(13))
  R++
  R++ // 빈줄
  R++ // 빈줄

  // 헤더
  COLS.forEach((col, c) =>
    setCell(ws, R, c, col.label, hdrStyle(C.h2Bg, C.h2Fg, col.align)),
  )
  R++

  // 시약별 집계
  const map = new Map()
  monthlyLogs.forEach((log) => {
    if (!map.has(log.reagentId)) {
      const rg = reagents.find((r) => r.id === log.reagentId) ?? {}
      map.set(log.reagentId, {
        name:         rg.name ?? log.reagentName,
        manufacturer: rg.manufacturer ?? '',
        lotNo:        log.lotNo,
        totalQty:     0,
        currentStock: rg.currentStock ?? 0,
        minStock:     rg.minStock ?? 0,
      })
    }
    map.get(log.reagentId).totalQty += log.qty
  })

  const rows = [...map.values()].sort((a, b) => a.name.localeCompare(b.name))

  if (rows.length === 0) {
    setCell(ws, R, 0, '이번 달 출고된 시약이 없습니다.',
      { font: { name: FONT, sz: 10, italic: true, color: { rgb: C.muted } },
        alignment: { horizontal: 'center' }, border: border('thin') })
    R++
  } else {
    rows.forEach((row, idx) => {
      const isLow    = row.currentStock < row.minStock
      const statusTx = isLow ? '⚠ 재고 부족' : '● 정상'
      const stColor  = isLow ? C.danger : C.safe
      const stockFont = isLow
        ? { bold: true, color: { rgb: C.danger } }
        : { color: { rgb: C.safe } }

      const cells = [
        { v: row.name,         align: 'left',   font: {} },
        { v: row.manufacturer, align: 'center', font: {} },
        { v: row.lotNo,        align: 'center', font: {} },
        { v: row.totalQty,     align: 'center', font: { bold: true } },
        { v: row.currentStock, align: 'center', font: stockFont },
        { v: row.minStock,     align: 'center', font: {} },
        { v: statusTx,         align: 'center', font: { bold: true, color: { rgb: stColor } } },
      ]

      cells.forEach(({ v, align, font }, c) =>
        setCell(ws, R, c, v, dataStyle(idx, align, font)),
      )
      R++
    })

    // 합계 행
    const grandTotal = rows.reduce((s, r) => s + r.totalQty, 0)
    const sumStyle = {
      font: { name: FONT, sz: 10, bold: true, color: { rgb: C.h2Fg } },
      fill: { patternType: 'solid', fgColor: { rgb: C.h2Bg } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: border('medium'),
    }
    for (let c = 0; c < NC - 1; c++)
      setCell(ws, R, c, c === 0 ? `총  ${rows.length}종` : '', sumStyle)
    setCell(ws, R, 3, grandTotal, sumStyle)
    R++
  }

  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } }]
  ws['!ref']  = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: R - 1, c: NC - 1 })
  ws['!cols'] = COLS.map((col) => ({ wch: col.wch }))
  ws['!rows'] = [
    { hpt: 28 }, // 제목
    { hpt: 6  },
    { hpt: 6  },
    { hpt: 22 }, // 헤더
  ]

  XLSX.utils.book_append_sheet(wb, ws, '시약별 월간 요약')
}
