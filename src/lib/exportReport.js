import XLSX from 'xlsx-js-style'

// ── 색상 ──────────────────────────────────────────────────────────────────
const C = {
  // 고정 열 헤더 (딥 블루)
  fixHdrBg: '1E3A8A', fixHdrFg: 'FFFFFF',
  // 날짜 열 헤더 (미드 블루)
  dayHdrBg: '1D4ED8', dayHdrFg: 'FFFFFF',
  // 날짜 열 헤더 – 해당 월 범위 외
  dayHdrOffBg: 'CBD5E1', dayHdrOffFg: '64748B',
  // 타이틀
  titleBg: 'DBEAFE', titleFg: '1E3A8A',
  // 데이터 행
  rowOdd: 'FFFFFF', rowEven: 'EFF6FF',
  // 날짜 셀 – 출고 있음
  dayActiveBg: 'DBEAFE', dayActiveFg: '1E40AF',
  // 날짜 셀 – 월 범위 외
  dayOffBg: 'F1F5F9',
  // 합계 행
  totalBg: 'FEF9C3', totalFg: '78350F',
  // 상태 색
  danger: 'B91C1C', safe: '15803D', muted: '94A3B8',
  // 테두리
  bThin: 'CBD5E1', bMid: '475569',
}

const FONT = '맑은 고딕'

// ── 헬퍼 ──────────────────────────────────────────────────────────────────
const borders = (w = 'thin') => {
  const s = { style: w, color: { rgb: w === 'medium' ? C.bMid : C.bThin } }
  return { top: s, bottom: s, left: s, right: s }
}

/** 셀 쓰기 – 값 타입 자동 감지 */
const wc = (ws, r, c, v, s) => {
  const t = v === '' || v == null ? 's' : typeof v === 'number' ? 'n' : 's'
  ws[XLSX.utils.encode_cell({ r, c })] = { v: v ?? '', t, s }
}

// ── 고정 컬럼 정의 (A~I) ───────────────────────────────────────────────────
const FIXED = [
  { label: '시약명',     wch: 28, align: 'left'   },
  { label: '제조사',     wch: 14, align: 'center' },
  { label: 'Lot No.',    wch: 16, align: 'center' },
  { label: '유효기간',   wch: 12, align: 'center' },
  { label: '입고일자',   wch: 12, align: 'center' },
  { label: '입고량',     wch:  8, align: 'center' },
  { label: '전월이월량', wch: 10, align: 'center' },
  { label: '당월재고',   wch: 10, align: 'center' },
  { label: '사용합계',   wch: 10, align: 'center' },
]
const NF = FIXED.length // 9

// ── 메인 진입점 ────────────────────────────────────────────────────────────
export function generateMonthlyReport(logs, reagents, year, month, inboundLogs = []) {
  const ym        = `${year}-${String(month).padStart(2, '0')}`
  const lastDay   = new Date(year, month, 0).getDate()
  const monthly   = logs.filter((l) => l.datetime.startsWith(ym))

  // ── 시약별 · 일별 사용량 ─────────────────────────────────────────────
  const usageByRg = new Map()
  monthly.forEach((log) => {
    const day = parseInt(log.datetime.slice(8, 10), 10)
    if (!usageByRg.has(log.reagentId)) usageByRg.set(log.reagentId, new Map())
    const dm = usageByRg.get(log.reagentId)
    dm.set(day, (dm.get(day) ?? 0) + log.qty)
  })

  // ── 날짜별 전체 합계 ──────────────────────────────────────────────────
  const dayTotals = new Map()
  for (let d = 1; d <= lastDay; d++) {
    const tot = monthly
      .filter((l) => parseInt(l.datetime.slice(8, 10), 10) === d)
      .reduce((s, l) => s + l.qty, 0)
    dayTotals.set(d, tot)
  }

  // ── 시약별 당월 입고량 ────────────────────────────────────────────────
  const monthlyInbound = new Map()
  inboundLogs
    .filter((l) => l.datetime.startsWith(ym))
    .forEach((l) => {
      monthlyInbound.set(l.reagentId, (monthlyInbound.get(l.reagentId) ?? 0) + l.qty)
    })

  const wb = XLSX.utils.book_new()
  buildPivotSheet(wb, reagents, usageByRg, dayTotals, monthlyInbound, year, month, lastDay, monthly.length)
  XLSX.writeFile(wb, `${year}년_${month}월_시약사용대장.xlsx`)
}

// ── 피벗 시트 빌더 ────────────────────────────────────────────────────────
function buildPivotSheet(wb, reagents, usageByRg, dayTotals, monthlyInbound, year, month, lastDay, totalLogs) {
  const ws   = {}
  const TCOL = NF + 31 // 전체 컬럼 수 (고정9 + 날짜31)
  let R = 0

  // ── 1행: 타이틀 ──────────────────────────────────────────────────────
  const titleS = {
    font: { name: FONT, sz: 14, bold: true, color: { rgb: C.titleFg } },
    fill: { patternType: 'solid', fgColor: { rgb: C.titleBg } },
    alignment: { horizontal: 'center', vertical: 'center' },
  }
  wc(ws, R, 0, `${year}년 ${month}월  시약 사용 현황`, titleS)
  // 출력 메타 (우측 끝)
  wc(ws, R, TCOL - 1,
    `출력일: ${new Date().toLocaleDateString('ko-KR')}  /  총 ${totalLogs}건`,
    { font: { name: FONT, sz: 9, italic: true, color: { rgb: C.muted } },
      alignment: { horizontal: 'right', vertical: 'center' } })
  R++

  // 빈 행
  R++

  // ── 3행: 헤더 ────────────────────────────────────────────────────────
  // 고정 헤더 (A~I)
  FIXED.forEach((col, c) => {
    wc(ws, R, c, col.label, {
      font:      { name: FONT, sz: 10, bold: true, color: { rgb: C.fixHdrFg } },
      fill:      { patternType: 'solid', fgColor: { rgb: C.fixHdrBg } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border:    borders('medium'),
    })
  })
  // 날짜 헤더 (1~31)
  for (let d = 1; d <= 31; d++) {
    const c       = NF + d - 1
    const inMonth = d <= lastDay
    wc(ws, R, c, d, {
      font:      { name: FONT, sz: 9, bold: true,
                   color: { rgb: inMonth ? C.dayHdrFg : C.dayHdrOffFg } },
      fill:      { patternType: 'solid',
                   fgColor: { rgb: inMonth ? C.dayHdrBg : C.dayHdrOffBg } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border:    borders('medium'),
    })
  }
  R++

  // ── 4행~: 데이터 행 ──────────────────────────────────────────────────
  reagents.forEach((rg, idx) => {
    const dm          = usageByRg.get(rg.id) ?? new Map()
    const monthlyUsed = [...dm.values()].reduce((s, v) => s + v, 0)
    const inboundQty  = monthlyInbound.get(rg.id) ?? 0
    // 전월이월량 = 현재재고 + 당월사용량 - 당월입고량
    const carriedOver = rg.currentStock + monthlyUsed - inboundQty
    const isLow       = rg.currentStock < rg.minStock
    const bg          = idx % 2 === 0 ? C.rowOdd : C.rowEven

    const ds = (align = 'left', fontEx = {}) => ({
      font:      { name: FONT, sz: 10, color: { rgb: '1E293B' }, ...fontEx },
      fill:      { patternType: 'solid', fgColor: { rgb: bg } },
      alignment: { horizontal: align, vertical: 'center' },
      border:    borders('thin'),
    })

    // 고정 열 값
    const fixedCells = [
      { v: rg.name,                                        align: 'left',   fx: {} },
      { v: rg.manufacturer,                                align: 'center', fx: {} },
      { v: rg.lotNo,                                       align: 'center', fx: {} },
      { v: rg.expiryDate ?? '',                            align: 'center', fx: {} },
      { v: rg.createdAt ? rg.createdAt.slice(0, 10) : '-', align: 'center', fx: {} },
      { v: inboundQty,                                     align: 'center', fx: {} },
      { v: carriedOver,                                    align: 'center', fx: {} },
      { v: rg.currentStock, align: 'center',
        fx: isLow ? { bold: true, color: { rgb: C.danger } } : { color: { rgb: C.safe } } },
      { v: monthlyUsed, align: 'center',
        fx: monthlyUsed > 0 ? { bold: true, color: { rgb: C.dayHdrFg } } : { color: { rgb: C.muted } } },
    ]
    fixedCells.forEach(({ v, align, fx }, c) => wc(ws, R, c, v, ds(align, fx)))

    // 날짜 열 (1~31)
    for (let d = 1; d <= 31; d++) {
      const c       = NF + d - 1
      const inMonth = d <= lastDay
      const qty     = dm.get(d)

      if (!inMonth) {
        wc(ws, R, c, '', {
          fill:   { patternType: 'solid', fgColor: { rgb: C.dayOffBg } },
          border: borders('thin'),
        })
      } else if (qty) {
        wc(ws, R, c, qty, {
          font:      { name: FONT, sz: 9, bold: true, color: { rgb: C.dayActiveFg } },
          fill:      { patternType: 'solid', fgColor: { rgb: C.dayActiveBg } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border:    borders('thin'),
        })
      } else {
        wc(ws, R, c, '', {
          fill:      { patternType: 'solid', fgColor: { rgb: bg } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border:    borders('thin'),
        })
      }
    }
    R++
  })

  // ── 합계 행 ────────────────────────────────────────────────────────
  const ts = (align = 'center') => ({
    font:      { name: FONT, sz: 10, bold: true, color: { rgb: C.totalFg } },
    fill:      { patternType: 'solid', fgColor: { rgb: C.totalBg } },
    alignment: { horizontal: align, vertical: 'center' },
    border:    borders('medium'),
  })

  // 고정 열 – 합계 행
  const grandTotal = [...usageByRg.values()].reduce(
    (s, dm) => s + [...dm.values()].reduce((a, b) => a + b, 0), 0
  )
  const summaryVals = [
    `총 ${reagents.length}종`, '', '', '', '', '', '', '', grandTotal,
  ]
  summaryVals.forEach((v, c) => wc(ws, R, c, v, ts(c === 0 ? 'center' : 'center')))

  // 날짜 열 – 합계 행
  for (let d = 1; d <= 31; d++) {
    const c       = NF + d - 1
    const inMonth = d <= lastDay
    if (inMonth) {
      const tot = dayTotals.get(d) ?? 0
      wc(ws, R, c, tot || '', tot ? ts() : {
        fill:   { patternType: 'solid', fgColor: { rgb: C.totalBg } },
        border: borders('medium'),
      })
    } else {
      wc(ws, R, c, '', {
        fill:   { patternType: 'solid', fgColor: { rgb: C.dayOffBg } },
        border: borders('thin'),
      })
    }
  }
  R++

  // ── 시트 메타데이터 ───────────────────────────────────────────────
  // 제목 행 열 병합 (A~마지막-1)
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: TCOL - 2 } },
  ]
  ws['!ref'] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: R - 1, c: TCOL - 1 })

  // 열 너비
  ws['!cols'] = [
    ...FIXED.map((col) => ({ wch: col.wch })),
    ...Array(31).fill(null).map(() => ({ wch: 4.5 })),
  ]

  // 행 높이
  ws['!rows'] = [
    { hpt: 30 }, // 타이틀
    { hpt: 5  }, // 빈 줄
    { hpt: 22 }, // 헤더
  ]

  // 틀 고정: 3행 + 9열(I열) 이후 고정 → 스크롤 시 시약명·헤더 고정
  ws['!freeze'] = { xSplit: NF, ySplit: 3 }

  XLSX.utils.book_append_sheet(wb, ws, '월간 시약 사용현황')
}
