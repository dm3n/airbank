import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

// ─── Interfaces ────────────────────────────────────────────────────────────────
interface WorkbookCell {
  section: string
  row_key: string
  period: string
  raw_value: number | null
  display_value: string | null
  is_overridden: boolean
  confidence: number | null
  source_excerpt: string | null
}
interface AuditEntry {
  edited_at: string
  old_value: number | null
  new_value: number | null
  note: string | null
  cell: { section: string; row_key: string; period: string } | null
}
interface WorkbookFlag {
  section: string
  row_key: string
  period: string | null
  flag_type: string
  severity: string
  title: string
  body: string | null
  created_by_ai: boolean
}

// ─── Number Formatters ─────────────────────────────────────────────────────────
function fmt$(v: number | null | undefined, compact = false): string {
  if (v == null) return '—'
  if (compact) {
    if (Math.abs(v) >= 1_000_000) return `$${(Math.abs(v) / 1e6).toFixed(1)}M`
    if (Math.abs(v) >= 1_000) return `$${(Math.abs(v) / 1e3).toFixed(0)}K`
  }
  const s = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(Math.abs(v))
  return v < 0 ? `(${s})` : s
}
function fmt$sign(v: number | null | undefined): string {
  if (v == null || v === 0) return '—'
  const s = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(Math.abs(v))
  return v > 0 ? `+${s}` : `(${s})`
}
function fmtPct(v: number | null | undefined, decimals = 1): string {
  if (v == null) return '—'
  return `${v.toFixed(decimals)}%`
}
function yoy(curr: number | null, prev: number | null): string {
  if (curr == null || prev == null || prev === 0) return '—'
  const p = ((curr - prev) / Math.abs(prev)) * 100
  return `${p >= 0 ? '+' : ''}${p.toFixed(1)}%`
}
function pctOfRev(val: number | null, rev: number | null): string {
  if (val == null || rev == null || rev === 0) return '—'
  return `${((val / rev) * 100).toFixed(1)}%`
}

// ─── Data Helpers ──────────────────────────────────────────────────────────────
function getV(cells: WorkbookCell[], period: string, ...keys: string[]): number | null {
  for (const k of keys) {
    const c = cells.find(x => x.period === period && x.row_key === k)
    if (c?.raw_value != null) return c.raw_value
  }
  return null
}
function humanize(s: string) {
  return s.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

// ─── Excel Builder ─────────────────────────────────────────────────────────────
async function buildExcel(
  companyName: string,
  cells: WorkbookCell[],
  auditEntries: AuditEntry[],
  periods: string[],
  flags: WorkbookFlag[],
): Promise<Buffer> {
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Airbank QoE Platform'
  wb.created = new Date()

  // Sort periods chronologically (TTM always last)
  const sortedPeriods = [...periods].sort((a, b) => {
    if (a === 'TTM') return 1
    if (b === 'TTM') return -1
    return a.localeCompare(b)
  })

  const ORANGE    = 'FFC86400'
  const DARK      = 'FF3D1C00'
  const NAVY      = 'FF1B2A4A'
  const BLUE_LIGHT = 'FFEFF6FF'
  const ORANGE_LIGHT = 'FFFFF3E0'
  const CREAM     = 'FFFBF7F0'
  const MID_GRAY  = 'FFE0E0E0'
  const TEXT      = 'FF1A1A1A'
  const GREEN_FILL = 'FFE8F5E9'
  const RED_FILL  = 'FFFFEBEE'
  const WHITE     = 'FFFFFFFF'
  const SUBTOTAL_FILL = 'FFFFF8E1'

  // Style helpers
  const hdrFont   = (sz = 10) => ({ bold: true, size: sz, color: { argb: WHITE }, name: 'Calibri' })
  const bodyFont  = (sz = 9, bold = false) => ({ bold, size: sz, name: 'Calibri', color: { argb: TEXT } })
  const fillSolid = (argb: string) => ({ type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb } })
  const thinBorder = { style: 'thin' as const, color: { argb: MID_GRAY } }
  const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder }
  const bottomBorder = { bottom: { style: 'medium' as const, color: { argb: ORANGE } } }

  function setHdrRow(row: InstanceType<typeof ExcelJS.Workbook>['worksheets'][0]['getRow'] extends (n: number) => infer R ? R : never) {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill = fillSolid(NAVY)
      cell.font = hdrFont(9)
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
      cell.border = { bottom: { style: 'medium', color: { argb: ORANGE } } }
    })
    row.height = 26
  }

  function setSubtotalRow(row: any, fillArgb = SUBTOTAL_FILL) {
    row.eachCell({ includeEmpty: true }, (cell: any) => {
      if (!cell.value && cell.value !== 0) return
      cell.fill = fillSolid(fillArgb)
      cell.font = bodyFont(9, true)
      cell.border = { top: { style: 'thin', color: { argb: ORANGE } }, bottom: { style: 'medium', color: { argb: ORANGE } } }
    })
  }

  function setSectionHdrRow(row: any) {
    row.eachCell({ includeEmpty: true }, (cell: any) => {
      cell.fill = fillSolid(ORANGE_LIGHT)
      cell.font = { bold: true, size: 8, name: 'Calibri', italic: true, color: { argb: DARK } }
    })
    row.height = 16
  }

  const ACCT_FMT = '#,##0_);(#,##0)'
  const ACCT_RED_FMT = '#,##0_);[Red](#,##0)'
  const PCT_FMT = '0.0%'
  const DELTA_FMT = '+#,##0;(#,##0);"-"'

  // ── Sheet 1: Cover ───────────────────────────────────────────────────────────
  const cover = wb.addWorksheet('Cover')
  cover.views = [{ showGridLines: false }]
  cover.properties.defaultRowHeight = 18

  const cCols = 8
  ;[10, 20, 16, 16, 16, 16, 16, 14].forEach((w, i) => { cover.getColumn(i + 1).width = w })

  const coverMerge = (r1: number, r2: number, v: any, style: Partial<InstanceType<typeof ExcelJS.Workbook>['worksheets'][0]['getCell'] extends (a: string) => infer R ? R : never> = {}) => {
    cover.mergeCells(r1, 1, r2, cCols)
    const cell = cover.getCell(r1, 1)
    cell.value = v
    Object.assign(cell, style)
  }

  // Title block
  cover.getRow(1).height = 8
  cover.mergeCells(2, 1, 4, cCols)
  const titleCell = cover.getCell(2, 1)
  titleCell.value = 'QUALITY OF EARNINGS REPORT'
  titleCell.fill = fillSolid(NAVY)
  titleCell.font = { bold: true, size: 16, name: 'Calibri', color: { argb: WHITE } }
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' }
  cover.getRow(2).height = 32; cover.getRow(3).height = 10; cover.getRow(4).height = 6

  cover.mergeCells(5, 1, 6, cCols)
  const compCell = cover.getCell(5, 1)
  compCell.value = companyName
  compCell.fill = fillSolid(ORANGE)
  compCell.font = { bold: true, size: 20, name: 'Calibri', color: { argb: WHITE } }
  compCell.alignment = { vertical: 'middle', horizontal: 'center' }
  cover.getRow(5).height = 38; cover.getRow(6).height = 6

  // Meta block
  const metaRows = [
    ['Report Date', new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })],
    ['Reporting Periods', sortedPeriods.join('  |  ')],
    ['Prepared By', 'Airbank Quality of Earnings Platform'],
    ['Report Type', 'Quality of Earnings Analysis — For Discussion Purposes Only'],
  ]
  metaRows.forEach(([label, value], i) => {
    const rn = 8 + i
    cover.mergeCells(rn, 1, rn, 2); cover.mergeCells(rn, 3, rn, cCols)
    const lCell = cover.getCell(rn, 1)
    const vCell = cover.getCell(rn, 3)
    lCell.value = label; vCell.value = value
    lCell.font = bodyFont(9, true); vCell.font = bodyFont(9)
    lCell.fill = fillSolid(CREAM); vCell.fill = fillSolid(WHITE)
    cover.getRow(rn).height = 20
  })

  cover.getRow(13).height = 12

  // Financial summary table
  cover.mergeCells(14, 1, 14, cCols)
  const sumHdr = cover.getCell(14, 1)
  sumHdr.value = 'FINANCIAL SUMMARY HIGHLIGHTS  (' + (sortedPeriods[sortedPeriods.length - 1] ?? 'Latest') + ')'
  sumHdr.fill = fillSolid(NAVY); sumHdr.font = hdrFont(10)
  sumHdr.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  cover.getRow(14).height = 24

  const latestP = sortedPeriods[sortedPeriods.length - 1]
  const prevP   = sortedPeriods.length >= 2 ? sortedPeriods[sortedPeriods.length - 2] : null
  const IS = cells.filter(c => c.section === 'income-statement')
  const QOE = cells.filter(c => c.section === 'qoe')

  const TTMRevenue = getV(IS, latestP, 'total_net_revenue')
  const TTMGrossProfit = getV(IS, latestP, 'gross_profit')
  const TTMEbitda = getV(QOE, latestP, 'diligence_adjusted_ebitda') ?? getV(QOE, latestP, 'ebitda_as_defined') ?? getV(IS, latestP, 'operating_income_ebitda')
  const TTMNetIncome = getV(IS, latestP, 'net_income') ?? getV(QOE, latestP, 'net_income')
  const grossMarginPct = TTMRevenue && TTMGrossProfit ? (TTMGrossProfit / TTMRevenue) * 100 : null
  const ebitdaMarginPct = TTMRevenue && TTMEbitda ? (TTMEbitda / TTMRevenue) * 100 : null
  const prevRevenue = prevP ? getV(IS, prevP, 'total_net_revenue') : null
  const revGrowth = TTMRevenue && prevRevenue && prevRevenue !== 0 ? ((TTMRevenue - prevRevenue) / Math.abs(prevRevenue)) * 100 : null

  const sumData = [
    ['Total Revenue', TTMRevenue, revGrowth != null ? `${revGrowth >= 0 ? '+' : ''}${revGrowth.toFixed(1)}% YoY` : '—'],
    ['Gross Profit', TTMGrossProfit, grossMarginPct != null ? `${grossMarginPct.toFixed(1)}% Gross Margin` : '—'],
    ['Diligence-Adjusted EBITDA', TTMEbitda, ebitdaMarginPct != null ? `${ebitdaMarginPct.toFixed(1)}% EBITDA Margin` : '—'],
    ['Net Income', TTMNetIncome, ''],
    ['Open Flags', flags.length, flags.filter(f => f.severity === 'critical').length + ' Critical'],
    ['Analyst Overrides', auditEntries.length, 'QoE Adjustments'],
  ]

  sumData.forEach(([label, value, note], i) => {
    const rn = 15 + i
    cover.mergeCells(rn, 1, rn, 2); cover.mergeCells(rn, 3, rn, 5); cover.mergeCells(rn, 6, rn, cCols)
    const l = cover.getCell(rn, 1), v = cover.getCell(rn, 3), n = cover.getCell(rn, 6)
    l.value = String(label); v.value = typeof value === 'number' ? value : value; n.value = String(note)
    l.fill = fillSolid(i % 2 === 0 ? CREAM : WHITE); v.fill = fillSolid(i % 2 === 0 ? CREAM : WHITE); n.fill = fillSolid(i % 2 === 0 ? CREAM : WHITE)
    l.font = bodyFont(9, true); v.font = bodyFont(10, true); n.font = bodyFont(8)
    if (typeof value === 'number' && i < 4) v.numFmt = ACCT_FMT
    n.font = { ...n.font, color: { argb: 'FF888888' } }
    cover.getRow(rn).height = 22
  })

  cover.getRow(22).height = 14

  // Disclaimer
  cover.mergeCells(23, 1, 26, cCols)
  const disc = cover.getCell(23, 1)
  disc.value = 'CONFIDENTIAL\n\nThis Quality of Earnings report has been prepared by Airbank Platform for internal use and the exclusive benefit of the named company. The information contained herein is based on data provided by management and has not been independently audited or verified. This report is prepared for discussion purposes only and does not constitute a formal audit, review, or compilation. It may not be distributed, reproduced, or used without prior written consent.'
  disc.fill = fillSolid('FFFFF9E8')
  disc.font = { size: 8, name: 'Calibri', color: { argb: 'FF6B5C4A' }, italic: true }
  disc.alignment = { wrapText: true, vertical: 'top' }
  disc.border = { top: { style: 'medium', color: { argb: ORANGE } }, bottom: { style: 'thin', color: { argb: MID_GRAY } }, left: { style: 'thin', color: { argb: MID_GRAY } }, right: { style: 'thin', color: { argb: MID_GRAY } } }
  cover.getRow(23).height = 60

  // ── Sheet 2: QoE Summary ─────────────────────────────────────────────────────
  const qoeS = wb.addWorksheet('QoE Summary')
  qoeS.views = [{ state: 'frozen', ySplit: 2 }]
  ;[38, ...sortedPeriods.map(() => 16), 16, 14].forEach((w, i) => { qoeS.getColumn(i + 1).width = w })

  const qHdrRow = qoeS.getRow(1)
  qHdrRow.getCell(1).value = `${companyName} — Quality of Earnings Summary`
  qHdrRow.getCell(1).font = { bold: true, size: 12, name: 'Calibri', color: { argb: NAVY } }
  qHdrRow.getCell(1).fill = fillSolid(CREAM)
  qoeS.mergeCells(1, 1, 1, 2 + sortedPeriods.length)
  qHdrRow.height = 22

  const qColHdr = qoeS.getRow(2)
  qColHdr.getCell(1).value = 'Metric'
  sortedPeriods.forEach((p, i) => { qColHdr.getCell(2 + i).value = p })
  qColHdr.getCell(2 + sortedPeriods.length).value = 'YoY Growth'
  setHdrRow(qColHdr as any)

  let qr = 3
  const addQoeSection = (label: string) => {
    const row = qoeS.getRow(qr++)
    row.getCell(1).value = label
    setSectionHdrRow(row as any)
    qoeS.mergeCells(qr - 1, 1, qr - 1, 2 + sortedPeriods.length)
  }

  const addQoeRow = (label: string, rowKey: string, section: string, isSubtotal = false, isBold = false) => {
    const row = qoeS.getRow(qr++)
    row.getCell(1).value = label
    row.getCell(1).font = bodyFont(9, isBold || isSubtotal)
    row.getCell(1).fill = fillSolid(isSubtotal ? SUBTOTAL_FILL : (qr % 2 === 0 ? BLUE_LIGHT : WHITE))
    const sectionCells = cells.filter(c => c.section === section)
    sortedPeriods.forEach((p, i) => {
      const val = getV(sectionCells, p, rowKey)
      const cell = row.getCell(2 + i)
      cell.value = val
      cell.numFmt = ACCT_FMT
      cell.font = bodyFont(9, isBold || isSubtotal)
      cell.fill = fillSolid(isSubtotal ? SUBTOTAL_FILL : (qr % 2 === 0 ? BLUE_LIGHT : WHITE))
      cell.alignment = { horizontal: 'right' }
    })
    // YoY Growth
    const last2 = sortedPeriods.slice(-2)
    if (last2.length === 2) {
      const curr = getV(cells.filter(c => c.section === section), last2[1], rowKey)
      const prev = getV(cells.filter(c => c.section === section), last2[0], rowKey)
      const yoyCell = row.getCell(2 + sortedPeriods.length)
      if (curr != null && prev != null && prev !== 0) {
        const pct = ((curr - prev) / Math.abs(prev))
        yoyCell.value = pct
        yoyCell.numFmt = '+0.0%;-0.0%;"-"'
        yoyCell.font = { ...bodyFont(9), color: { argb: pct >= 0 ? 'FF2E7D32' : 'FFC62828' } }
      } else {
        yoyCell.value = '—'
        yoyCell.font = bodyFont(9)
      }
      yoyCell.fill = fillSolid(isSubtotal ? SUBTOTAL_FILL : (qr % 2 === 0 ? BLUE_LIGHT : WHITE))
    }
    if (isSubtotal) setSubtotalRow(row as any, SUBTOTAL_FILL)
    row.height = 16
    return row
  }

  addQoeSection('REVENUE')
  addQoeRow('Total Net Revenue', 'total_net_revenue', 'income-statement', false, false)
  addQoeRow('Gross Profit', 'gross_profit', 'income-statement', true, true)
  addQoeRow('Gross Margin %', 'gross_profit', 'income-statement', false, false) // will override with pct

  // Override Gross Margin % row with computed pct values
  const gmRow = qoeS.getRow(qr - 1)
  gmRow.getCell(1).value = 'Gross Margin %'
  sortedPeriods.forEach((p, i) => {
    const rev = getV(IS, p, 'total_net_revenue')
    const gp  = getV(IS, p, 'gross_profit')
    const cell = gmRow.getCell(2 + i)
    cell.value = (rev && gp) ? gp / rev : null
    cell.numFmt = PCT_FMT
    cell.font = { ...bodyFont(9), italic: true }
  })

  addQoeSection('QUALITY OF EARNINGS BRIDGE')
  addQoeRow('Net Income', 'net_income', 'qoe')
  addQoeRow('+ Interest Expense', 'interest_expense', 'qoe')
  addQoeRow('+ Tax Provision', 'tax_provision', 'qoe')
  addQoeRow('+ Depreciation', 'depreciation', 'qoe')
  addQoeRow('+ Amortization', 'amortization', 'qoe')
  addQoeRow('= EBITDA, as Defined', 'ebitda_as_defined', 'qoe', true, true)
  addQoeRow('Management Adjustments (Total)', 'total_mgmt_adjustments', 'qoe')
  addQoeRow('= Management-Adjusted EBITDA', 'mgmt_adjusted_ebitda', 'qoe', true, true)
  addQoeRow('Diligence Adjustments (Total)', 'total_diligence_adjustments', 'qoe')
  addQoeRow('= Diligence-Adjusted EBITDA', 'diligence_adjusted_ebitda', 'qoe', true, true)

  // Adjusted EBITDA Margin %
  const adjMarginRow = qoeS.getRow(qr++)
  adjMarginRow.getCell(1).value = 'Adjusted EBITDA Margin %'
  adjMarginRow.getCell(1).font = { ...bodyFont(9), italic: true }
  sortedPeriods.forEach((p, i) => {
    const rev = getV(IS, p, 'total_net_revenue')
    const ebitda = getV(QOE, p, 'diligence_adjusted_ebitda')
    const cell = adjMarginRow.getCell(2 + i)
    cell.value = (rev && ebitda) ? ebitda / rev : null
    cell.numFmt = PCT_FMT
    cell.font = { ...bodyFont(9), italic: true }
    cell.fill = fillSolid(ORANGE_LIGHT)
  })
  adjMarginRow.getCell(1).fill = fillSolid(ORANGE_LIGHT)
  adjMarginRow.height = 16

  addQoeSection('BALANCE SHEET SNAPSHOT')
  addQoeRow('Total Current Assets', 'total_current_assets', 'balance-sheet')
  addQoeRow('Total Current Liabilities', 'total_current_liabilities', 'balance-sheet')
  addQoeRow('Net Working Capital', 'total_current_assets', 'balance-sheet', true, true) // override below
  const nwcRow = qoeS.getRow(qr - 1)
  nwcRow.getCell(1).value = 'Net Working Capital'
  sortedPeriods.forEach((p, i) => {
    const ca = getV(cells.filter(c => c.section === 'balance-sheet'), p, 'total_current_assets')
    const cl = getV(cells.filter(c => c.section === 'balance-sheet'), p, 'total_current_liabilities')
    const cell = nwcRow.getCell(2 + i)
    cell.value = (ca != null && cl != null) ? ca - cl : null
    cell.numFmt = ACCT_FMT
    cell.font = bodyFont(9, true)
    cell.fill = fillSolid(SUBTOTAL_FILL)
  })

  addQoeRow('Total Assets', 'total_assets', 'balance-sheet')
  addQoeRow('Total Liabilities', 'total_liabilities', 'balance-sheet')
  addQoeRow("Total Stockholders' Equity", 'total_equity', 'balance-sheet', true, true)

  // Auto-filter
  qoeS.autoFilter = { from: 'A2', to: { row: 2, column: 2 + sortedPeriods.length } }

  // ── Sheet 3: Income Statement ────────────────────────────────────────────────
  const isWS = wb.addWorksheet('Income Statement')
  isWS.views = [{ state: 'frozen', ySplit: 3 }]
  ;[38, ...sortedPeriods.map(() => 16), 14, 14].forEach((w, i) => { isWS.getColumn(i + 1).width = w })

  // Title
  isWS.mergeCells(1, 1, 1, 2 + sortedPeriods.length + 1)
  const isTitleCell = isWS.getCell(1, 1)
  isTitleCell.value = `${companyName} — Income Statement (Adjusted)`
  isTitleCell.font = { bold: true, size: 12, name: 'Calibri', color: { argb: NAVY } }
  isTitleCell.fill = fillSolid(CREAM)
  isWS.getRow(1).height = 22

  // Period labels row
  isWS.mergeCells(2, 1, 2, 2 + sortedPeriods.length + 1)
  const isSubTitle = isWS.getCell(2, 1)
  isSubTitle.value = `Reporting Periods: ${sortedPeriods.join(' | ')}  |  All figures in USD`
  isSubTitle.font = bodyFont(8); isSubTitle.fill = fillSolid(CREAM)
  isWS.getRow(2).height = 16

  // Column headers
  const isHdr = isWS.getRow(3)
  isHdr.getCell(1).value = 'Line Item'
  sortedPeriods.forEach((p, i) => { isHdr.getCell(2 + i).value = p })
  isHdr.getCell(2 + sortedPeriods.length).value = '% of Rev (TTM)'
  isHdr.getCell(3 + sortedPeriods.length).value = 'YoY Growth'
  setHdrRow(isHdr as any)
  isHdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }

  let isr = 4
  const ttmRev = getV(IS, latestP, 'total_net_revenue')
  const prevRev = prevP ? getV(IS, prevP, 'total_net_revenue') : null

  // Income statement sections
  const IS_SECTIONS = [
    { header: 'REVENUE', rows: [
      { k: 'gross_product_sales', l: 'Gross Product Sales', sub: false },
      { k: 'returns_allowances', l: 'Returns & Allowances', sub: false },
      { k: 'promotional_discounts', l: 'Promotional Discounts', sub: false },
      { k: 'net_product_sales', l: 'Net Product Sales', sub: true },
      { k: 'shipping_revenue', l: 'Shipping Revenue', sub: false },
      { k: 'other_revenue', l: 'Other Revenue', sub: false },
      { k: 'total_net_revenue', l: 'Total Net Revenue', sub: true },
    ]},
    { header: 'COST OF GOODS SOLD', rows: [
      { k: 'product_cogs', l: 'Product Cost of Goods Sold', sub: false },
      { k: 'warehousing_fulfillment', l: 'Warehousing & Fulfillment', sub: false },
      { k: 'shipping_freight_out', l: 'Shipping & Freight Out', sub: false },
      { k: 'inventory_adjustments', l: 'Inventory Adjustments', sub: false },
      { k: 'total_cogs', l: 'Total Cost of Goods Sold', sub: true },
    ]},
    { header: 'GROSS PROFIT', rows: [
      { k: 'gross_profit', l: 'Gross Profit', sub: true },
    ]},
    { header: 'OPERATING EXPENSES', rows: [
      { k: 'sales_marketing', l: 'Sales & Marketing', sub: false },
      { k: 'salaries_wages_payroll', l: 'Salaries, Wages & Payroll Tax', sub: false },
      { k: 'employee_benefits_401k', l: 'Employee Benefits & 401k', sub: false },
      { k: 'rent_occupancy', l: 'Rent & Occupancy', sub: false },
      { k: 'professional_services', l: 'Professional Services', sub: false },
      { k: 'technology_software', l: 'Technology & Software', sub: false },
      { k: 'insurance', l: 'Insurance', sub: false },
      { k: 'payment_processing_fees', l: 'Payment Processing Fees', sub: false },
      { k: 'travel_entertainment', l: 'Travel & Entertainment', sub: false },
      { k: 'office_general_admin', l: 'Office & General Admin', sub: false },
      { k: 'total_operating_expenses', l: 'Total Operating Expenses', sub: true },
    ]},
    { header: 'EBITDA & BELOW', rows: [
      { k: 'operating_income_ebitda', l: 'Operating Income (EBITDA)', sub: true },
      { k: 'interest_expense', l: 'Interest Expense', sub: false },
      { k: 'depreciation_amortization', l: 'Depreciation & Amortization', sub: false },
      { k: 'income_before_tax', l: 'Income Before Tax', sub: true },
      { k: 'tax_provision', l: 'Tax Provision', sub: false },
      { k: 'net_income', l: 'Net Income', sub: true },
    ]},
  ]

  IS_SECTIONS.forEach(sec => {
    // Section header
    const shRow = isWS.getRow(isr++)
    shRow.getCell(1).value = sec.header
    setSectionHdrRow(shRow as any)
    isWS.mergeCells(isr - 1, 1, isr - 1, 3 + sortedPeriods.length)

    sec.rows.forEach(({ k, l, sub }) => {
      const row = isWS.getRow(isr++)
      const isAlt = isr % 2 === 0
      row.getCell(1).value = l
      row.getCell(1).font = bodyFont(9, sub)
      row.getCell(1).fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))

      sortedPeriods.forEach((p, pi) => {
        const val = getV(IS, p, k)
        const c = row.getCell(2 + pi)
        c.value = val; c.numFmt = ACCT_RED_FMT
        c.font = bodyFont(9, sub)
        c.fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))
        c.alignment = { horizontal: 'right' }
      })

      // % of Rev (TTM)
      const ttmVal = getV(IS, latestP, k)
      const pctCell = row.getCell(2 + sortedPeriods.length)
      if (ttmVal != null && ttmRev && ttmRev !== 0) {
        pctCell.value = ttmVal / ttmRev
        pctCell.numFmt = PCT_FMT
      } else { pctCell.value = null }
      pctCell.font = { ...bodyFont(8), italic: true }
      pctCell.fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))

      // YoY Growth
      const yoyCell = row.getCell(3 + sortedPeriods.length)
      if (prevP) {
        const c = getV(IS, latestP, k), p = getV(IS, prevP, k)
        if (c != null && p != null && p !== 0) {
          const pct = (c - p) / Math.abs(p)
          yoyCell.value = pct
          yoyCell.numFmt = '+0.0%;-0.0%;"-"'
          yoyCell.font = { ...bodyFont(8), color: { argb: pct >= 0 ? 'FF2E7D32' : 'FFC62828' } }
        }
      }
      yoyCell.fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))

      if (sub) setSubtotalRow(row as any)
      row.height = 16
    })
  })

  isWS.autoFilter = { from: 'A3', to: { row: 3, column: 3 + sortedPeriods.length } }

  // ── Sheet 4: EBITDA Bridge ───────────────────────────────────────────────────
  const bridgeWS = wb.addWorksheet('EBITDA Bridge')
  bridgeWS.views = [{ state: 'frozen', ySplit: 2 }]
  ;[42, 18, ...sortedPeriods.map(() => 16)].forEach((w, i) => { bridgeWS.getColumn(i + 1).width = w })

  bridgeWS.mergeCells(1, 1, 1, 2 + sortedPeriods.length)
  const bTitle = bridgeWS.getCell(1, 1)
  bTitle.value = `${companyName} — EBITDA Bridge & Quality of Earnings`
  bTitle.font = { bold: true, size: 12, name: 'Calibri', color: { argb: NAVY } }
  bTitle.fill = fillSolid(CREAM); bridgeWS.getRow(1).height = 22

  const bHdr = bridgeWS.getRow(2)
  bHdr.getCell(1).value = 'Line Item'
  bHdr.getCell(2).value = 'Type'
  sortedPeriods.forEach((p, i) => { bHdr.getCell(3 + i).value = p })
  setHdrRow(bHdr as any); bHdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }

  let br = 3
  const addBridgeRow = (label: string, type: string, rowKey: string, section: string, style: 'normal' | 'subtotal' | 'headline' | 'section' = 'normal') => {
    const row = bridgeWS.getRow(br++)
    const fill = style === 'headline' ? ORANGE : style === 'subtotal' ? SUBTOTAL_FILL : style === 'section' ? ORANGE_LIGHT : (br % 2 === 0 ? BLUE_LIGHT : WHITE)
    const fontColor = style === 'headline' ? WHITE : TEXT
    row.getCell(1).value = label
    row.getCell(1).font = { bold: style !== 'normal', size: 9, name: 'Calibri', color: { argb: fontColor } }
    row.getCell(1).fill = fillSolid(fill)
    row.getCell(2).value = type
    row.getCell(2).font = { size: 8, name: 'Calibri', italic: true, color: { argb: type === 'Add-back' ? 'FF2E7D32' : type === 'Reduction' ? 'FFC62828' : fontColor } }
    row.getCell(2).fill = fillSolid(fill)
    const sectionCells = cells.filter(c => c.section === section)
    sortedPeriods.forEach((p, i) => {
      const c = row.getCell(3 + i)
      c.value = getV(sectionCells, p, rowKey)
      c.numFmt = ACCT_FMT
      c.font = { bold: style !== 'normal', size: 9, name: 'Calibri', color: { argb: fontColor } }
      c.fill = fillSolid(fill)
      c.alignment = { horizontal: 'right' }
    })
    if (style === 'subtotal' || style === 'headline') setSubtotalRow(row as any, fill)
    row.height = style === 'headline' ? 22 : 16
  }

  const addBridgeSep = (label: string) => {
    const row = bridgeWS.getRow(br++)
    bridgeWS.mergeCells(br - 1, 1, br - 1, 2 + sortedPeriods.length)
    row.getCell(1).value = label; setSectionHdrRow(row as any); row.height = 16
  }

  addBridgeSep('STARTING POINT')
  addBridgeRow('Net Income', 'Basis', 'net_income', 'qoe')
  addBridgeSep('EBITDA ADD-BACKS')
  addBridgeRow('+ Interest Expense', 'Add-back', 'interest_expense', 'qoe')
  addBridgeRow('+ Tax Provision', 'Add-back', 'tax_provision', 'qoe')
  addBridgeRow('+ Depreciation', 'Add-back', 'depreciation', 'qoe')
  addBridgeRow('+ Amortization', 'Add-back', 'amortization', 'qoe')
  addBridgeRow('= EBITDA, as Defined', 'Subtotal', 'ebitda_as_defined', 'qoe', 'subtotal')
  addBridgeSep('MANAGEMENT ADJUSTMENTS')
  addBridgeRow('a) Charitable Contributions', 'Add-back', 'adj_charitable_contributions', 'qoe')
  addBridgeRow('b) Owner Tax & Legal', 'Add-back', 'adj_owner_tax_legal', 'qoe')
  addBridgeRow('c) Excess Owner Compensation', 'Add-back', 'adj_excess_owner_comp', 'qoe')
  addBridgeRow('d) Personal Expenses', 'Add-back', 'adj_personal_expenses', 'qoe')
  addBridgeRow('= Total Management Adjustments', 'Subtotal', 'total_mgmt_adjustments', 'qoe', 'subtotal')
  addBridgeRow('= Management-Adjusted EBITDA', 'Subtotal', 'mgmt_adjusted_ebitda', 'qoe', 'subtotal')
  addBridgeSep('DILIGENCE ADJUSTMENTS')
  addBridgeRow('e) Professional Fees - One-time', 'Add-back', 'adj_professional_fees_onetime', 'qoe')
  addBridgeRow('f) Executive Search Fees', 'Add-back', 'adj_executive_search', 'qoe')
  addBridgeRow('g) Facility Relocation', 'Add-back', 'adj_facility_relocation', 'qoe')
  addBridgeRow('h) Severance', 'Add-back', 'adj_severance', 'qoe')
  addBridgeRow('i) Above-Market Rent', 'Reduction', 'adj_above_market_rent', 'qoe')
  addBridgeRow('j) Normalize Owner Compensation', 'Reduction', 'adj_normalize_owner_comp', 'qoe')
  addBridgeRow('= Total Diligence Adjustments', 'Subtotal', 'total_diligence_adjustments', 'qoe', 'subtotal')
  bridgeWS.getRow(br++).height = 6
  addBridgeRow('⭐  DILIGENCE-ADJUSTED EBITDA', 'Headline', 'diligence_adjusted_ebitda', 'qoe', 'headline')

  // Margin row
  const marginBRow = bridgeWS.getRow(br++)
  bridgeWS.mergeCells(br - 1, 1, br - 1, 2)
  marginBRow.getCell(1).value = 'Adjusted EBITDA Margin %'
  marginBRow.getCell(1).font = { size: 9, name: 'Calibri', italic: true, color: { argb: 'FF6B5C4A' } }
  sortedPeriods.forEach((p, i) => {
    const rev = getV(IS, p, 'total_net_revenue')
    const e = getV(QOE, p, 'diligence_adjusted_ebitda')
    const c = marginBRow.getCell(3 + i)
    c.value = (rev && e) ? e / rev : null; c.numFmt = PCT_FMT
    c.font = { size: 9, name: 'Calibri', italic: true }; c.alignment = { horizontal: 'right' }
  })
  marginBRow.height = 16

  // ── Sheet 5: Adjustments Schedule ───────────────────────────────────────────
  const adjWS = wb.addWorksheet('Adjustments Schedule')
  adjWS.views = [{ state: 'frozen', ySplit: 2 }]
  ;[22, 22, 36, 10, 16, 16, 16, 12, 40].forEach((w, i) => { adjWS.getColumn(i + 1).width = w })

  adjWS.mergeCells(1, 1, 1, 9)
  const adjTitle = adjWS.getCell(1, 1)
  adjTitle.value = `${companyName} — Adjustments Schedule (All Analyst Overrides)`
  adjTitle.font = { bold: true, size: 12, name: 'Calibri', color: { argb: NAVY } }
  adjTitle.fill = fillSolid(CREAM); adjWS.getRow(1).height = 22

  const adjHdr = adjWS.getRow(2)
  ;['Date', 'Section', 'Line Item', 'Period', 'Reported', 'Adjusted', 'Change ($)', 'Change (%)', 'Analyst Note'].forEach((h, i) => {
    adjHdr.getCell(i + 1).value = h
  })
  setHdrRow(adjHdr as any); adjHdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }

  // Build adjustments from audit entries (first reported = earliest old_value)
  const adjByKey = new Map<string, { section: string; rowKey: string; period: string; reported: number | null; adjusted: number | null; note: string | null; date: string }>()
  for (const e of [...auditEntries].reverse()) {
    if (!e.cell) continue
    const k = `${e.cell.section}::${e.cell.row_key}::${e.cell.period}`
    if (!adjByKey.has(k)) {
      adjByKey.set(k, { section: e.cell.section, rowKey: e.cell.row_key, period: e.cell.period, reported: e.old_value, adjusted: e.new_value, note: e.note, date: e.edited_at })
    } else {
      const ex = adjByKey.get(k)!; ex.adjusted = e.new_value; ex.note = e.note ?? ex.note
    }
  }

  let ar = 3
  for (const adj of adjByKey.values()) {
    const row = adjWS.getRow(ar++)
    const isAlt = ar % 2 === 0
    const delta = adj.adjusted != null && adj.reported != null ? adj.adjusted - adj.reported : null
    const deltaPct = delta != null && adj.reported != null && adj.reported !== 0 ? delta / Math.abs(adj.reported) : null

    row.getCell(1).value = new Date(adj.date).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })
    row.getCell(2).value = humanize(adj.section)
    row.getCell(3).value = humanize(adj.rowKey)
    row.getCell(4).value = adj.period
    row.getCell(5).value = adj.reported; row.getCell(5).numFmt = ACCT_FMT
    row.getCell(6).value = adj.adjusted; row.getCell(6).numFmt = ACCT_FMT
    row.getCell(7).value = delta; row.getCell(7).numFmt = DELTA_FMT
    if (delta != null) row.getCell(7).font = { ...bodyFont(9), color: { argb: delta >= 0 ? 'FF2E7D32' : 'FFC62828' } }
    row.getCell(8).value = deltaPct; row.getCell(8).numFmt = '+0.0%;-0.0%;"-"'
    if (deltaPct != null) row.getCell(8).font = { ...bodyFont(9), color: { argb: deltaPct >= 0 ? 'FF2E7D32' : 'FFC62828' } }
    row.getCell(9).value = adj.note ?? ''

    const fill = isAlt ? BLUE_LIGHT : WHITE
    row.eachCell({ includeEmpty: false }, c => { if (!c.fill || (c.fill as any).type === 'none') (c as any).fill = fillSolid(fill) })
    row.eachCell({ includeEmpty: false }, c => { if (!c.font) c.font = bodyFont(9) })
    row.height = 16
  }

  if (ar === 3) {
    adjWS.mergeCells(3, 1, 3, 9)
    adjWS.getCell(3, 1).value = 'No analyst overrides recorded for this workbook.'
    adjWS.getCell(3, 1).font = { italic: true, size: 9, name: 'Calibri', color: { argb: '88888888' } }
  }

  adjWS.autoFilter = { from: 'A2', to: 'I2' }

  // ── Sheet 6: Balance Sheet ───────────────────────────────────────────────────
  const bsWS = wb.addWorksheet('Balance Sheet')
  bsWS.views = [{ state: 'frozen', ySplit: 2 }]
  ;[40, ...sortedPeriods.map(() => 16)].forEach((w, i) => { bsWS.getColumn(i + 1).width = w })

  bsWS.mergeCells(1, 1, 1, 1 + sortedPeriods.length)
  const bsTitle = bsWS.getCell(1, 1)
  bsTitle.value = `${companyName} — Balance Sheet`
  bsTitle.font = { bold: true, size: 12, name: 'Calibri', color: { argb: NAVY } }
  bsTitle.fill = fillSolid(CREAM); bsWS.getRow(1).height = 22

  const bsHdr = bsWS.getRow(2)
  bsHdr.getCell(1).value = 'Line Item'
  sortedPeriods.forEach((p, i) => { bsHdr.getCell(2 + i).value = p })
  setHdrRow(bsHdr as any); bsHdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }

  const BS = cells.filter(c => c.section === 'balance-sheet')
  let bsr = 3

  const BS_SECTIONS = [
    { header: 'CURRENT ASSETS', rows: [
      { k: 'cash_equivalents', l: 'Cash & Cash Equivalents' },
      { k: 'accounts_receivable', l: 'Accounts Receivable' },
      { k: 'allowance_doubtful', l: 'Allowance for Doubtful Accounts' },
      { k: 'total_inventory', l: 'Total Inventory' },
      { k: 'prepaid_expenses', l: 'Prepaid Expenses' },
      { k: 'other_current_assets', l: 'Other Current Assets' },
      { k: 'total_current_assets', l: 'Total Current Assets', sub: true },
    ]},
    { header: 'LONG-TERM ASSETS', rows: [
      { k: 'ppe_gross', l: 'Property, Plant & Equipment' },
      { k: 'accumulated_depreciation', l: 'Less: Accumulated Depreciation' },
      { k: 'net_fixed_assets', l: 'Net Fixed Assets', sub: true },
      { k: 'intangible_assets', l: 'Intangible Assets' },
      { k: 'total_assets', l: 'Total Assets', sub: true },
    ]},
    { header: 'CURRENT LIABILITIES', rows: [
      { k: 'accounts_payable_trade', l: 'Accounts Payable' },
      { k: 'credit_cards_payable', l: 'Credit Cards Payable' },
      { k: 'accrued_payroll_benefits', l: 'Accrued Payroll & Benefits' },
      { k: 'accrued_expenses_other', l: 'Accrued Expenses - Other' },
      { k: 'sales_tax_payable', l: 'Sales Tax Payable' },
      { k: 'deferred_revenue', l: 'Deferred Revenue' },
      { k: 'current_portion_lt_debt', l: 'Current Portion - LT Debt' },
      { k: 'total_current_liabilities', l: 'Total Current Liabilities', sub: true },
    ]},
    { header: "LONG-TERM LIABILITIES & EQUITY", rows: [
      { k: 'lt_debt_net', l: 'Long-Term Debt, Net of Current' },
      { k: 'total_liabilities', l: 'Total Liabilities', sub: true },
      { k: 'common_stock', l: 'Common Stock' },
      { k: 'additional_paid_in_capital', l: 'Additional Paid-in Capital' },
      { k: 'retained_earnings', l: 'Retained Earnings' },
      { k: 'current_year_net_income', l: 'Current Year Net Income' },
      { k: 'total_equity', l: "Total Stockholders' Equity", sub: true },
      { k: 'total_liabilities_equity', l: 'Total Liabilities & Equity', sub: true },
    ]},
  ]

  BS_SECTIONS.forEach(sec => {
    const shRow = bsWS.getRow(bsr++)
    shRow.getCell(1).value = sec.header; setSectionHdrRow(shRow as any)
    bsWS.mergeCells(bsr - 1, 1, bsr - 1, 1 + sortedPeriods.length)

    sec.rows.forEach(({ k, l, sub }: any) => {
      const row = bsWS.getRow(bsr++)
      const isAlt = bsr % 2 === 0
      row.getCell(1).value = l; row.getCell(1).font = bodyFont(9, sub ?? false)
      row.getCell(1).fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))
      sortedPeriods.forEach((p, i) => {
        const c = row.getCell(2 + i)
        c.value = getV(BS, p, k); c.numFmt = ACCT_RED_FMT
        c.font = bodyFont(9, sub ?? false)
        c.fill = fillSolid(sub ? SUBTOTAL_FILL : (isAlt ? BLUE_LIGHT : WHITE))
        c.alignment = { horizontal: 'right' }
      })
      if (sub) setSubtotalRow(row as any)
      row.height = 16
    })
  })

  // ── Sheet 7: Audit Trail ─────────────────────────────────────────────────────
  const auditWS = wb.addWorksheet('Audit Trail')
  auditWS.views = [{ state: 'frozen', ySplit: 1 }]
  ;[22, 20, 36, 10, 15, 15, 15, 40].forEach((w, i) => { auditWS.getColumn(i + 1).width = w })

  const auditHdr = auditWS.getRow(1)
  ;['Date / Time', 'Section', 'Line Item', 'Period', 'Old Value', 'New Value', 'Change ($)', 'Analyst Note'].forEach((h, i) => {
    auditHdr.getCell(i + 1).value = h
  })
  setHdrRow(auditHdr as any); auditHdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }

  const sortedAudit = [...auditEntries].sort((a, b) => new Date(b.edited_at).getTime() - new Date(a.edited_at).getTime())
  sortedAudit.forEach((e, i) => {
    const row = auditWS.getRow(i + 2)
    const isAlt = (i + 2) % 2 === 0
    const delta = e.new_value != null && e.old_value != null ? e.new_value - e.old_value : null
    row.getCell(1).value = new Date(e.edited_at).toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' })
    row.getCell(2).value = e.cell ? humanize(e.cell.section) : ''
    row.getCell(3).value = e.cell ? humanize(e.cell.row_key) : ''
    row.getCell(4).value = e.cell?.period ?? ''
    row.getCell(5).value = e.old_value; row.getCell(5).numFmt = ACCT_FMT
    row.getCell(6).value = e.new_value; row.getCell(6).numFmt = ACCT_FMT
    row.getCell(7).value = delta; row.getCell(7).numFmt = DELTA_FMT
    if (delta != null) row.getCell(7).font = { ...bodyFont(9), color: { argb: delta >= 0 ? 'FF2E7D32' : 'FFC62828' } }
    row.getCell(8).value = e.note ?? ''
    const fill = isAlt ? BLUE_LIGHT : WHITE
    row.eachCell({ includeEmpty: false }, c => { if (!c.font) c.font = bodyFont(9); (c as any).fill = fillSolid(fill) })
    row.height = 16
  })
  auditWS.autoFilter = { from: 'A1', to: 'H1' }

  return Buffer.from(await wb.xlsx.writeBuffer())
}

// ─── PDF Builder ───────────────────────────────────────────────────────────────
async function buildPdf(
  companyName: string,
  cells: WorkbookCell[],
  auditEntries: AuditEntry[],
  flags: WorkbookFlag[],
  periods: string[],
): Promise<Buffer> {
  const PDFDocument = (await import('pdfkit')).default

  const W = 612, H = 792, M = 48, UW = 516

  // Color palette (Peermount-inspired warm professional)
  const ORANGE   = '#C17A2E'
  const DARK     = '#2B1200'
  const OL       = '#FDF5E4'   // orange light fill
  const TABLEHDR = '#F0E4CC'
  const ALTROW   = '#FAF6F0'
  const SEP      = '#D4B896'
  const TEXT     = '#1A1A1A'
  const GRAY     = '#6B5C4A'
  const LG       = '#F5F5F5'
  const WHITE    = '#FFFFFF'
  const GREEN    = '#2E7D32'
  const RED      = '#C62828'

  const sortedPeriods = [...periods].sort((a, b) => {
    if (a === 'TTM') return 1; if (b === 'TTM') return -1; return a.localeCompare(b)
  })
  const latestP = sortedPeriods[sortedPeriods.length - 1]
  const prevP   = sortedPeriods.length >= 2 ? sortedPeriods[sortedPeriods.length - 2] : null

  const IS = cells.filter(c => c.section === 'income-statement')
  const QOE = cells.filter(c => c.section === 'qoe')
  const BS = cells.filter(c => c.section === 'balance-sheet')

  const TTMRevenue   = getV(IS, latestP, 'total_net_revenue')
  const TTMGrossProfit = getV(IS, latestP, 'gross_profit')
  const TTMCogs      = getV(IS, latestP, 'total_cogs')
  const TTMOpex      = getV(IS, latestP, 'total_operating_expenses')
  const TTMEbitda    = getV(QOE, latestP, 'diligence_adjusted_ebitda') ?? getV(QOE, latestP, 'ebitda_as_defined') ?? getV(IS, latestP, 'operating_income_ebitda')
  const TTMNetIncome = getV(IS, latestP, 'net_income') ?? getV(QOE, latestP, 'net_income')
  const prevRevenue  = prevP ? getV(IS, prevP, 'total_net_revenue') : null
  const prevEbitda   = prevP ? (getV(QOE, prevP, 'diligence_adjusted_ebitda') ?? getV(QOE, prevP, 'ebitda_as_defined')) : null

  const grossMarginPct  = TTMRevenue && TTMGrossProfit ? (TTMGrossProfit / TTMRevenue) * 100 : null
  const ebitdaMarginPct = TTMRevenue && TTMEbitda ? (TTMEbitda / TTMRevenue) * 100 : null
  const revGrowthPct    = TTMRevenue && prevRevenue && prevRevenue !== 0 ? ((TTMRevenue - prevRevenue) / Math.abs(prevRevenue)) * 100 : null
  const totalAdjAmt     = auditEntries.reduce((s, e) => s + ((e.new_value ?? 0) - (e.old_value ?? 0)), 0)

  const doc = new PDFDocument({ size: [W, H], margins: { top: M, bottom: M, left: M, right: M }, autoFirstPage: false, bufferPages: true })
  const chunks: Buffer[] = []
  doc.on('data', (c: Buffer) => chunks.push(c))
  doc.on('error', (e: Error) => { throw e })

  let totalPages = 0

  function startPage(): number {
    doc.addPage()
    totalPages++
    return totalPages
  }

  // ── Helpers ──────────────────────────────────────────────────────────────────
  function contentHeader(title: string, subtitle?: string) {
    doc.rect(0, 0, W, 52).fill(ORANGE)
    doc.rect(0, 52, W, 3).fill(DARK)
    doc.fillColor(WHITE).font('Helvetica-Bold').fontSize(16).text(title, M, 14, { width: UW - 80 })
    if (subtitle) doc.fillColor('#F5C88A').font('Helvetica').fontSize(9).text(subtitle, M, 35, { width: UW })
    doc.fillColor(WHITE).font('Helvetica-Bold').fontSize(9).text('AIRBANK', W - M - 60, 20, { width: 60, align: 'right' })
  }

  function footer(pg: number) {
    doc.moveTo(M, H - 30).lineTo(W - M, H - 30).dash(3, { space: 3 }).strokeColor(SEP).lineWidth(0.5).stroke().undash()
    doc.fillColor(GRAY).font('Helvetica').fontSize(7.5)
    doc.text(companyName + ' — Quality of Earnings Report', M, H - 22, { width: 260 })
    doc.text('CONFIDENTIAL', W / 2 - 35, H - 22, { width: 70, align: 'center' })
    doc.text(`Page ${pg}`, W - M - 60, H - 22, { width: 60, align: 'right' })
  }

  function drawSep(y: number): number {
    doc.moveTo(M, y).lineTo(W - M, y).dash(4, { space: 3 }).strokeColor(SEP).lineWidth(0.5).stroke().undash()
    return y + 10
  }

  function drawCalloutPair(x: number, y: number, boxW: number, lbl1: string, val1: string, lbl2: string, val2: string, height = 54) {
    const bw = (boxW - 8) / 2
    // Box 1 (Reported — dark)
    doc.rect(x, y, bw, height).fillAndStroke(LG, SEP)
    doc.fillColor(TEXT).font('Helvetica-Bold').fontSize(16).text(val1, x + 4, y + 8, { width: bw - 8, align: 'center' })
    doc.fillColor(GRAY).font('Helvetica').fontSize(7).text(lbl1, x + 4, y + height - 14, { width: bw - 8, align: 'center' })
    // Box 2 (Adjusted — orange)
    doc.rect(x + bw + 8, y, bw, height).fillAndStroke(OL, ORANGE)
    doc.fillColor(DARK).font('Helvetica-Bold').fontSize(16).text(val2, x + bw + 12, y + 8, { width: bw - 8, align: 'center' })
    doc.fillColor(GRAY).font('Helvetica').fontSize(7).text(lbl2, x + bw + 12, y + height - 14, { width: bw - 8, align: 'center' })
  }

  function drawTable(y: number, headers: string[], rows: string[][], colWidths: number[], boldRows: Set<number> = new Set()): number {
    const ROW_H = 16, HDR_H = 18
    let cx = M
    doc.rect(M, y, UW, HDR_H).fill(TABLEHDR)
    headers.forEach((h, i) => {
      doc.fillColor(DARK).font('Helvetica-Bold').fontSize(8)
         .text(h, cx + 3, y + 5, { width: colWidths[i] - 6, align: i === 0 ? 'left' : 'right' })
      cx += colWidths[i]
    })
    y += HDR_H

    rows.forEach((row, ri) => {
      const isBold = boldRows.has(ri)
      if (isBold) {
        doc.rect(M, y, UW, ROW_H).fill(OL)
        doc.moveTo(M, y).lineTo(W - M, y).strokeColor(ORANGE).lineWidth(0.5).stroke()
      } else if (ri % 2 === 0) {
        doc.rect(M, y, UW, ROW_H).fill(ALTROW)
      }
      cx = M
      row.forEach((val, i) => {
        const isNeg = val.startsWith('(') || val.startsWith('-')
        doc.fillColor(isNeg && i > 0 ? RED : TEXT)
           .font(isBold ? 'Helvetica-Bold' : 'Helvetica').fontSize(8)
           .text(val ?? '—', cx + 3, y + 3, { width: colWidths[i] - 6, align: i === 0 ? 'left' : 'right' })
        cx += colWidths[i]
      })
      doc.moveTo(M, y + ROW_H).lineTo(W - M, y + ROW_H).strokeColor('#EDE0D0').lineWidth(0.25).stroke()
      y += ROW_H
    })
    return y + 4
  }

  function drawNumberedItem(x: number, y: number, itemW: number, num: number, body: string): number {
    const r = 9
    doc.circle(x + r, y + r, r).fill(ORANGE)
    doc.fillColor(WHITE).font('Helvetica-Bold').fontSize(9).text(String(num), x, y + 4, { width: r * 2, align: 'center' })
    const tx = x + r * 2 + 7
    doc.fillColor(TEXT).font('Helvetica').fontSize(8.5).text(body, tx, y + 2, { width: itemW - r * 2 - 10, lineGap: 1.5 })
    const h = doc.heightOfString(body, { width: itemW - r * 2 - 10, lineGap: 1.5 })
    return y + Math.max(h + 8, r * 2 + 4)
  }

  // ── PAGE 1: COVER ────────────────────────────────────────────────────────────
  startPage()

  // Two-tone header
  doc.rect(0, 0, W * 0.65, 110).fill(ORANGE)
  doc.rect(W * 0.65, 0, W * 0.35, 110).fill(DARK)
  doc.fillColor(WHITE).font('Helvetica-Bold').fontSize(18).text('AIRBANK', W - M - 80, 38, { width: 80, align: 'right' })
  doc.fillColor('#D4A06A').font('Helvetica').fontSize(8).text('Quality of Earnings Platform', W - M - 100, 62, { width: 100, align: 'right' })

  // Company name
  doc.fillColor(TEXT).font('Helvetica-Bold').fontSize(30).text(companyName, M, 130, { width: UW, align: 'center' })
  doc.fillColor(ORANGE).font('Helvetica-Bold').fontSize(16).text('Quality of Earnings Report', M, 174, { width: UW, align: 'center' })

  // Period chips
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })
  const periodStr = sortedPeriods.join('  |  ')

  doc.moveTo(M + 60, 210).lineTo(W - M - 60, 210).strokeColor(SEP).lineWidth(1).stroke()

  doc.fillColor(TEXT).font('Helvetica').fontSize(10)
  doc.text(`Report Date:  ${today}`, M, 224, { width: UW, align: 'center' })
  doc.text(`Reporting Periods:  ${periodStr}`, M, 244, { width: UW, align: 'center' })
  doc.text('Prepared By:  Airbank QoE Platform', M, 264, { width: UW, align: 'center' })

  // Summary financials box
  const sbY = 300
  doc.roundedRect(M, sbY, UW, 195, 5).fillAndStroke(OL, ORANGE)
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(10).text('FINANCIAL SUMMARY HIGHLIGHTS', M + 12, sbY + 10)
  doc.fillColor(GRAY).font('Helvetica').fontSize(8).text(`(${latestP ?? 'Latest Period'})`, M + 12, sbY + 26)

  const sumMetrics = [
    ['Total Revenue', fmt$(TTMRevenue, true), revGrowthPct != null ? `${revGrowthPct >= 0 ? '+' : ''}${revGrowthPct.toFixed(1)}% YoY` : ''],
    ['Gross Profit', fmt$(TTMGrossProfit, true), grossMarginPct != null ? `${grossMarginPct.toFixed(1)}% GM` : ''],
    ['Adj. EBITDA', fmt$(TTMEbitda, true), ebitdaMarginPct != null ? `${ebitdaMarginPct.toFixed(1)}% Margin` : ''],
    ['Net Income', fmt$(TTMNetIncome, true), ''],
  ]

  const bxW = (UW - 24) / 4
  sumMetrics.forEach(([lbl, val, sub], i) => {
    const bxX = M + 12 + i * (bxW + 6)
    doc.rect(bxX, sbY + 42, bxW, 80).fillAndStroke(WHITE, SEP)
    doc.fillColor(DARK).font('Helvetica-Bold').fontSize(17).text(val, bxX + 3, sbY + 53, { width: bxW - 6, align: 'center' })
    doc.fillColor(GRAY).font('Helvetica').fontSize(7).text(lbl, bxX + 3, sbY + 79, { width: bxW - 6, align: 'center' })
    if (sub) doc.fillColor(ORANGE).font('Helvetica').fontSize(6.5).text(sub, bxX + 3, sbY + 96, { width: bxW - 6, align: 'center' })
  })

  doc.fillColor(TEXT).font('Helvetica').fontSize(8.5)
  doc.text(`Analysis Sections Completed: ${[...new Set(cells.map(c => c.section))].length}`, M + 12, sbY + 136)
  doc.text(`Data Points Extracted: ${cells.length}`, M + 12 + 185, sbY + 136)
  doc.text(`Open Flags: ${flags.length}`, M + 12 + 370, sbY + 136)
  doc.text(`Analyst Overrides: ${auditEntries.length}`, M + 12, sbY + 153)
  doc.text(`Confidence Threshold Applied: 65%`, M + 12 + 185, sbY + 153)

  // Confidentiality
  doc.roundedRect(M, 520, UW, 50, 4).fillAndStroke('#FFF9F0', '#E8C8A0')
  doc.fillColor('#8B6530').font('Helvetica-Bold').fontSize(7.5).text('CONFIDENTIAL — FOR DISCUSSION PURPOSES ONLY', M + 10, 532)
  doc.fillColor(GRAY).font('Helvetica').fontSize(7.5).text(
    'This report has been prepared for internal use. Not for distribution without prior written consent. Does not constitute a formal audit.',
    M + 10, 546, { width: UW - 20 }
  )

  // Bottom bar
  doc.rect(0, H - 36, W, 36).fill(DARK)
  doc.fillColor('#A08060').font('Helvetica').fontSize(7.5).text('Prepared by Airbank Platform  |  Confidential', M, H - 22, { width: UW, align: 'center' })

  // ── PAGE 2: EXECUTIVE SUMMARY ────────────────────────────────────────────────
  startPage()
  contentHeader('Executive Summary', `${companyName} — Quality of Earnings Analysis`)
  let y = 68

  // Narrative boxes
  const narratives = [
    {
      title: 'Revenue Performance',
      text: TTMRevenue
        ? `${companyName} reported total revenue of ${fmt$(TTMRevenue)} for the period ending ${latestP}${revGrowthPct != null ? `, representing a ${revGrowthPct >= 0 ? '+' : ''}${revGrowthPct.toFixed(1)}% year-over-year change compared to ${prevP}` : ''}. ${grossMarginPct != null ? `Gross margin stands at ${grossMarginPct.toFixed(1)}%, reflecting the company's cost structure and pricing effectiveness.` : ''}`
        : 'Revenue data not yet extracted. Upload and analyze financial documents to populate this section.',
    },
    {
      title: 'Earnings Quality',
      text: TTMEbitda
        ? `Diligence-Adjusted EBITDA of ${fmt$(TTMEbitda)} (${ebitdaMarginPct != null ? ebitdaMarginPct.toFixed(1) + '% margin' : '—'}) incorporates ${auditEntries.length > 0 ? auditEntries.length + ' analyst override(s) totaling ' + fmt$sign(totalAdjAmt) : 'no analyst overrides to date'}. ${totalAdjAmt !== 0 ? `Total QoE adjustments of ${fmt$sign(totalAdjAmt)} have been applied to normalize recurring earnings power.` : 'No material adjustments were required, indicating high earnings quality.'}`
        : 'EBITDA data not yet extracted. Upload and analyze financial documents to populate this section.',
    },
    {
      title: 'Key Observations',
      text: flags.length > 0
        ? `${flags.length} open flag(s) identified: ${flags.slice(0, 3).map(f => f.title).join('; ')}${flags.length > 3 ? `; and ${flags.length - 3} additional item(s)` : ''}. Review the Flags panel for full detail.`
        : 'No open flags at this time. All extracted data points have been reviewed and no material concerns identified.',
    },
  ]

  narratives.forEach(({ title, text }) => {
    doc.rect(M, y, UW, 4).fill(ORANGE)
    doc.rect(M, y + 4, UW, 54).fill(LG)
    doc.fillColor(DARK).font('Helvetica-Bold').fontSize(9).text(title, M + 10, y + 10)
    doc.fillColor(TEXT).font('Helvetica').fontSize(8.5).text(text, M + 10, y + 24, { width: UW - 20, lineGap: 1.5 })
    y += 68
  })

  y = drawSep(y + 4)

  // Reported vs Adjusted Earnings table
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Reported vs. Adjusted Earnings', M, y)
  y += 20

  const qoeRows = [
    { k: 'net_income',              l: 'Net Income',                   sec: QOE },
    { k: 'interest_expense',        l: 'Interest Expense',             sec: QOE },
    { k: 'tax_provision',           l: 'Tax Provision',                sec: QOE },
    { k: 'depreciation',            l: 'Depreciation',                 sec: QOE },
    { k: 'amortization',            l: 'Amortization',                 sec: QOE },
    { k: 'ebitda_as_defined',       l: 'EBITDA, as Defined',           sec: QOE, bold: true },
    { k: 'total_mgmt_adjustments',  l: 'Total Management Adjustments', sec: QOE },
    { k: 'mgmt_adjusted_ebitda',    l: 'Management-Adjusted EBITDA',   sec: QOE, bold: true },
    { k: 'total_diligence_adjustments', l: 'Total Diligence Adjustments', sec: QOE },
    { k: 'diligence_adjusted_ebitda',   l: 'Diligence-Adjusted EBITDA', sec: QOE, bold: true },
  ]
  const PCOLS = sortedPeriods.length
  const colW0 = 180
  const colWP = (UW - colW0 - 60) / Math.max(PCOLS, 1)
  const colWidths = [colW0, ...sortedPeriods.map(() => Math.min(colWP, 80)), 60]

  const tblRows = qoeRows.map(({ k, l, sec }) => [
    l,
    ...sortedPeriods.map(p => fmt$(getV(sec, p, k))),
    yoy(getV(qoeRows.find(r => r.k === k)?.sec ?? QOE, latestP, k), prevP ? getV(qoeRows.find(r => r.k === k)?.sec ?? QOE, prevP, k) : null),
  ])
  const boldIdxs = new Set(qoeRows.map((r, i) => r.bold ? i : -1).filter(i => i >= 0))
  y = drawTable(y, ['Line Item', ...sortedPeriods, 'YoY'], tblRows, colWidths, boldIdxs)
  y += 4

  // Key Adjustments
  if (auditEntries.length > 0) {
    y = drawSep(y)
    doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Key Adjustments', M, y)
    y += 20

    // Group by row_key
    const byKey: Record<string, { desc: string; delta: number; period: string }> = {}
    for (const e of auditEntries) {
      if (!e.cell) continue
      const k = e.cell.row_key
      const delta = (e.new_value ?? 0) - (e.old_value ?? 0)
      if (!byKey[k]) byKey[k] = { desc: e.note ?? humanize(e.cell.row_key), delta: 0, period: e.cell.period }
      byKey[k].delta += delta
      if (e.note) byKey[k].desc = e.note
    }

    const adjList = Object.values(byKey).filter(a => Math.abs(a.delta) >= 1).slice(0, 6)
    const halfW = (UW - 16) / 2
    let numX = M, numY = y, col = 0

    for (let i = 0; i < adjList.length; i++) {
      const adj = adjList[i]
      const text = `${adj.desc}  (${adj.period}: ${fmt$sign(adj.delta)})`
      const startY = numY
      numY = drawNumberedItem(numX, numY, halfW, i + 1, text)
      if (col === 0) { col = 1; numX = M + halfW + 16; numY = startY }
      else { col = 0; numX = M }
    }
    y = Math.max(numY, y) + 8
  }

  footer(2)

  // ── PAGE 3: REVENUE & COGS ANALYSIS ─────────────────────────────────────────
  startPage()
  contentHeader('Data Analysis — Revenue & Cost of Goods Sold', `${latestP} Analysis  |  ${companyName}`)
  y = 68

  // Revenue Analysis
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Revenue Analysis', M, y)
  y += 20

  const prevGP = prevP ? getV(IS, prevP, 'gross_profit') : null
  drawCalloutPair(M, y, UW / 2 - 4, 'Reported Revenue', fmt$(TTMRevenue, true), 'Prior Period', fmt$(prevRevenue, true))
  drawCalloutPair(M + UW / 2 + 4, y, UW / 2 - 4, 'Gross Profit', fmt$(TTMGrossProfit, true), 'Gross Margin %', grossMarginPct != null ? `${grossMarginPct.toFixed(1)}%` : '—')
  y += 68

  doc.fillColor(TEXT).font('Helvetica').fontSize(9).text(
    TTMRevenue
      ? `Total net revenue of ${fmt$(TTMRevenue)} represents${revGrowthPct != null ? ` a ${Math.abs(revGrowthPct).toFixed(1)}% ${revGrowthPct >= 0 ? 'increase' : 'decrease'} year-over-year` : ' the company\'s top-line results'} for ${latestP}. ${grossMarginPct != null ? `The ${grossMarginPct.toFixed(1)}% gross margin reflects the relationship between net revenues and direct costs of sales.` : ''}`
      : 'Revenue data will appear here after uploading and analyzing financial documents.',
    M, y, { width: UW, lineGap: 2 }
  )
  y += 32

  // Revenue table
  const revTableRows = [
    ['Gross Product Sales', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'gross_product_sales'))), pctOfRev(getV(IS, latestP, 'gross_product_sales'), TTMRevenue)],
    ['Returns & Allowances', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'returns_allowances'))), pctOfRev(getV(IS, latestP, 'returns_allowances'), TTMRevenue)],
    ['Promotional Discounts', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'promotional_discounts'))), pctOfRev(getV(IS, latestP, 'promotional_discounts'), TTMRevenue)],
    ['Net Product Sales', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'net_product_sales'))), pctOfRev(getV(IS, latestP, 'net_product_sales'), TTMRevenue)],
    ['Shipping Revenue', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'shipping_revenue'))), pctOfRev(getV(IS, latestP, 'shipping_revenue'), TTMRevenue)],
    ['Total Net Revenue', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'total_net_revenue'))), '100.0%'],
  ]
  const pColW = Math.min((UW - 180 - 55) / Math.max(PCOLS, 1), 75)
  const revColW = [180, ...sortedPeriods.map(() => pColW), 55]
  y = drawTable(y, ['Line Item', ...sortedPeriods, '% Rev'], revTableRows, revColW, new Set([3, 5]))
  y += 6

  y = drawSep(y)

  // COGS Analysis
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Cost of Goods Sold (COGS) & Gross Margin', M, y)
  y += 20

  const prevRev  = prevP ? getV(IS, prevP, 'total_net_revenue') : null
  const prevCogs = prevP ? getV(IS, prevP, 'total_cogs') : null
  const prevGM = prevP && prevRev ? (getV(IS, prevP, 'gross_profit') ?? 0) / prevRev * 100 : null

  drawCalloutPair(M, y, UW / 2 - 4, 'Reported COGS', fmt$(TTMCogs, true), 'Prior Period COGS', fmt$(prevCogs, true))
  drawCalloutPair(M + UW / 2 + 4, y, UW / 2 - 4, 'Gross Margin (TTM)', grossMarginPct != null ? `${grossMarginPct.toFixed(1)}%` : '—', 'Prior Period GM', prevGM != null ? `${prevGM.toFixed(1)}%` : '—')
  y += 68

  doc.fillColor(TEXT).font('Helvetica').fontSize(9).text(
    TTMCogs
      ? `Cost of goods sold of ${fmt$(TTMCogs)} represents ${pctOfRev(TTMCogs, TTMRevenue)} of total net revenue. ${TTMGrossProfit ? `Gross profit of ${fmt$(TTMGrossProfit)} yields a ${fmtPct(grossMarginPct)} gross margin, ` : ''}${prevGM != null && grossMarginPct != null ? `compared to ${prevGM.toFixed(1)}% in ${prevP}, a change of ${(grossMarginPct - prevGM) >= 0 ? '+' : ''}${(grossMarginPct - prevGM).toFixed(1)} percentage points.` : 'which is used as the basis for all profitability analysis.'}`
      : 'COGS data will appear here after uploading and analyzing financial documents.',
    M, y, { width: UW, lineGap: 2 }
  )
  y += 32

  const cogsTableRows = [
    ['Product Cost of Goods Sold', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'product_cogs'))), pctOfRev(getV(IS, latestP, 'product_cogs'), TTMRevenue)],
    ['Warehousing & Fulfillment', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'warehousing_fulfillment'))), pctOfRev(getV(IS, latestP, 'warehousing_fulfillment'), TTMRevenue)],
    ['Shipping & Freight Out', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'shipping_freight_out'))), pctOfRev(getV(IS, latestP, 'shipping_freight_out'), TTMRevenue)],
    ['Inventory Adjustments', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'inventory_adjustments'))), pctOfRev(getV(IS, latestP, 'inventory_adjustments'), TTMRevenue)],
    ['Total COGS', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'total_cogs'))), pctOfRev(TTMCogs, TTMRevenue)],
    ['Gross Profit', ...sortedPeriods.map(p => fmt$(getV(IS, p, 'gross_profit'))), fmtPct(grossMarginPct)],
  ]
  y = drawTable(y, ['Line Item', ...sortedPeriods, '% Rev'], cogsTableRows, revColW, new Set([4, 5]))

  footer(3)

  // ── PAGE 4: EBITDA & OPEX ────────────────────────────────────────────────────
  startPage()
  contentHeader('Data Analysis — EBITDA Bridge & Operating Expenses', `${latestP} Analysis  |  ${companyName}`)
  y = 68

  // EBITDA Analysis
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('EBITDA Analysis', M, y)
  y += 20

  const TTMEbitdaDefined = getV(QOE, latestP, 'ebitda_as_defined')
  const TTMMgmtAdj = getV(QOE, latestP, 'total_mgmt_adjustments')
  const TTMDilAdj = getV(QOE, latestP, 'total_diligence_adjustments')

  drawCalloutPair(M, y, UW / 3 - 4, 'EBITDA as Defined', fmt$(TTMEbitdaDefined, true), 'Adj. EBITDA', fmt$(TTMEbitda, true))
  drawCalloutPair(M + UW / 3 + 4, y, UW / 3 - 4, 'Total Mgmt Adj.', fmt$sign(TTMMgmtAdj), 'Total Diligence Adj.', fmt$sign(TTMDilAdj))
  drawCalloutPair(M + (UW / 3 + 4) * 2, y, UW / 3 - 4, 'EBITDA Margin', fmtPct(ebitdaMarginPct), 'YoY EBITDA', yoy(TTMEbitda, prevEbitda))
  y += 68

  doc.fillColor(TEXT).font('Helvetica').fontSize(9).text(
    TTMEbitda
      ? `Diligence-Adjusted EBITDA of ${fmt$(TTMEbitda)} (${fmtPct(ebitdaMarginPct)} margin) reflects ${auditEntries.length > 0 ? `${auditEntries.length} quality of earnings adjustments totaling ${fmt$sign(totalAdjAmt)}. These adjustments normalize the earnings base to reflect the company's sustainable operating performance.` : 'no material quality of earnings adjustments, indicating that reported earnings are a reliable proxy for sustainable earnings power.'}`
      : 'EBITDA data will appear here after uploading and analyzing financial documents.',
    M, y, { width: UW, lineGap: 2 }
  )
  y += 32

  const bridgeTableRows = [
    ['Net Income', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'net_income'))), ''],
    ['+ Interest Expense', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'interest_expense'))), ''],
    ['+ Tax Provision', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'tax_provision'))), ''],
    ['+ Depreciation & Amortization', ...sortedPeriods.map(p => {
      const d = getV(QOE, p, 'depreciation') ?? 0
      const a = getV(QOE, p, 'amortization') ?? 0
      return fmt$(d + a !== 0 ? d + a : null)
    }), ''],
    ['= EBITDA, as Defined', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'ebitda_as_defined'))), yoy(TTMEbitdaDefined, prevP ? getV(QOE, prevP, 'ebitda_as_defined') : null)],
    ['Total Management Adjustments', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'total_mgmt_adjustments'))), ''],
    ['= Management-Adjusted EBITDA', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'mgmt_adjusted_ebitda'))), ''],
    ['Total Diligence Adjustments', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'total_diligence_adjustments'))), ''],
    ['= Diligence-Adjusted EBITDA', ...sortedPeriods.map(p => fmt$(getV(QOE, p, 'diligence_adjusted_ebitda'))), yoy(TTMEbitda, prevEbitda)],
  ]
  const bridgeColW = [190, ...sortedPeriods.map(() => Math.min(pColW, 76)), 60]
  y = drawTable(y, ['Line Item', ...sortedPeriods, 'YoY'], bridgeTableRows, bridgeColW, new Set([4, 6, 8]))
  y += 6

  y = drawSep(y)

  // OpEx Analysis
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Operating Expenses', M, y)
  y += 20

  const prevOpex = prevP ? getV(IS, prevP, 'total_operating_expenses') : null
  drawCalloutPair(M, y, UW / 2 - 4, 'Total OpEx', fmt$(TTMOpex, true), 'Prior Period OpEx', fmt$(prevOpex, true))
  drawCalloutPair(M + UW / 2 + 4, y, UW / 2 - 4, 'OpEx % of Revenue', pctOfRev(TTMOpex, TTMRevenue), 'YoY OpEx Change', yoy(TTMOpex, prevOpex))
  y += 68

  doc.fillColor(TEXT).font('Helvetica').fontSize(9).text(
    TTMOpex
      ? `Total operating expenses of ${fmt$(TTMOpex)} represent ${pctOfRev(TTMOpex, TTMRevenue)} of total net revenue${prevOpex != null ? `, compared to ${fmt$(prevOpex)} (${pctOfRev(prevOpex, prevRevenue)}) in ${prevP}` : ''}. The largest expense categories are analyzed below.`
      : 'Operating expense data will appear here after uploading and analyzing financial documents.',
    M, y, { width: UW, lineGap: 2 }
  )
  y += 32

  const opexKeys = [
    ['sales_marketing', 'Sales & Marketing'],
    ['salaries_wages_payroll', 'Salaries, Wages & Payroll'],
    ['employee_benefits_401k', 'Employee Benefits & 401k'],
    ['rent_occupancy', 'Rent & Occupancy'],
    ['professional_services', 'Professional Services'],
    ['technology_software', 'Technology & Software'],
    ['total_operating_expenses', 'Total Operating Expenses'],
  ]
  const opexRows = opexKeys.map(([k, l]) => [l, ...sortedPeriods.map(p => fmt$(getV(IS, p, k))), pctOfRev(getV(IS, latestP, k), TTMRevenue)])
  y = drawTable(y, ['Line Item', ...sortedPeriods, '% Rev'], opexRows, revColW, new Set([opexKeys.length - 1]))

  footer(4)

  // ── PAGE 5: CONCERNS & CONCLUSIONS ──────────────────────────────────────────
  startPage()
  contentHeader('Key Areas of Concern & Conclusions', `${companyName} — Summary Findings`)
  y = 68

  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Key Areas of Concern', M, y)
  y += 16

  const openFlags = flags.slice(0, 8)
  if (openFlags.length === 0) {
    doc.rect(M, y, UW, 44).fillAndStroke(LG, SEP)
    doc.fillColor(GRAY).font('Helvetica-Oblique').fontSize(9).text('No open flags — all items reviewed and no material concerns identified.', M + 12, y + 14, { width: UW - 24 })
    y += 56
  } else {
    openFlags.forEach((flag, i) => {
      const BADGE_W = 55, BADGE_H = 14
      const badgeColor = flag.severity === 'critical' ? RED : flag.severity === 'warning' ? '#E65100' : GRAY
      const rowH = 44
      if (i % 2 === 0) doc.rect(M, y, UW, rowH).fill(ALTROW)
      doc.rect(M + 6, y + 5, BADGE_W, BADGE_H).fill(badgeColor)
      doc.fillColor(WHITE).font('Helvetica-Bold').fontSize(6.5).text(flag.severity.toUpperCase(), M + 6, y + 8, { width: BADGE_W, align: 'center' })
      doc.fillColor(TEXT).font('Helvetica-Bold').fontSize(8.5).text(flag.title, M + BADGE_W + 14, y + 4, { width: UW - BADGE_W - 20 })
      const bodyText = (flag.body ?? '').substring(0, 140) + ((flag.body?.length ?? 0) > 140 ? '...' : '')
      doc.fillColor(GRAY).font('Helvetica').fontSize(8).text(bodyText, M + BADGE_W + 14, y + 17, { width: UW - BADGE_W - 20 })
      doc.fillColor(GRAY).font('Helvetica-Oblique').fontSize(7).text(`${humanize(flag.section)} › ${humanize(flag.row_key)}${flag.period ? ` › ${flag.period}` : ''}`, M + BADGE_W + 14, y + 32, { width: UW - BADGE_W - 20 })
      doc.moveTo(M, y + rowH).lineTo(W - M, y + rowH).strokeColor('#EDE0D0').lineWidth(0.25).stroke()
      y += rowH
      if (y > H - 200) { y = H - 200 } // safety guard
    })
    if (flags.length > 8) {
      doc.fillColor(GRAY).font('Helvetica-Oblique').fontSize(8).text(`... and ${flags.length - 8} additional flag(s). See the Flags panel in the Airbank workbook for full detail.`, M + 6, y + 4)
      y += 20
    }
  }

  y = drawSep(y + 10)

  // Conclusions
  doc.fillColor(DARK).font('Helvetica-Bold').fontSize(13).text('Conclusions & Recommendations', M, y)
  y += 18

  const criticalFlags = flags.filter(f => f.severity === 'critical').length
  const conclusions = [
    `Revenue: ${TTMRevenue ? `${companyName} generated total net revenue of ${fmt$(TTMRevenue)} in ${latestP}${revGrowthPct != null ? `, representing a ${Math.abs(revGrowthPct).toFixed(1)}% ${revGrowthPct >= 0 ? 'increase' : 'decrease'} year-over-year` : ''}. ${grossMarginPct != null ? `Gross margin of ${grossMarginPct.toFixed(1)}% indicates ${grossMarginPct > 40 ? 'strong' : grossMarginPct > 25 ? 'acceptable' : 'tight'} cost controls.` : ''}` : 'Revenue data pending analysis.'}`,
    `Earnings Quality: ${TTMEbitda ? `Diligence-Adjusted EBITDA of ${fmt$(TTMEbitda)} (${fmtPct(ebitdaMarginPct)}) reflects ${auditEntries.length} QoE adjustment(s) totaling ${fmt$sign(totalAdjAmt)}. ${Math.abs(totalAdjAmt) > 0 ? 'These adjustments normalize one-time and non-recurring items from the earnings base.' : 'No material adjustments were required.'}` : 'EBITDA analysis pending.'}`,
    `Open Items: ${criticalFlags > 0 ? `${criticalFlags} critical flag(s) require resolution prior to any transaction close. ` : 'No critical flags identified. '}${flags.length > 0 ? `A total of ${flags.length} flag(s) are open and require review.` : 'All flags have been resolved.'}`,
    `Recommendation: ${auditEntries.length > 0 ? `The ${auditEntries.length} analyst override(s) applied represent management\'s best estimate of sustainable earnings. Independent verification of supporting documentation is recommended for all diligence adjustments.` : 'Continue the analysis process by uploading all relevant financial documents and completing the AI-assisted extraction.'}`,
  ]

  conclusions.forEach((text, i) => {
    const parts = text.split(': ')
    const label = parts[0]
    const body  = parts.slice(1).join(': ')
    if (y + 40 > H - 55) return
    doc.rect(M, y, 3, 28).fill(ORANGE)
    doc.fillColor(DARK).font('Helvetica-Bold').fontSize(9).text(label + ':', M + 10, y + 2)
    doc.fillColor(TEXT).font('Helvetica').fontSize(8.5).text(body, M + 10, y + 14, { width: UW - 14, lineGap: 1.5 })
    y += 38
  })

  // Final disclaimer
  y += 8
  doc.rect(M, y, UW, 32).fill(LG)
  doc.fillColor(GRAY).font('Helvetica-Oblique').fontSize(7.5).text(
    'This report has been prepared by Airbank QoE Platform for informational purposes only. It does not constitute a formal audit, review, or compilation engagement. All figures are based on management-provided data and have not been independently verified.',
    M + 8, y + 8, { width: UW - 16, lineGap: 1.5 }
  )

  footer(5)

  // ── Apply footers with correct total page count ─────────────────────────────
  doc.end()

  return new Promise((resolve, reject) => {
    doc.on('end', () => resolve(Buffer.concat(chunks)))
    doc.on('error', reject)
  })
}

// ─── Google Sheets ─────────────────────────────────────────────────────────────
async function exportToGoogleSheets(
  companyName: string,
  cells: WorkbookCell[],
  periods: string[],
): Promise<string> {
  const { google } = await import('googleapis')
  const credentialsJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON
  if (!credentialsJson) throw new Error('Google credentials not configured')

  const credentials = JSON.parse(credentialsJson)
  const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] })
  const sheets = google.sheets({ version: 'v4', auth })

  const spreadsheet = await sheets.spreadsheets.create({
    requestBody: { properties: { title: `QoE Workbook - ${companyName}` } },
  })
  const spreadsheetId = spreadsheet.data.spreadsheetId!
  const sectionList = [...new Set(cells.map(c => c.section))]
  const sheetRequests = sectionList.slice(1).map((s, i) => ({
    addSheet: { properties: { title: s.replace(/-/g, ' ').slice(0, 31), index: i + 1 } },
  }))
  if (sheetRequests.length > 0) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: sheetRequests } })
  }

  const valueRanges = sectionList.map(section => {
    const sc = cells.filter(c => c.section === section)
    const rowKeys = [...new Set(sc.map(c => c.row_key))]
    const sheetTitle = section.replace(/-/g, ' ').slice(0, 31)
    const values = [
      ['Line Item', ...periods],
      ...rowKeys.map(rk => [humanize(rk), ...periods.map(p => sc.find(c => c.row_key === rk && c.period === p)?.raw_value ?? '')]),
    ]
    return { range: `${sheetTitle}!A1`, values }
  })

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'RAW', data: valueRanges },
  })

  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}`
}

// ─── Route Handler ─────────────────────────────────────────────────────────────
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json()
  const format: 'excel' | 'pdf' | 'sheets' | 'airtable' = body.format ?? 'excel'

  const serviceClient = createSupabaseServiceClient()

  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, company_name, periods, created_by')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Workbook not found' }, { status: 404 })

  const [{ data: cells }, { data: auditEntries }, { data: flagsData }] = await Promise.all([
    serviceClient
      .from('workbook_cells')
      .select('section, row_key, period, raw_value, display_value, is_overridden, confidence, source_excerpt')
      .eq('workbook_id', workbookId)
      .order('section').order('row_key'),
    serviceClient
      .from('audit_entries')
      .select('edited_at, old_value, new_value, note, cell:workbook_cells(section, row_key, period)')
      .eq('workbook_id', workbookId)
      .order('edited_at', { ascending: false }),
    serviceClient
      .from('cell_flags')
      .select('section, row_key, period, flag_type, severity, title, body, resolved_at, created_by_ai')
      .eq('workbook_id', workbookId)
      .is('resolved_at', null)
      .order('created_at', { ascending: false }),
  ])

  const periods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']
  const allCells = (cells ?? []) as WorkbookCell[]
  const allAudit = ((auditEntries ?? []) as unknown) as AuditEntry[]
  const allFlags = ((flagsData ?? []) as unknown) as WorkbookFlag[]

  const safe = workbook.company_name.replace(/[^a-z0-9]/gi, '_')

  if (format === 'excel') {
    try {
      const buf = await buildExcel(workbook.company_name, allCells, allAudit, periods, allFlags)
      return new NextResponse(buf as unknown as BodyInit, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Content-Disposition': `attachment; filename="${safe}_QoE.xlsx"`,
        },
      })
    } catch (err) {
      return NextResponse.json({ error: String(err) }, { status: 500 })
    }
  }

  if (format === 'pdf') {
    try {
      const buf = await buildPdf(workbook.company_name, allCells, allAudit, allFlags, periods)
      return new NextResponse(buf as unknown as BodyInit, {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${safe}_QoE.pdf"`,
        },
      })
    } catch (err) {
      return NextResponse.json({ error: String(err) }, { status: 500 })
    }
  }

  if (format === 'sheets') {
    try {
      const url = await exportToGoogleSheets(workbook.company_name, allCells, periods)
      return NextResponse.json({ url })
    } catch (err) {
      return NextResponse.json({ error: String(err) }, { status: 500 })
    }
  }

  return NextResponse.json({ error: `Format "${format}" not yet implemented` }, { status: 400 })
}
