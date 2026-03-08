'use client'

import { useState, use, useEffect, useCallback } from 'react'
import { useRouter } from 'next/navigation'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { LineChart, Line, BarChart, Bar, ComposedChart, PieChart as RechartsPie, Pie, Cell, AreaChart, Area, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LabelList } from 'recharts'
import { ChartContainer, ChartTooltip, ChartTooltipContent, ChartLegend, ChartLegendContent } from '@/components/ui/chart'
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Download, ChevronDown, FileSpreadsheet, Sheet, Database, FileBarChart, TrendingUp, FileText, Scale, DollarSign, ShoppingCart, Calendar, Banknote, Package, ClipboardCheck, BarChart3, PieChart, Settings, Loader2, Shield, BookOpen, Sparkles } from 'lucide-react'
import Image from 'next/image'
import { AuditableCell, type SourceRef, type CellFlag } from '@/components/auditable-cell'
import { WorkbookSettingsDialog } from '@/components/workbook-settings-dialog'
import { DocumentViewerPanel } from '@/components/document-viewer-panel'
import { RiskDiligenceSection } from '@/components/risk-diligence-section'
import { CompleteQoeSection } from '@/components/complete-qoe-section'
import { useLayoutContext } from '@/lib/layout-context'

const formatCurrency = (value: number) => {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 }).format(value)
}

const formatPercent = (value: number) => {
  return `${value.toFixed(1)}%`
}

export default function WorkbookPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params)

  const DEMO_NAMES: { [key: string]: string } = {
    '1': 'Acme Corp',
    '2': 'TechStart Inc',
    '3': 'Global Industries',
    '4': 'Finance Group',
    '5': 'Retail Solutions',
  }

  const isDemoWorkbook = /^\d$/.test(id)
  const router = useRouter()

  const [activeSection, setActiveSection] = useState('overview')
  const [settingsOpen, setSettingsOpen] = useState(false)

  // ── Live data layer ─────────────────────────────────────────────
  interface LiveCell {
    id: string
    section: string
    row_key: string
    period: string
    raw_value: number | null
    display_value: string | null
    is_calculated: boolean
    is_overridden: boolean
    source_page: number | null
    source_excerpt: string | null
    confidence: number | null
    source_document: { id: string; file_name: string } | null
    flags?: CellFlag[]
  }

  const [liveCells, setLiveCells] = useState<LiveCell[]>([])
  const [cellsLoading, setCellsLoading] = useState(false)
  const [viewerOpen, setViewerOpen] = useState(false)
  const [viewerSource, setViewerSource] = useState<SourceRef | null>(null)
  const [exporting, setExporting] = useState(false)
  const [exportError, setExportError] = useState<string | null>(null)
  const [liveWorkbookName, setLiveWorkbookName] = useState<string | null>(null)
  const [workbookStatus, setWorkbookStatus] = useState<string | null>(null)
  const [flags, setFlags] = useState<{ id: string; resolved_at: string | null }[]>([])
  const { chatOpen, openChat, closeChat, setCellRef, flagsRefreshRef } = useLayoutContext()

  const workbookName = liveWorkbookName ?? DEMO_NAMES[id] ?? 'Workbook'

  useEffect(() => {
    const name = DEMO_NAMES[id]
    if (name) document.title = `Airbank - ${name}`
    else document.title = 'Airbank - Workbook'
  }, [id])

  const fetchCells = useCallback(async () => {
    if (isDemoWorkbook) return
    setCellsLoading(true)
    try {
      const [cellsRes, wbRes, flagsRes] = await Promise.all([
        fetch(`/api/workbooks/${id}/cells`),
        fetch(`/api/workbooks/${id}`),
        fetch(`/api/workbooks/${id}/flags`),
      ])
      if (cellsRes.ok) {
        const data: LiveCell[] = await cellsRes.json()
        setLiveCells(data)
      }
      if (wbRes.ok) {
        const wb = await wbRes.json()
        if (wb.company_name) { setLiveWorkbookName(wb.company_name); document.title = `Airbank - ${wb.company_name}` }
        if (wb.status) setWorkbookStatus(wb.status)
      }
      if (flagsRes.ok) {
        const flagData = await flagsRes.json()
        setFlags(Array.isArray(flagData) ? flagData : [])
      }
    } catch {
      // API not configured — use hardcoded data
    } finally {
      setCellsLoading(false)
    }
  }, [id])

  useEffect(() => {
    fetchCells()
  }, [fetchCells])

  // Register fetchCells so the AI panel (in layout) can refresh cells after flag changes
  useEffect(() => {
    flagsRefreshRef.current = fetchCells
    return () => { flagsRefreshRef.current = undefined }
  }, [fetchCells, flagsRefreshRef])

  /** Get a numeric value from live cells, falling back to the hardcoded value. */
  const getLiveValue = useCallback(
    (section: string, rowKey: string, period: string): number | null => {
      const cell = liveCells.find(
        (c) => c.section === section && c.row_key === rowKey && c.period === period
      )
      return cell?.raw_value ?? null
    },
    [liveCells]
  )

  /** Build a SourceRef for a cell (if live data is available). */
  const getCellSourceRef = useCallback(
    (section: string, rowKey: string, period: string): SourceRef | null => {
      const cell = liveCells.find(
        (c) => c.section === section && c.row_key === rowKey && c.period === period
      )
      if (!cell || !cell.source_document) return null
      return {
        documentId: cell.source_document.id,
        documentName: cell.source_document.file_name,
        page: cell.source_page,
        excerpt: cell.source_excerpt ?? '',
        confidence: cell.confidence ?? 0,
      }
    },
    [liveCells]
  )

  /** Get the cell id for persistence. */
  const getCellId = useCallback(
    (section: string, rowKey: string, period: string): string | undefined => {
      return liveCells.find(
        (c) => c.section === section && c.row_key === rowKey && c.period === period
      )?.id
    },
    [liveCells]
  )

  /** Get flags for a specific cell. */
  const getCellFlags = useCallback(
    (section: string, rowKey: string, period: string): CellFlag[] => {
      const cell = liveCells.find(
        (c) => c.section === section && c.row_key === rowKey && c.period === period
      )
      return cell?.flags ?? []
    },
    [liveCells]
  )

  const openFlagCount = flags.filter(f => !f.resolved_at).length

  /** Create a flag from an AuditableCell and refresh. */
  const handleFlagCreate = useCallback(
    async (section: string, rowKey: string, period: string, flag: { title: string; body: string; flag_type: string }) => {
      if (!id || isDemoWorkbook) return
      const cellId = getCellId(section, rowKey, period)
      await fetch(`/api/workbooks/${id}/flags`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          section,
          row_key: rowKey,
          period,
          flag_type: flag.flag_type,
          severity: 'warning',
          title: flag.title,
          body: flag.body,
          cell_id: cellId ?? null,
        }),
      })
      await fetchCells()
    },
    [id, isDemoWorkbook, getCellId, fetchCells]
  )

  const handleViewSource = useCallback((sourceRef: SourceRef) => {
    setViewerSource(sourceRef)
    setViewerOpen(true)
  }, [])

  const handleExport = useCallback(
    async (format: 'excel' | 'pdf' | 'sheets' | 'airtable') => {
      if (isDemoWorkbook) return
      setExporting(true)
      setExportError(null)
      try {
        const res = await fetch(`/api/workbooks/${id}/export`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ format }),
        })
        if (!res.ok) {
          const ct = res.headers.get('content-type') ?? ''
          const body = ct.includes('application/json')
            ? ((await res.json()) as { error?: string }).error ?? 'Export failed'
            : await res.text()
          throw new Error(body)
        }
        if (format === 'excel' || format === 'pdf') {
          const blob = await res.blob()
          const url = URL.createObjectURL(blob)
          const a = document.createElement('a')
          a.href = url
          const ext = format === 'pdf' ? 'pdf' : 'xlsx'
          a.download = `${workbookName.replace(/[^a-z0-9]/gi, '_')}_QoE.${ext}`
          a.click()
          URL.revokeObjectURL(url)
        } else if (format === 'sheets') {
          const data = await res.json()
          window.open(data.url, '_blank')
        }
      } catch (err) {
        setExportError(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    },
    [id, isDemoWorkbook, workbookName]
  )

  /** Merge helper: returns formatted live value or falls back to hardcoded. */
  const getDisplayValue = useCallback(
    (section: string, rowKey: string, period: string, fallback: number, isPercent = false): string => {
      const live = getLiveValue(section, rowKey, period)
      const val = live !== null ? live : fallback
      return isPercent ? formatPercent(val) : formatCurrency(val)
    },
    [getLiveValue]
  )
  // ── End live data layer ─────────────────────────────────────────

  // Sidebar sections
  const sections = [
    { id: 'complete-qoe', name: 'Complete QoE', icon: BookOpen },
    { id: 'overview', name: 'Overview', icon: FileBarChart },
    { id: 'qoe', name: 'Quality of Earnings', icon: TrendingUp },
    { id: 'income-statement', name: 'Income Statement', icon: FileText },
    { id: 'balance-sheet', name: 'Balance Sheet', icon: Scale },
    { id: 'sales-channel', name: 'Sales by Channel', icon: ShoppingCart },
    { id: 'margins-month', name: 'Margins by Month', icon: Calendar },
    { id: 'proof-cash', name: 'Proof of Cash', icon: Banknote },
    { id: 'working-capital', name: 'Working Capital', icon: DollarSign },
    { id: 'cogs-vendors', name: 'COGS Vendors', icon: Package },
    { id: 'testing', name: 'AP/Accrual Testing', icon: ClipboardCheck },
    { id: 'charts', name: 'Charts & Analytics', icon: BarChart3 },
    { id: 'risk-diligence', name: 'Risk & Diligence', icon: Shield },
  ]

  // Overview Data (FY20-FY22 + TTM)
  const overviewData = [
    { metric: 'Reported Revenues', fy20: 42187456, fy21: 51482903, fy22: 62745381, ttm: 68293742 },
    { metric: 'EBITDA, as defined', fy20: 4630517, fy21: 6177949, fy22: 7217823, ttm: 7960816 },
    { metric: 'Diligence-Adjusted EBITDA', fy20: 4875291, fy21: 6524835, fy22: 7682947, ttm: 8547239 },
    { metric: 'Adjusted EBITDA Margin', fy20: 11.6, fy21: 12.7, fy22: 12.2, ttm: 12.5, isPercent: true },
  ]

  // Quality of Earnings Data
  const qoeData = [
    { item: 'Net Income', fy20: 4558429, fy21: 6091874, fy22: 7120305, ttm: 7858378, source: 'Audited Financial Statements - Year End Dec 31' },
    { item: 'Interest Expense', fy20: 42850, fy21: 52387, fy22: 61245, ttm: 65182, source: 'General Ledger Acct #7100 - Interest Expense' },
    { item: 'Tax Provision', fy20: 23185, fy21: 27038, fy22: 29583, ttm: 31256, source: 'Form 1120 - Federal & State Income Tax Returns' },
    { item: 'Depreciation', fy20: 4853, fy21: 5250, fy22: 5390, ttm: 4700, source: 'Fixed Asset Schedule - Acc Depreciation Roll-Forward' },
    { item: 'Amortization', fy20: 1200, fy21: 1400, fy22: 1300, ttm: 1300, source: 'Intangible Asset Amortization Schedule - Software & Patents' },
    { item: 'EBITDA, as defined', fy20: 4630517, fy21: 6177949, fy22: 7217823, ttm: 7960816, source: 'Calculated per Management Methodology', isBold: true },
    { item: 'a) Charitable Contributions', fy20: 15000, fy21: 248500, fy22: 287350, ttm: 315920, source: 'GL Acct #6850 - Discretionary charitable donations to local foundation', isAdjustment: true },
    { item: 'b) Owner Tax & Legal', fy20: 18500, fy21: 18500, fy22: 21200, ttm: 22450, source: 'GL Acct #6420 - Personal CPA & estate planning fees', isAdjustment: true },
    { item: 'c) Excess Owner Compensation', fy20: 185200, fy21: 208750, fy22: 236280, ttm: 248190, source: 'Payroll Register - Excess above market CEO salary of $175K', isAdjustment: true },
    { item: 'd) Personal Expenses', fy20: 26574, fy21: 31850, fy22: 35894, ttm: 39863, source: 'GL Acct #6900 - Country club, personal travel, family expenses', isAdjustment: true },
    { item: 'Total Management Adjustments', fy20: 245274, fy21: 507600, fy22: 580724, ttm: 626423, source: 'Sum of Management Add-backs', isBold: true },
    { item: 'Management-Adjusted EBITDA', fy20: 4875791, fy21: 6685549, fy22: 7798547, ttm: 8587239, source: 'EBITDA as Defined + Management Adjustments', isBold: true },
    { item: 'e) Professional Fees - One-time', fy20: 0, fy21: 52400, fy22: 58900, ttm: 68200, source: 'GL Acct #6425 - Litigation settlement & SEC filing costs related to terminated acquisition', isAdjustment: true },
    { item: 'f) Executive Search Fees', fy20: 0, fy21: 65000, fy22: 48000, ttm: 52000, source: 'GL Acct #6340 - Non-recurring CFO & VP Sales recruiting expenses', isAdjustment: true },
    { item: 'g) Facility Relocation', fy20: 0, fy21: 0, fy22: 38750, ttm: 41500, source: 'GL Acct #6520 - One-time warehouse relocation & build-out costs', isAdjustment: true },
    { item: 'h) Severance', fy20: 0, fy21: 0, fy22: 85400, ttm: 92300, source: 'GL Acct #6120 - Restructuring severance for 3 terminated employees', isAdjustment: true },
    { item: 'i) Above-Market Rent', fy20: 0, fy21: -158714, fy22: -162500, ttm: -168900, source: 'GL Acct #6500 - Normalize to market rate of $18/sqft vs $28/sqft paid (related party lease)', isAdjustment: true },
    { item: 'j) Normalize Owner Comp', fy20: -244500, fy21: -382300, fy22: -404000, ttm: -425900, source: 'Payroll Register - Reduce total owner comp to market CEO salary of $175K + benefits', isAdjustment: true },
    { item: 'Total Diligence Adjustments', fy20: -244500, fy21: -423614, fy22: -335450, ttm: -340800, source: 'Sum of Diligence Add-backs (Net)', isBold: true },
    { item: 'Diligence-Adjusted EBITDA', fy20: 4875291, fy21: 6524835, fy22: 7682947, ttm: 8547239, source: 'Final Normalized Run-Rate EBITDA - Buyer Perspective', isBold: true },
  ]

  // Income Statement Data
  const incomeStatementData = [
    { item: 'Gross Product Sales', fy20: 47832185, fy21: 58926741, fy22: 71584293, ttm: 77842953, pct: 114, category: 'Revenue' },
    { item: 'Returns & Allowances', fy20: -2847293, fy21: -3428574, fy22: -4183928, ttm: -4562485, pct: -7, category: 'Revenue' },
    { item: 'Promotional Discounts', fy20: -3218547, fy21: -3956382, fy22: -4813574, ttm: -5246193, pct: -8, category: 'Revenue' },
    { item: 'Net Product Sales', fy20: 41766345, fy21: 51541785, fy22: 62586791, ttm: 68034275, pct: 100, category: 'Revenue', isBold: true },
    { item: 'Shipping Revenue', fy20: 421111, fy21: 514829, fy22: 625868, ttm: 680343, pct: 1, category: 'Revenue' },
    { item: 'Other Revenue', fy20: 0, fy21: -573711, fy22: -467278, ttm: -420876, pct: -1, category: 'Revenue' },
    { item: 'Total Net Revenue', fy20: 42187456, fy21: 51482903, fy22: 62745381, ttm: 68293742, pct: 100, category: 'Revenue', isBold: true },
    { item: 'Product Cost of Goods Sold', fy20: 12656238, fy21: 15444871, fy22: 18823614, ttm: 20488124, pct: 30, category: 'COGS' },
    { item: 'Warehousing & Fulfillment', fy20: 3374995, fy21: 4118632, fy22: 5019631, ttm: 5463500, pct: 8, category: 'COGS' },
    { item: 'Shipping & Freight Out', fy20: 2109372, fy21: 2574145, fy22: 3137272, ttm: 3414687, pct: 5, category: 'COGS' },
    { item: 'Inventory Adjustments', fy20: 84375, fy21: 102966, fy22: 125491, ttm: 136587, pct: 0, category: 'COGS' },
    { item: 'Total Cost of Goods Sold', fy20: 18224980, fy21: 22240614, fy22: 27106008, ttm: 29502898, pct: 43, category: 'COGS', isBold: true },
    { item: 'Gross Profit', fy20: 23962476, fy21: 29242289, fy22: 35639373, ttm: 38790844, pct: 57, category: 'Gross Profit', isBold: true },
    { item: 'Sales & Marketing', fy20: 10547114, fy21: 12870726, fy22: 15686345, ttm: 17073436, pct: 25, category: 'OpEx' },
    { item: 'Salaries, Wages & Payroll Tax', fy20: 2531923, fy21: 3088974, fy22: 3764574, ttm: 4097686, pct: 6, category: 'OpEx' },
    { item: 'Employee Benefits & 401k', fy20: 316490, fy21: 386122, fy22: 470572, ttm: 512211, pct: 1, category: 'OpEx' },
    { item: 'Rent & Occupancy', fy20: 336000, fy21: 441000, fy22: 457500, ttm: 468000, pct: 1, category: 'OpEx' },
    { item: 'Professional Services', fy20: 591123, fy21: 721083, fy22: 878636, ttm: 956314, pct: 1, category: 'OpEx' },
    { item: 'Technology & Software', fy20: 548274, fy21: 669045, fy22: 815229, ttm: 887662, pct: 1, category: 'OpEx' },
    { item: 'Insurance', fy20: 210937, fy21: 257418, fy22: 313636, ttm: 341484, pct: 0, category: 'OpEx' },
    { item: 'Payment Processing Fees', fy20: 1687498, fy21: 2059316, fy22: 2509815, ttm: 2731750, pct: 4, category: 'OpEx' },
    { item: 'Travel & Entertainment', fy20: 168750, fy21: 205932, fy22: 250982, ttm: 273175, pct: 0, category: 'OpEx' },
    { item: 'Office & General Admin', fy20: 253124, fy21: 308897, fy22: 376473, ttm: 409763, pct: 1, category: 'OpEx' },
    { item: 'Total Operating Expenses', fy20: 19331959, fy21: 23581462, fy22: 28721589, ttm: 31251028, pct: 46, category: 'OpEx', isBold: true },
    { item: 'Operating Income (EBITDA)', fy20: 4630517, fy21: 6177949, fy22: 7217823, ttm: 7960816, pct: 12, category: 'Operating Income', isBold: true },
    { item: 'Interest Expense', fy20: -42850, fy21: -52387, fy22: -61245, ttm: -65182, pct: 0, category: 'Other' },
    { item: 'Depreciation & Amortization', fy20: -6053, fy21: -6650, fy22: -6690, ttm: -6000, pct: 0, category: 'Other' },
    { item: 'Income Before Tax', fy20: 4581614, fy21: 6118912, fy22: 7149888, ttm: 7889634, pct: 12, category: 'Other', isBold: true },
    { item: 'Tax Provision', fy20: -23185, fy21: -27038, fy22: -29583, ttm: -31256, pct: 0, category: 'Other' },
    { item: 'Net Income', fy20: 4558429, fy21: 6091874, fy22: 7120305, ttm: 7858378, pct: 12, category: 'Net Income', isBold: true },
  ]

  // Balance Sheet Data
  const balanceSheetData = [
    { item: 'Cash & Cash Equivalents', jan21: 1547293, jan22: 2183746, jan23: 2847562, source: 'Bank Reconciliations - Wells Fargo Acct ***4729 & ***2183', category: 'Current Assets' },
    { item: 'Accounts Receivable', jan21: 1284738, jan22: 1586420, jan23: 1924857, source: 'A/R Aging Detail Report by Customer as of 12/31', category: 'Current Assets' },
    { item: 'Allowance for Doubtful Accounts', jan21: -25695, jan22: -31728, jan23: -38497, source: 'Historical Bad Debt Analysis - 2% reserve methodology', category: 'Current Assets' },
    { item: 'Net Accounts Receivable', jan21: 1259043, jan22: 1554692, jan23: 1886360, source: 'A/R net of allowance', category: 'Current Assets', isBold: true },
    { item: 'Inventory - Raw Materials', jan21: 284759, jan22: 352185, jan23: 428473, source: 'Physical Inventory Count & Valuation - FIFO method', category: 'Current Assets' },
    { item: 'Inventory - Finished Goods', jan21: 758024, jan22: 937493, jan23: 1140266, source: 'Perpetual Inventory System Report', category: 'Current Assets' },
    { item: 'Total Inventory', jan21: 1042783, jan22: 1289678, jan23: 1568739, source: 'Sum of inventory components', category: 'Current Assets', isBold: true },
    { item: 'Prepaid Expenses', jan21: 42187, jan22: 51483, jan23: 62745, source: 'Prepaid Schedule - Insurance, Software, Rent', category: 'Current Assets' },
    { item: 'Other Current Assets', jan21: 18500, jan22: 22800, jan23: 27700, source: 'Deposits & miscellaneous receivables', category: 'Current Assets' },
    { item: 'Total Current Assets', jan21: 3909806, jan22: 5102399, jan23: 6393106, source: 'Sum of all current assets', category: 'Current Assets', isBold: true },
    { item: 'Property, Plant & Equipment', jan21: 127450, jan22: 145280, jan23: 168950, source: 'Fixed Asset Register - Historical Cost', category: 'Fixed Assets' },
    { item: 'Less: Accumulated Depreciation', jan21: -89614, jan22: -94864, jan23: -99254, source: 'Accumulated Depreciation Schedule', category: 'Fixed Assets' },
    { item: 'Net Fixed Assets', jan21: 37836, jan22: 50416, jan23: 69696, source: 'PP&E net book value', category: 'Fixed Assets', isBold: true },
    { item: 'Intangible Assets - Software', jan21: 24000, jan22: 28000, jan23: 26000, source: 'Capitalized Software Development Costs', category: 'Fixed Assets' },
    { item: 'Total Assets', jan21: 3971642, jan22: 5180815, jan23: 6488802, source: 'Sum of all asset categories', category: 'Total Assets', isBold: true },
    { item: 'Accounts Payable - Trade', jan21: 789342, jan22: 976528, jan23: 1188247, source: 'A/P Aging by Vendor as of 12/31', category: 'Current Liabilities' },
    { item: 'Credit Cards Payable', jan21: 94531, jan22: 116966, jan23: 142295, source: 'Amex Corporate Card Statement - Balance Forward', category: 'Current Liabilities' },
    { item: 'Accrued Payroll & Benefits', jan21: 227384, jan22: 281292, jan23: 342187, source: 'Payroll Accrual Schedule - Last pay period + PTO liability', category: 'Current Liabilities' },
    { item: 'Accrued Expenses - Other', jan21: 168750, jan22: 208593, jan23: 253827, source: 'Accrued Utilities, Professional Fees, Marketing', category: 'Current Liabilities' },
    { item: 'Sales Tax Payable', jan21: 126656, jan22: 156735, jan23: 190734, source: 'Sales Tax Returns - State Filings (CA, TX, NY, FL)', category: 'Current Liabilities' },
    { item: 'Deferred Revenue', jan21: 0, jan22: 85000, jan23: 127500, source: 'Customer Deposits & Prepayments', category: 'Current Liabilities' },
    { item: 'Current Portion - LT Debt', jan21: 125000, jan22: 125000, jan23: 125000, source: 'Term Loan Agreement - Bank of America', category: 'Current Liabilities' },
    { item: 'Total Current Liabilities', jan21: 1531663, jan22: 1950114, jan23: 2369790, source: 'Sum of all current liabilities', category: 'Current Liabilities', isBold: true },
    { item: 'Long-Term Debt, Net of Current', jan21: 875000, jan22: 750000, jan23: 625000, source: 'Term Loan Schedule - Original $1.5M, 5yr amortization', category: 'LT Liabilities' },
    { item: 'Total Liabilities', jan21: 2406663, jan22: 2700114, jan23: 2994790, source: 'Sum of current + long-term liabilities', category: 'Total Liabilities', isBold: true },
    { item: 'Common Stock', jan21: 100000, jan22: 100000, jan23: 100000, source: 'Certificate of Incorporation - 10,000 shares authorized, 1,000 issued', category: 'Equity' },
    { item: 'Additional Paid-in Capital', jan21: 0, jan22: 0, jan23: 0, source: 'No additional capital contributions', category: 'Equity' },
    { item: 'Retained Earnings', jan21: -93450, jan22: 1288827, jan23: 2673707, source: 'Retained Earnings Roll-Forward', category: 'Equity' },
    { item: 'Current Year Net Income', jan21: 1558429, jan22: 2091874, jan23: 2720305, source: 'Net Income per books', category: 'Equity' },
    { item: 'Total Stockholders\' Equity', jan21: 1564979, jan22: 2480701, jan23: 3494012, source: 'Sum of equity accounts', category: 'Equity', isBold: true },
    { item: 'Total Liabilities & Equity', jan21: 3971642, jan22: 5180815, jan23: 6488802, source: 'Must equal Total Assets', category: 'Total', isBold: true },
  ]

  // Sales by Channel Data (TTM Jan-23)
  const salesChannelData = [
    { product: 'Premium Series A - Commercial', revenue: 10244061, pct: 15.0, qty: 58520, avgPrice: 175.05, source: 'NetSuite Sales Report by SKU - Item #PS-A-COM-001' },
    { product: 'Standard Line B - Residential', revenue: 8195249, pct: 12.0, qty: 124857, avgPrice: 65.65, source: 'NetSuite Sales Report by SKU - Item #STD-B-RES-005' },
    { product: 'Pro Edition C - Industrial', revenue: 6829374, pct: 10.0, qty: 42684, avgPrice: 160.00, source: 'NetSuite Sales Report by SKU - Item #PRO-C-IND-120' },
    { product: 'Value Pack D - Bulk', revenue: 6146436, pct: 9.0, qty: 256893, avgPrice: 23.93, source: 'NetSuite Sales Report by SKU - Item #VAL-D-BLK-250' },
    { product: 'Elite Model E - Enterprise', revenue: 5463498, pct: 8.0, qty: 21859, avgPrice: 249.95, source: 'NetSuite Sales Report by SKU - Item #ELT-E-ENT-999' },
    { product: 'Classic Series F - Commercial', revenue: 4780562, pct: 7.0, qty: 95611, avgPrice: 50.00, source: 'NetSuite Sales Report by SKU - Item #CLS-F-COM-050' },
    { product: 'Compact Unit G - Residential', revenue: 4097623, pct: 6.0, qty: 136587, avgPrice: 30.00, source: 'NetSuite Sales Report by SKU - Item #CMP-G-RES-030' },
    { product: 'Advanced Kit H - Industrial', revenue: 3414687, pct: 5.0, qty: 42684, avgPrice: 80.00, source: 'NetSuite Sales Report by SKU - Item #ADV-H-IND-080' },
    { product: 'Starter Set I - Entry Level', revenue: 2731749, pct: 4.0, qty: 227323, avgPrice: 12.02, source: 'NetSuite Sales Report by SKU - Item #STR-I-ENT-012' },
    { product: 'Professional J - Specialty', revenue: 2048812, pct: 3.0, qty: 20488, avgPrice: 100.00, source: 'NetSuite Sales Report by SKU - Item #PRF-J-SPC-100' },
    { product: 'All Other Products (840 SKUs)', revenue: 14341691, pct: 21.0, qty: 487382, avgPrice: 29.43, source: 'NetSuite Sales Summary - Remaining SKUs aggregated' },
    { product: 'Total Product Sales', revenue: 68293742, pct: 100.0, qty: 1514888, avgPrice: 45.08, source: 'NetSuite TTM Revenue Roll-up Report', isBold: true },
  ]

  // Margins by Month Data (Last 12 months)
  const months = ['Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan']
  const marginsByMonthData = [
    { month: 'Feb', revenue: 5284750, cogs: 2272443, opex: 2431065 },
    { month: 'Mar', revenue: 5692410, cogs: 2447656, opex: 2619509 },
    { month: 'Apr', revenue: 5418293, cogs: 2329886, opex: 2492354 },
    { month: 'May', revenue: 5651847, cogs: 2430294, opex: 2599850 },
    { month: 'Jun', revenue: 5892156, cogs: 2533827, opex: 2710593 },
    { month: 'Jul', revenue: 5729384, cogs: 2463875, opex: 2635514 },
    { month: 'Aug', revenue: 5968247, cogs: 2566346, opex: 2745673 },
    { month: 'Sep', revenue: 5647291, cogs: 2428335, opex: 2597756 },
    { month: 'Oct', revenue: 5915832, cogs: 2543868, opex: 2723934 },
    { month: 'Nov', revenue: 5738549, cogs: 2467819, opex: 2639736 },
    { month: 'Dec', revenue: 5692847, cogs: 2447824, opex: 2619691 },
    { month: 'Jan', revenue: 5662136, cogs: 2434613, opex: 2605351 },
  ]

  // Proof of Cash Data
  const proofOfCashData = [
    { month: 'Feb', deposits: 5318475, beginningAR: 1586420, endingAR: 1624738, revenueGL: 5284750, nonRevDeposits: 72043, variance: 0 },
    { month: 'Mar', deposits: 5728549, beginningAR: 1624738, endingAR: 1668294, revenueGL: 5692410, nonRevDeposits: 79985, variance: 0 },
    { month: 'Apr', deposits: 5447182, beginningAR: 1668294, endingAR: 1705847, revenueGL: 5418293, nonRevDeposits: 66636, variance: 0 },
    { month: 'May', deposits: 5682594, beginningAR: 1705847, endingAR: 1745938, revenueGL: 5651847, nonRevDeposits: 70838, variance: 0 },
    { month: 'Jun', deposits: 5924738, beginningAR: 1745938, endingAR: 1788562, revenueGL: 5892156, nonRevDeposits: 78218, variance: 0 },
    { month: 'Jul', deposits: 5761847, beginningAR: 1788562, endingAR: 1826485, revenueGL: 5729384, nonRevDeposits: 74528, variance: 0 },
    { month: 'Aug', deposits: 6002185, beginningAR: 1826485, endingAR: 1868394, revenueGL: 5968247, nonRevDeposits: 85341, variance: 0 },
    { month: 'Sep', deposits: 5677582, beginningAR: 1868394, endingAR: 1904728, revenueGL: 5647291, nonRevDeposits: 67025, variance: 0 },
    { month: 'Oct', deposits: 5948294, beginningAR: 1904728, endingAR: 1947183, revenueGL: 5915832, nonRevDeposits: 75117, variance: 0 },
    { month: 'Nov', deposits: 5770482, beginningAR: 1947183, endingAR: 1985946, revenueGL: 5738549, nonRevDeposits: 70154, variance: 0 },
    { month: 'Dec', deposits: 5725947, beginningAR: 1985946, endingAR: 2018046, revenueGL: 5692847, nonRevDeposits: 65146, variance: 0 },
    { month: 'Jan', deposits: 5693728, beginningAR: 2018046, endingAR: 2053638, revenueGL: 5662136, nonRevDeposits: 67046, variance: 0 },
  ]

  // Working Capital Data
  const workingCapitalData = [
    { item: 'Accounts Receivable', current: 1886360, normalized: 1823947, adjustment: -62413, days: 10.1, targetDays: 9.8, source: 'A/R Aging Report - DSO calculation based on TTM revenue $68.3M' },
    { item: 'Inventory', current: 1568739, normalized: 1687429, adjustment: 118690, days: 19.4, targetDays: 20.9, source: 'Inventory Analysis - DIO calculation vs. COGS run-rate' },
    { item: 'Prepaid Expenses', current: 62745, normalized: 62745, adjustment: 0, days: 0.3, targetDays: 0.3, source: 'Prepaid Schedule - Insurance, Rent, Software normalized' },
    { item: 'Other Current Assets', current: 27700, normalized: 27700, adjustment: 0, days: 0.1, targetDays: 0.1, source: 'Deposits & misc receivables - no adjustment required' },
    { item: 'Accounts Payable', current: -1188247, normalized: -1247586, adjustment: -59339, days: 14.7, targetDays: 15.5, source: 'A/P Aging Report - DPO calculation vs. COGS + OpEx' },
    { item: 'Accrued Expenses', current: -342187, normalized: -342187, adjustment: 0, days: 1.8, targetDays: 1.8, source: 'Payroll & expense accruals normalized to run-rate' },
    { item: 'Credit Cards Payable', current: -142295, normalized: -142295, adjustment: 0, days: 0.8, targetDays: 0.8, source: 'Corporate cards - normal operating level' },
    { item: 'Sales Tax Payable', current: -190734, normalized: -190734, adjustment: 0, days: 1.0, targetDays: 1.0, source: 'Sales tax liability - statutory requirement' },
    { item: 'Deferred Revenue', current: -127500, normalized: -127500, adjustment: 0, days: 0.7, targetDays: 0.7, source: 'Customer deposits - no normalization' },
    { item: 'Adjusted Net Working Capital', current: 1554581, normalized: 1550519, adjustment: -4062, days: 8.3, targetDays: 8.3, source: 'Sum of normalized NWC components', isBold: true },
  ]

  // COGS Vendors Data
  const cogsVendorsData = [
    { vendor: 'Global Manufacturing Inc.', fy20: 3647797, fy21: 4448123, fy22: 5421201, ttm: 5900580, pct: 20.0, source: 'A/P Subledger by Vendor - Primary component supplier (China)' },
    { vendor: 'Pacific Logistics Group', fy20: 2738348, fy21: 3336092, fy22: 4065901, ttm: 4425435, pct: 15.0, source: 'A/P Subledger - Freight forwarding & warehousing services' },
    { vendor: 'Midwest Distribution Co.', fy20: 2190678, fy21: 2668877, fy22: 3253921, ttm: 3540348, pct: 12.0, source: 'A/P Subledger - Fulfillment & 3PL services (7 warehouses)' },
    { vendor: 'TechComponents LLC', fy20: 1825562, fy21: 2224061, fy22: 2712601, ttm: 2951348, pct: 10.0, source: 'A/P Subledger - Electronic components & assemblies' },
    { vendor: 'Atlantic Packaging Solutions', fy20: 1642486, fy21: 2001907, fy22: 2440741, ttm: 2655261, pct: 9.0, source: 'A/P Subledger - Custom packaging materials & design' },
    { vendor: 'Sterling Materials Group', fy20: 1095991, fy21: 1334604, fy22: 1626361, ttm: 1770174, pct: 6.0, source: 'A/P Subledger - Raw materials supplier (metals & plastics)' },
    { vendor: 'Express Freight Services', fy20: 912493, fy21: 1112170, fy22: 1355300, ttm: 1475145, pct: 5.0, source: 'A/P Subledger - Last-mile delivery & expedited shipping' },
    { vendor: 'Quality Control Labs', fy20: 729994, fy21: 889336, fy22: 1084241, ttm: 1180116, pct: 4.0, source: 'A/P Subledger - Product testing & certifications (ISO/UL)' },
    { vendor: 'All Other Vendors (143 vendors)', fy20: 3442631, fy21: 4196444, fy22: 5145741, ttm: 5604491, pct: 19.0, source: 'A/P Subledger Summary - Remaining supplier base' },
    { vendor: 'Total COGS', fy20: 18224980, fy21: 22240614, fy22: 27106008, ttm: 29502898, pct: 100.0, source: 'Total COGS per Income Statement', isBold: true },
  ]

  // AP/Accrual Testing Data
  const testingData = [
    { checkNum: 'ACH-18472', invoice: 'GMI-2023-0847', payee: 'Global Manufacturing Inc.', checkDate: '2023-01-05', checkAmt: 287450, periodAmt: 287450, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18485', invoice: 'PLG-INV-4729', payee: 'Pacific Logistics Group', checkDate: '2023-01-08', checkAmt: 156780, periodAmt: 156780, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18503', invoice: 'MDC-2023-0194', payee: 'Midwest Distribution Co.', checkDate: '2023-01-12', checkAmt: 198650, periodAmt: 198650, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18518', invoice: 'TC-47283', payee: 'TechComponents LLC', checkDate: '2023-01-15', checkAmt: 124830, periodAmt: 124830, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'CHK-9284', invoice: 'APS-JAN-2023', payee: 'Atlantic Packaging Solutions', checkDate: '2023-01-18', checkAmt: 89470, periodAmt: 89470, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18547', invoice: 'SMG-8472-A', payee: 'Sterling Materials Group', checkDate: '2023-01-22', checkAmt: 67240, periodAmt: 67240, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'CHK-9301', invoice: 'EFS-012023', payee: 'Express Freight Services', checkDate: '2023-01-25', checkAmt: 52830, periodAmt: 52830, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18574', invoice: 'QCL-2023-Q1', payee: 'Quality Control Labs', checkDate: '2023-01-28', checkAmt: 41850, periodAmt: 18950, futureAmt: 22900, notAccrued: 0 },
    { checkNum: 'CHK-9318', invoice: 'AWS-JAN-2023', payee: 'Amazon Web Services', checkDate: '2023-01-30', checkAmt: 28470, periodAmt: 28470, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18592', invoice: 'SHP-847291', payee: 'Shopify Inc.', checkDate: '2023-01-31', checkAmt: 15680, periodAmt: 15680, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'CHK-9325', invoice: 'INS-Q1-2023', payee: 'Liberty Mutual Insurance', checkDate: '2023-02-01', checkAmt: 47250, periodAmt: 0, futureAmt: 47250, notAccrued: 0 },
    { checkNum: 'ACH-18608', invoice: 'META-JAN-2023', payee: 'Meta Platforms Inc.', checkDate: '2023-02-02', checkAmt: 68940, periodAmt: 68940, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'ACH-18621', invoice: 'GOOG-ADV-847', payee: 'Google LLC', checkDate: '2023-02-03', checkAmt: 124680, periodAmt: 124680, futureAmt: 0, notAccrued: 0 },
    { checkNum: 'CHK-9342', invoice: 'CTR-FEB-2023', payee: 'Smith & Associates Consulting', checkDate: '2023-02-05', checkAmt: 35800, periodAmt: 0, futureAmt: 35800, notAccrued: 0 },
    { checkNum: 'ACH-18647', invoice: 'SAL-JAN-2023', payee: 'Salesforce.com Inc.', checkDate: '2023-02-06', checkAmt: 12450, periodAmt: 12450, futureAmt: 0, notAccrued: 0 },
  ]

  const testingTotals = {
    totalTested: testingData.reduce((sum, row) => sum + row.checkAmt, 0),
    totalDisbursements: 5847293,
    percentTested: 23.7,
    totalError: 0,
  }

  const renderContent = () => {
    switch (activeSection) {
      case 'overview':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Overview</h2>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Metric ($ in USD)</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">TTM Jan-23</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {overviewData.map((row, idx) => (
                    <TableRow key={idx} className={row.metric.includes('Adjusted EBITDA') && !row.metric.includes('Margin') ? 'bg-muted/50' : ''}>
                      <TableCell className="font-medium">{row.metric}</TableCell>
                      <TableCell className="text-right font-mono">
                        {row.isPercent ? formatPercent(row.fy20) : formatCurrency(row.fy20)}
                      </TableCell>
                      <TableCell className="text-right font-mono">
                        {row.isPercent ? formatPercent(row.fy21) : formatCurrency(row.fy21)}
                      </TableCell>
                      <TableCell className="text-right font-mono">
                        {row.isPercent ? formatPercent(row.fy22) : formatCurrency(row.fy22)}
                      </TableCell>
                      <TableCell className="text-right font-mono">
                        {row.isPercent ? formatPercent(row.ttm) : formatCurrency(row.ttm)}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'qoe': {
        const qoeTtmRev = getLiveValue('overview', 'total_revenue', 'TTM') ?? 68293742
        const qoeTtmBase = getLiveValue('qoe', 'ebitda_as_defined', 'TTM') ?? 7960816
        const qoeTtmAdj = getLiveValue('qoe', 'diligence_adjusted_ebitda', 'TTM') ?? 8547239
        const qoeNetAdj = qoeTtmAdj - qoeTtmBase
        const qoeAdjMargin = qoeTtmRev > 0 ? (qoeTtmAdj / qoeTtmRev) * 100 : 12.5

        // Build TTM EBITDA waterfall — range bar approach: [start, end] per bar
        type WfItem = { name: string; value: number; isBase?: boolean; isEnd?: boolean }
        type WfPoint = { name: string; range: [number, number]; rawValue: number; isBase: boolean; isEnd: boolean }
        const qoeWfItems: WfItem[] = [
          { name: 'EBITDA as Defined', value: qoeTtmBase, isBase: true },
          { name: 'a) Charitable', value: getLiveValue('qoe', 'adj_charitable_contributions', 'TTM') ?? 315920 },
          { name: 'b) Owner Tax', value: getLiveValue('qoe', 'adj_owner_tax_legal', 'TTM') ?? 22450 },
          { name: 'c) Excess Comp', value: getLiveValue('qoe', 'adj_excess_owner_comp', 'TTM') ?? 248190 },
          { name: 'd) Personal Exp', value: getLiveValue('qoe', 'adj_personal_expenses', 'TTM') ?? 39863 },
          { name: 'e) Prof. Fees', value: getLiveValue('qoe', 'adj_professional_fees_onetime', 'TTM') ?? 68200 },
          { name: 'f) Exec. Search', value: getLiveValue('qoe', 'adj_executive_search', 'TTM') ?? 52000 },
          { name: 'g) Relocation', value: getLiveValue('qoe', 'adj_facility_relocation', 'TTM') ?? 41500 },
          { name: 'h) Severance', value: getLiveValue('qoe', 'adj_severance', 'TTM') ?? 92300 },
          { name: 'i) Above-Mkt Rent', value: getLiveValue('qoe', 'adj_above_market_rent', 'TTM') ?? -168900 },
          { name: 'j) Owner Comp', value: getLiveValue('qoe', 'adj_normalize_owner_comp', 'TTM') ?? -425900 },
          { name: 'Adj. EBITDA', value: 0, isEnd: true },
        ]
        let qoeWfRunning = 0
        const qoeWfData: WfPoint[] = qoeWfItems.map((item) => {
          if (item.isBase) { qoeWfRunning = item.value; return { name: item.name, range: [0, item.value], rawValue: item.value, isBase: true, isEnd: false } }
          if (item.isEnd) { return { name: item.name, range: [0, qoeWfRunning], rawValue: qoeWfRunning, isBase: false, isEnd: true } }
          const lo = item.value >= 0 ? qoeWfRunning : qoeWfRunning + item.value
          const hi = item.value >= 0 ? qoeWfRunning + item.value : qoeWfRunning
          const pt: WfPoint = { name: item.name, range: [lo, hi], rawValue: item.value, isBase: false, isEnd: false }
          qoeWfRunning += item.value
          return pt
        })
        const qoeWfYMin = qoeTtmBase * 0.97
        const qoeWfYMax = Math.max(...qoeWfData.map((p) => p.range[1])) * 1.02

        // Margin trend
        const qoeMarginTrend = [
          { period: 'FY20', unadj: (4630517 / 42187456) * 100, adj: (4875291 / 42187456) * 100 },
          { period: 'FY21', unadj: (6177949 / 51482903) * 100, adj: (6524835 / 51482903) * 100 },
          { period: 'FY22', unadj: (7217823 / 62745381) * 100, adj: (7682947 / 62745381) * 100 },
          { period: 'TTM',  unadj: qoeTtmRev > 0 ? (qoeTtmBase / qoeTtmRev) * 100 : 11.7, adj: qoeAdjMargin },
        ]

        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Quality of Earnings</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />
                Loading live data...
              </div>
            )}

            {/* KPI Cards */}
            <div className="grid grid-cols-4 gap-3 mb-6">
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">TTM Revenue</p>
                <p className="text-xl font-bold">{formatCurrency(qoeTtmRev)}</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">EBITDA as Defined</p>
                <p className="text-xl font-bold">{formatCurrency(qoeTtmBase)}</p>
                <p className="text-xs text-muted-foreground mt-1">{formatPercent(qoeTtmRev > 0 ? (qoeTtmBase / qoeTtmRev) * 100 : 11.7)} margin</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">Adj. EBITDA</p>
                <p className="text-xl font-bold text-[#1E3A5F]">{formatCurrency(qoeTtmAdj)}</p>
                <p className="text-xs text-muted-foreground mt-1">{formatPercent(qoeAdjMargin)} margin</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">Total Net Adjustments</p>
                <p className="text-xl font-bold text-emerald-600">+{formatCurrency(qoeNetAdj)}</p>
                <p className="text-xs text-muted-foreground mt-1">10 discrete add-backs</p>
              </div>
            </div>

            {/* Charts row */}
            <div className="grid grid-cols-2 gap-6 mb-8">
              {/* EBITDA Waterfall */}
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-3">EBITDA Bridge — TTM</h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={qoeWfData} margin={{ top: 5, right: 10, left: 20, bottom: 44 }} barCategoryGap="20%">
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-40} textAnchor="end" interval={0} />
                      <YAxis domain={[qoeWfYMin, qoeWfYMax]} tickFormatter={(v) => `$${(v / 1e6).toFixed(1)}M`} tick={{ fontSize: 9 }} width={55} />
                      <Tooltip
                        content={({ active, payload, label }) => {
                          if (!active || !payload?.length) return null
                          const d = (payload[0] as { payload: WfPoint }).payload
                          return (
                            <div className="bg-white border border-gray-200 rounded px-3 py-2 text-xs shadow">
                              <div className="font-semibold mb-1">{label}</div>
                              <div>{d.isBase || d.isEnd ? formatCurrency(d.range[1]) : `${d.rawValue >= 0 ? '+' : ''}${formatCurrency(d.rawValue)}`}</div>
                            </div>
                          )
                        }}
                      />
                      <Bar dataKey="range">
                        {qoeWfData.map((entry, idx) => (
                          <Cell key={idx} fill={entry.isBase || entry.isEnd ? '#1E3A5F' : entry.rawValue >= 0 ? '#10b981' : '#ef4444'} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Margin Trend */}
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-3">EBITDA Margin Trend</h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={qoeMarginTrend} margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={(v: number) => `${v.toFixed(1)}%`} tick={{ fontSize: 10 }} width={45} domain={[9, 15]} />
                      <Tooltip formatter={(v: unknown) => [`${typeof v === 'number' ? v.toFixed(1) : v}%`]} />
                      <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'adj' ? 'Adj. EBITDA Margin' : 'EBITDA as Defined Margin'} />
                      <Line type="monotone" dataKey="unadj" name="unadj" stroke="#94a3b8" strokeWidth={2} dot={{ r: 4 }} />
                      <Line type="monotone" dataKey="adj" name="adj" stroke="#1E3A5F" strokeWidth={2} dot={{ r: 4 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Full Adjustment Schedule */}
            <h3 className="text-base font-semibold mb-3">Full Adjustment Schedule</h3>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">TTM Jan-23</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {qoeData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    return (
                    <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                      <TableCell className={row.isAdjustment ? 'pl-8 italic text-sm' : row.isBold ? 'font-semibold' : ''}>
                        {row.item}
                      </TableCell>
                      {(['FY20', 'FY21', 'FY22', 'TTM'] as const).map((period, pi) => {
                        const fallbackVal = pi === 0 ? row.fy20 : pi === 1 ? row.fy21 : pi === 2 ? row.fy22 : row.ttm
                        const displayVal = getDisplayValue('qoe', rowKey, period, fallbackVal)
                        const srcRef = getCellSourceRef('qoe', rowKey, period)
                        const cellId = getCellId('qoe', rowKey, period)
                        return (
                          <TableCell key={period} className="text-right font-mono">
                            <AuditableCell
                              value={displayVal}
                              source={row.source}
                              sourceRef={srcRef}
                              cellId={cellId}
                              workbookId={isDemoWorkbook ? undefined : id}
                              isEditable={!row.isBold}
                              onViewSource={handleViewSource}
                            />
                          </TableCell>
                        )
                      })}
                    </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )
      }

      case 'income-statement':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Income Statement - Adjusted</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />Loading live data...
              </div>
            )}
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">TTM Jan-23</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">% of Sales</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {incomeStatementData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        {(['FY20', 'FY21', 'FY22', 'TTM'] as const).map((period, pi) => {
                          const fallback = pi === 0 ? row.fy20 : pi === 1 ? row.fy21 : pi === 2 ? row.fy22 : row.ttm
                          const val = getLiveValue('income-statement', rowKey, period) ?? fallback
                          const srcRef = getCellSourceRef('income-statement', rowKey, period)
                          const cellId = getCellId('income-statement', rowKey, period)
                          return (
                            <TableCell key={period} className="text-right font-mono text-sm">
                              <AuditableCell
                                value={formatCurrency(val)}
                                source=""
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                isEditable={!row.isBold}
                                onViewSource={handleViewSource}
                              />
                            </TableCell>
                          )
                        })}
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.pct}%
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'balance-sheet':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Balance Sheet - Adjusted</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />Loading live data...
              </div>
            )}
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2021</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2022</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2023</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {balanceSheetData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    const bsPeriods = ['FY20', 'FY21', 'FY22'] as const
                    const fallbacks = [row.jan21, row.jan22, row.jan23]
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        {bsPeriods.map((period, pi) => {
                          const displayVal = getDisplayValue('balance-sheet', rowKey, period, fallbacks[pi])
                          const srcRef = getCellSourceRef('balance-sheet', rowKey, period)
                          const cellId = getCellId('balance-sheet', rowKey, period)
                          return (
                            <TableCell key={period} className="text-right font-mono">
                              <AuditableCell
                                value={displayVal}
                                source={row.source}
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                isEditable={!row.isBold}
                                onViewSource={handleViewSource}
                              />
                            </TableCell>
                          )
                        })}
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'sales-channel':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Sales by Channel</h2>
            <p className="text-sm text-muted-foreground mb-4">Product mix and SKU analysis for TTM Jan-23</p>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[200px]">Product/SKU</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Revenue</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">% of Total</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Qty Sold</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Avg Price</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {salesChannelData.map((row, idx) => {
                    const rowKey = row.product.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 40)
                    const displayVal = getDisplayValue('sales-channel', rowKey, 'TTM', row.revenue)
                    const srcRef = getCellSourceRef('sales-channel', rowKey, 'TTM')
                    const cellId = getCellId('sales-channel', rowKey, 'TTM')
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.product}</TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell
                            value={displayVal}
                            source={row.source}
                            sourceRef={srcRef}
                            cellId={cellId}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono">{formatPercent(row.pct)}</TableCell>
                        <TableCell className="text-right font-mono">{row.qty.toLocaleString()}</TableCell>
                        <TableCell className="text-right font-mono">{formatCurrency(row.avgPrice)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'margins-month':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Margins by Month</h2>
            <p className="text-sm text-muted-foreground mb-4">Monthly profitability analysis - Last 12 months</p>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Month</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Revenue</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">COGS</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">OpEx</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Gross Margin</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">GM %</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Net Margin</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">NM %</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {marginsByMonthData.map((row, idx) => {
                    // Use live data for revenue/cogs/opex if available, keyed by month
                    const liveRev = getLiveValue('margins-month', 'revenue', row.month) ?? row.revenue
                    const liveCogs = getLiveValue('margins-month', 'cogs', row.month) ?? row.cogs
                    const liveOpex = getLiveValue('margins-month', 'opex', row.month) ?? row.opex
                    const grossMargin = liveRev - liveCogs
                    const netMargin = liveRev - liveCogs - liveOpex
                    const gmPct = liveRev > 0 ? (grossMargin / liveRev) * 100 : 0
                    const nmPct = liveRev > 0 ? (netMargin / liveRev) * 100 : 0
                    return (
                      <TableRow key={idx}>
                        <TableCell className="font-medium">{row.month}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveRev)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveCogs)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveOpex)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(grossMargin)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(gmPct)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(netMargin)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(nmPct)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'proof-cash':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Proof of Cash</h2>
            <p className="text-sm text-muted-foreground mb-4">Bank deposit reconciliation to reported revenue</p>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Month</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Bank Deposits</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Less: Beg AR</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Add: End AR</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Less: Non-Rev</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Revenue GL</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Variance</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {proofOfCashData.map((row, idx) => {
                    const liveDeposits = getLiveValue('proof-cash', 'bank_deposits', row.month) ?? row.deposits
                    const liveBegAR = getLiveValue('proof-cash', 'beginning_ar', row.month) ?? row.beginningAR
                    const liveEndAR = getLiveValue('proof-cash', 'ending_ar', row.month) ?? row.endingAR
                    const liveNonRev = getLiveValue('proof-cash', 'non_rev_deposits', row.month) ?? row.nonRevDeposits
                    const liveRevGL = getLiveValue('proof-cash', 'revenue_gl', row.month) ?? row.revenueGL
                    const calculated = liveDeposits - liveBegAR + liveEndAR - liveNonRev
                    const variance = calculated - liveRevGL
                    return (
                      <TableRow key={idx}>
                        <TableCell className="font-medium">{row.month}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveDeposits)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveBegAR)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveEndAR)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveNonRev)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveRevGL)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          <span className={Math.abs(variance) > 50000 ? 'text-red-500' : 'text-green-600'}>
                            {formatCurrency(variance)}
                          </span>
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'working-capital':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Adjusted Net Working Capital Analysis</h2>
            <p className="text-sm text-muted-foreground mb-4">Working capital components with DSO, DPO, and DOH metrics</p>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[250px]">Component</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Current</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">Days</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">Target</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Normalized</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Adjustment</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {workingCapitalData.map((row, idx) => {
                    const rowKey = 'wc_' + row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    const currentVal = getLiveValue('working-capital', rowKey, 'TTM') ?? row.current
                    const normalizedVal = getLiveValue('working-capital', rowKey + '_normalized', 'TTM') ?? row.normalized
                    const adjustment = normalizedVal - currentVal
                    const srcRef = getCellSourceRef('working-capital', rowKey, 'TTM')
                    const cellId = getCellId('working-capital', rowKey, 'TTM')
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell
                            value={formatCurrency(currentVal)}
                            source={row.source}
                            sourceRef={srcRef}
                            cellId={cellId}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.days > 0 ? row.days.toFixed(1) : '-'}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.targetDays > 0 ? row.targetDays.toFixed(1) : '-'}
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell value={formatCurrency(normalizedVal)} source={row.source} isEditable={false} />
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          {adjustment !== 0 ? formatCurrency(adjustment) : '-'}
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'cogs-vendors':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">COGS Expenditures by Vendor</h2>
            <p className="text-sm text-muted-foreground mb-4">Vendor concentration and supplier analysis</p>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[200px]">Vendor</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">TTM Jan-23</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">% of COGS</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {cogsVendorsData.map((row, idx) => {
                    const rowKey = 'vendor_' + row.vendor.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 30)
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.vendor}</TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY20') ?? row.fy20)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY21') ?? row.fy21)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY22') ?? row.fy22)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          <AuditableCell
                            value={formatCurrency(getLiveValue('cogs-vendors', rowKey, 'TTM') ?? row.ttm)}
                            source={row.source}
                            sourceRef={getCellSourceRef('cogs-vendors', rowKey, 'TTM')}
                            cellId={getCellId('cogs-vendors', rowKey, 'TTM')}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(row.pct)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'testing':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Accounts Payable / Accrued Expense Testing</h2>
            <p className="text-sm text-muted-foreground mb-4">Sample testing of vendor invoices and disbursements</p>
            <div className="rounded-md border overflow-x-auto mb-6">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Check #</TableHead>
                    <TableHead className="font-semibold w-[100px]">Invoice #</TableHead>
                    <TableHead className="font-semibold w-[180px]">Payee</TableHead>
                    <TableHead className="font-semibold w-[110px]">Date</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Check Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Period Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Future Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Not Accrued</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {testingData.map((row, idx) => (
                    <TableRow key={idx}>
                      <TableCell className="font-mono text-sm">{row.checkNum}</TableCell>
                      <TableCell className="font-mono text-sm">{row.invoice}</TableCell>
                      <TableCell className="text-sm">{row.payee}</TableCell>
                      <TableCell className="text-sm">{row.checkDate}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.checkAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.periodAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.futureAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">
                        {row.notAccrued > 0 ? (
                          <span className="text-red-500">{formatCurrency(row.notAccrued)}</span>
                        ) : '-'}
                      </TableCell>
                    </TableRow>
                  ))}
                  <TableRow className="bg-muted/50 font-semibold">
                    <TableCell colSpan={4}>Testing Summary</TableCell>
                    <TableCell className="text-right font-mono">{formatCurrency(testingTotals.totalTested)}</TableCell>
                    <TableCell colSpan={2} className="text-right">
                      <span className="text-sm text-muted-foreground">
                        {testingTotals.percentTested}% of {formatCurrency(testingTotals.totalDisbursements)} tested
                      </span>
                    </TableCell>
                    <TableCell className="text-right font-mono">
                      {testingTotals.totalError > 0 ? (
                        <span className="text-red-500">{formatCurrency(testingTotals.totalError)}</span>
                      ) : (
                        <span className="text-green-600">$0</span>
                      )}
                    </TableCell>
                  </TableRow>
                </TableBody>
              </Table>
            </div>
            <div className="flex items-center gap-2 text-sm">
              <Badge className="bg-green-600">No Material Errors Found</Badge>
              <span className="text-muted-foreground">Sample size: {testingTotals.percentTested}% of total disbursements</span>
            </div>
          </div>
        )

      case 'charts': {
        const chartRevEbitda = [
          { period: 'FY20', revenue: 42.19, margin: 11.6 },
          { period: 'FY21', revenue: 51.48, margin: 12.7 },
          { period: 'FY22', revenue: 62.75, margin: 12.2 },
          { period: 'TTM',  revenue: 68.29, margin: 12.5 },
        ]
        const chartMonthly = [
          { month: 'Feb', revenue: 5.28, ebitda: 0.58 },
          { month: 'Mar', revenue: 5.69, ebitda: 0.63 },
          { month: 'Apr', revenue: 5.44, ebitda: 0.60 },
          { month: 'May', revenue: 5.85, ebitda: 0.64 },
          { month: 'Jun', revenue: 5.53, ebitda: 0.61 },
          { month: 'Jul', revenue: 5.62, ebitda: 0.62 },
          { month: 'Aug', revenue: 5.75, ebitda: 0.63 },
          { month: 'Sep', revenue: 5.49, ebitda: 0.61 },
          { month: 'Oct', revenue: 5.96, ebitda: 0.66 },
          { month: 'Nov', revenue: 5.38, ebitda: 0.59 },
          { month: 'Dec', revenue: 5.66, ebitda: 0.62 },
          { month: 'Jan', revenue: 5.66, ebitda: 0.62 },
        ]
        const chartCustomers = [
          { name: 'Acme Manufacturing', value: 14.2, pct: 20.8 },
          { name: 'Nexgen Solutions',   value: 9.6,  pct: 14.0 },
          { name: 'Bridgewater Corp',   value: 7.5,  pct: 11.0 },
          { name: 'Summit Enterprises', value: 5.5,  pct: 8.0  },
          { name: 'Crestwood Partners', value: 4.1,  pct: 6.0  },
          { name: 'Lakeside Industries',value: 3.4,  pct: 5.0  },
          { name: 'Horizon Group',      value: 2.7,  pct: 4.0  },
          { name: 'Pacific Dynamics',   value: 2.0,  pct: 3.0  },
          { name: 'Allied Services',    value: 1.4,  pct: 2.0  },
          { name: 'Meridian LLC',       value: 0.7,  pct: 1.0  },
        ]
        const chartCashConv = [
          { period: 'FY22', ebitda: 7.68, fcf: 6.74 },
          { period: 'TTM',  ebitda: 8.55, fcf: 7.46 },
        ]
        const chartMarginTrend = marginsByMonthData.map(m => ({
          month: m.month,
          grossMargin: ((m.revenue - m.cogs) / m.revenue) * 100,
          netMargin: ((m.revenue - m.cogs - m.opex) / m.revenue) * 100,
        }))

        return (
          <div className="max-w-5xl space-y-10">
            <div>
              <h2 className="text-2xl font-bold mb-1">Charts &amp; Analytics</h2>
              <p className="text-sm text-muted-foreground">Financial metrics and performance indicators</p>
            </div>

            {/* KPI Cards */}
            <div className="grid grid-cols-4 gap-4">
              {[
                { label: 'TTM Revenue',    value: '$68.3M', sub: '+32.6% vs. FY21' },
                { label: 'Adj. EBITDA',    value: '$8.55M', sub: '12.5% margin' },
                { label: 'Gross Profit',   value: '$38.8M', sub: '56.8% gross margin' },
                { label: 'FCF Conversion', value: '87.2%',  sub: '$7.46M free cash flow' },
              ].map((card) => (
                <div key={card.label} className="bg-slate-50 rounded p-4">
                  <p className="text-xs text-muted-foreground mb-1">{card.label}</p>
                  <p className="text-2xl font-bold">{card.value}</p>
                  <p className="text-xs text-muted-foreground mt-1">{card.sub}</p>
                </div>
              ))}
            </div>

            {/* Revenue + Margin ComposedChart */}
            <div>
              <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Revenue &amp; Adj. EBITDA Margin Trend</p>
              <div className="h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={chartRevEbitda} margin={{ top: 10, right: 50, left: 20, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                    <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                    <YAxis yAxisId="left" tickFormatter={(v) => `$${v}M`} tick={{ fontSize: 10 }} width={50} />
                    <YAxis yAxisId="right" orientation="right" tickFormatter={(v) => `${v.toFixed(1)}%`} tick={{ fontSize: 10 }} width={45} domain={[0, 20]} />
                    <Tooltip formatter={(v: unknown, name: unknown) => [name === 'margin' ? `${typeof v === 'number' ? v.toFixed(1) : v}%` : `$${v}M`, name === 'margin' ? 'Adj. EBITDA Margin' : 'Revenue']} />
                    <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'margin' ? 'Adj. EBITDA Margin %' : 'Revenue ($M)'} />
                    <Bar yAxisId="left" dataKey="revenue" name="revenue" fill="#1E3A5F" />
                    <Line yAxisId="right" type="monotone" dataKey="margin" name="margin" stroke="#10b981" strokeWidth={2} dot={{ r: 4 }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Monthly Revenue + Customer Concentration */}
            <div className="grid grid-cols-2 gap-8">
              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Monthly Revenue &amp; EBITDA (TTM)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartMonthly} margin={{ top: 5, right: 10, left: 20, bottom: 5 }} barCategoryGap="25%" barGap={2}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="month" tick={{ fontSize: 10 }} />
                      <YAxis tickFormatter={(v) => `$${v}M`} tick={{ fontSize: 10 }} width={46} />
                      <Tooltip formatter={(v: unknown, name: unknown) => [`$${v}M`, name === 'ebitda' ? 'EBITDA' : 'Revenue']} />
                      <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'ebitda' ? 'EBITDA' : 'Revenue'} />
                      <Bar dataKey="revenue" name="revenue" fill="#1E3A5F" />
                      <Bar dataKey="ebitda" name="ebitda" fill="#10b981" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Customer Concentration (TTM)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart layout="vertical" data={chartCustomers} margin={{ top: 5, right: 50, left: 110, bottom: 5 }} barCategoryGap="20%">
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
                      <XAxis type="number" tickFormatter={(v) => `$${v}M`} tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" tick={{ fontSize: 9 }} width={105} />
                      <Tooltip formatter={(v: unknown) => [`$${v}M`]} />
                      <Bar dataKey="value" name="Revenue ($M)">
                        {chartCustomers.map((entry, idx) => (
                          <Cell key={idx} fill={entry.pct >= 15 ? '#ef4444' : entry.pct >= 10 ? '#f59e0b' : '#1E3A5F'} />
                        ))}
                        <LabelList dataKey="pct" position="right" formatter={(v: unknown) => typeof v === 'number' ? `${v.toFixed(1)}%` : ''} style={{ fontSize: 10, fill: '#666' }} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Cash Conversion + Margin Trend */}
            <div className="grid grid-cols-2 gap-8">
              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Cash Conversion — EBITDA vs. FCF</p>
                <div className="h-48">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartCashConv} margin={{ top: 5, right: 10, left: 20, bottom: 5 }} barCategoryGap="30%" barGap={4}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={(v) => `$${v.toFixed(1)}M`} tick={{ fontSize: 10 }} width={50} />
                      <Tooltip formatter={(v: unknown, name: unknown) => [`$${typeof v === 'number' ? v.toFixed(2) : v}M`, name === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow']} />
                      <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow'} />
                      <Bar dataKey="ebitda" name="ebitda" fill="#1E3A5F" />
                      <Bar dataKey="fcf" name="fcf" fill="#10b981" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div>
                <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Monthly Margin Trend (TTM)</p>
                <div className="h-48">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={chartMarginTrend} margin={{ top: 5, right: 20, left: 10, bottom: 16 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="month" tick={{ fontSize: 10 }} angle={-35} textAnchor="end" interval={1} />
                      <YAxis tickFormatter={(v: number) => `${v.toFixed(0)}%`} tick={{ fontSize: 10 }} width={38} />
                      <Tooltip formatter={(v: unknown) => [`${typeof v === 'number' ? v.toFixed(1) : v}%`]} />
                      <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'grossMargin' ? 'Gross Margin' : 'EBITDA Margin'} />
                      <Line type="monotone" dataKey="grossMargin" name="grossMargin" stroke="#1E3A5F" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="netMargin" name="netMargin" stroke="#10b981" strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Liquidity Metrics */}
            <div>
              <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Balance Sheet Ratios</p>
              <div className="grid grid-cols-4 gap-4">
                {[
                  { label: 'Current Ratio',  value: '2.7×', sub: 'Current Assets / Current Liabilities', pct: 85 },
                  { label: 'Quick Ratio',    value: '2.0×', sub: '(CA − Inventory) / CL', pct: 75 },
                  { label: 'Debt / Equity',  value: '0.31×', sub: 'Total Debt / Equity', pct: 31 },
                  { label: 'Return on Equity', value: '32.4%', sub: 'Net Income / Avg. Equity', pct: 92 },
                ].map((m) => (
                  <div key={m.label} className="bg-slate-50 rounded p-4">
                    <p className="text-xs text-muted-foreground mb-1">{m.label}</p>
                    <p className="text-2xl font-bold">{m.value}</p>
                    <p className="text-xs text-muted-foreground mt-1">{m.sub}</p>
                    <div className="mt-2 h-1 bg-gray-200 rounded-full overflow-hidden">
                      <div className="h-full bg-[#1E3A5F] rounded-full" style={{ width: `${m.pct}%` }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )
      }
      case 'complete-qoe':
        return (
          <CompleteQoeSection
            liveCells={liveCells}
            workbookId={id}
            isDemoWorkbook={isDemoWorkbook}
            periods={['FY20', 'FY21', 'FY22', 'TTM']}
            onViewSource={handleViewSource}
            onCellSave={fetchCells}
            onCellReference={(ctx) => {
              setCellRef(ctx)
              openChat()
            }}
          />
        )

      case 'risk-diligence':
        return (
          <RiskDiligenceSection
            liveCells={liveCells}
            workbookId={id}
            isDemoWorkbook={isDemoWorkbook}
            periods={['FY20', 'FY21', 'FY22', 'TTM']}
            onViewSource={handleViewSource}
          />
        )

      default:
        return (
          <div className="text-center py-12">
            <h2 className="text-2xl font-bold mb-4">{sections.find(s => s.id === activeSection)?.name}</h2>
            <p className="text-muted-foreground mb-6">This section is under development</p>
            <Badge variant="outline">Coming Soon</Badge>
          </div>
        )
    }
  }

  return (
    <div className="flex h-full">
      {/* Workbook Sidebar */}
      <div className="w-64 border-r bg-background flex flex-col overflow-y-auto">
        <div className="p-4 border-b">
          <h2 className="font-semibold text-lg">{workbookName}</h2>
          <p className="text-xs text-muted-foreground mt-1">Workbook Sections</p>
        </div>

        <div className="flex-1 p-3">
          {sections.map((section) => {
            const Icon = section.icon
            return (
              <div
                key={section.id}
                onClick={() => setActiveSection(section.id)}
                className={`flex items-center px-3 py-2 rounded-md cursor-pointer hover:bg-accent mb-1 ${
                  activeSection === section.id ? 'bg-accent' : ''
                }`}
              >
                <Icon className="mr-2 h-4 w-4" />
                <span className="text-sm">{section.name}</span>
              </div>
            )
          })}
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 overflow-y-auto">
        <div className="p-8">
          <div className="flex items-center justify-between mb-8">
            <div className="flex items-center gap-3">
              <h1 className="text-xl font-semibold">{workbookName}</h1>
              {workbookStatus && !isDemoWorkbook && (
                <Badge
                  className={
                    workbookStatus === 'ready'
                      ? 'bg-blue-600 hover:bg-blue-700'
                      : workbookStatus === 'needs_input'
                      ? 'bg-gray-500 hover:bg-gray-600'
                      : workbookStatus === 'analyzing'
                      ? 'bg-gray-400 hover:bg-gray-500'
                      : workbookStatus === 'error'
                      ? 'bg-gray-500 hover:bg-gray-600'
                      : 'bg-gray-500 hover:bg-gray-600'
                  }
                >
                  {workbookStatus === 'ready'
                    ? 'Ready'
                    : workbookStatus === 'needs_input'
                    ? 'Needs Input'
                    : workbookStatus === 'analyzing'
                    ? 'Analyzing...'
                    : workbookStatus === 'error'
                    ? 'Error'
                    : workbookStatus}
                </Badge>
              )}
              {isDemoWorkbook && (
                <Badge variant="outline">Demo</Badge>
              )}
            </div>
            <div className="flex flex-col items-end gap-1">
              <div className="flex items-center gap-2">
                <Button
                  variant="outline"
                  size="icon"
                  onClick={() => chatOpen ? closeChat() : openChat()}
                  className="relative"
                  title="Workbook AI"
                >
                  <Sparkles className="h-4 w-4" />
                  {openFlagCount > 0 && (
                    <span className="absolute -top-1 -right-1 bg-blue-500 text-white text-[9px] font-semibold rounded-full h-4 w-4 flex items-center justify-center leading-none">
                      {openFlagCount}
                    </span>
                  )}
                </Button>
                <Button variant="outline" size="icon" onClick={() => setSettingsOpen(true)}>
                  <Settings className="h-4 w-4" />
                </Button>
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button disabled={exporting || isDemoWorkbook}>
                      {exporting ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <Download className="mr-2 h-4 w-4" />
                      )}
                      Export
                      <ChevronDown className="ml-2 h-4 w-4" />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end">
                    <DropdownMenuItem onClick={() => handleExport('excel')}>
                      <Image src="/integrations/excel.png" alt="Excel" width={16} height={16} className="mr-2 object-contain" />
                      Excel (.xlsx)
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => handleExport('sheets')}>
                      <Image src="/integrations/google-sheets.png" alt="Google Sheets" width={16} height={16} className="mr-2 object-contain" />
                      Google Sheets
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => handleExport('airtable')}>
                      <Image src="/integrations/airtable.png" alt="Airtable" width={16} height={16} className="mr-2 object-contain" />
                      Airtable
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => handleExport('pdf')}>
                      <FileText className="mr-2 h-4 w-4 text-gray-500" />
                      PDF
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>
              </div>
              {exportError && (
                <p className="text-xs text-gray-500 max-w-xs text-right leading-tight">
                  {exportError}
                </p>
              )}
            </div>
          </div>

          {renderContent()}
        </div>
      </div>

      {/* Workbook Settings Dialog */}
      <WorkbookSettingsDialog
        open={settingsOpen}
        onOpenChange={setSettingsOpen}
        workbookName={workbookName}
        workbookId={isDemoWorkbook ? undefined : id}
        onAnalysisComplete={fetchCells}
        onDeleted={() => router.push('/dashboard')}
      />

      {/* Document Viewer Panel */}
      <DocumentViewerPanel
        open={viewerOpen}
        onOpenChange={setViewerOpen}
        sourceRef={viewerSource}
      />

      {/* AI Chat Panel is rendered at layout level — see app/dashboard/layout.tsx */}
    </div>
  )
}
