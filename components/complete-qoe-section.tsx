'use client'

import { useEffect, useRef, useState, memo } from 'react'
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  Cell,
  LabelList,
  ComposedChart,
  Line,
} from 'recharts'
import { TrendingUp, AlertTriangle, CheckCircle } from 'lucide-react'
import { AuditableCell, type SourceRef, type CellFlag } from '@/components/auditable-cell'

// ─── Types ────────────────────────────────────────────────────────────────────

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

interface Props {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  periods: string[]
  onViewSource: (sourceRef: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}

// ─── Formatters ───────────────────────────────────────────────────────────────

const fmt = (v: number) =>
  Math.abs(v) >= 1e6
    ? `$${(v / 1e6).toFixed(1)}M`
    : `$${(v / 1e3).toFixed(0)}K`

const fmtFull = (v: number) =>
  new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 0,
  }).format(v)

const fmtPct = (v: number) => `${v.toFixed(1)}%`

// ─── Live cell helpers ────────────────────────────────────────────────────────

function getLive(
  cells: LiveCell[],
  section: string,
  rowKey: string,
  period: string
): number | null {
  return (
    cells.find(
      (c) => c.section === section && c.row_key === rowKey && c.period === period
    )?.raw_value ?? null
  )
}

function getSrcRef(
  cells: LiveCell[],
  section: string,
  rowKey: string,
  period: string
): SourceRef | null {
  const cell = cells.find(
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
}

function getCellId(
  cells: LiveCell[],
  section: string,
  rowKey: string,
  period: string
): string | undefined {
  return cells.find(
    (c) => c.section === section && c.row_key === rowKey && c.period === period
  )?.id
}

function getCellFlags(
  cells: LiveCell[],
  section: string,
  rowKey: string,
  period: string
): CellFlag[] {
  return (
    cells.find(
      (c) => c.section === section && c.row_key === rowKey && c.period === period
    )?.flags ?? []
  )
}

// ─── Waterfall helpers ────────────────────────────────────────────────────────

interface BridgeItem {
  name: string
  value: number
  isBase?: boolean
  isEnd?: boolean
}

interface WaterfallPoint {
  name: string
  range: [number, number]
  rawValue: number
  isBase: boolean
  isEnd: boolean
}

function buildWaterfallPoints(items: BridgeItem[]): WaterfallPoint[] {
  const result: WaterfallPoint[] = []
  let running = 0
  items.forEach((item) => {
    if (item.isBase) {
      result.push({ name: item.name, range: [0, item.value], rawValue: item.value, isBase: true, isEnd: false })
      running = item.value
    } else if (item.isEnd) {
      result.push({ name: item.name, range: [0, running], rawValue: running, isBase: false, isEnd: true })
    } else {
      const lo = item.value >= 0 ? running : running + item.value
      const hi = item.value >= 0 ? running + item.value : running
      result.push({ name: item.name, range: [lo, hi], rawValue: item.value, isBase: false, isEnd: false })
      running += item.value
    }
  })
  return result
}

// ─── Demo Data ────────────────────────────────────────────────────────────────

const OVERVIEW_DEMO = {
  revenue:   [42187456, 51482903, 62745381, 68293742],
  adjEbitda: [4875291,  6524835,  7682947,  8547239],
  margin:    [11.6, 12.7, 12.2, 12.5],
}

const DEMO_BRIDGE_ITEMS: BridgeItem[] = [
  { name: 'EBITDA as Defined', value: 7960816, isBase: true },
  { name: 'a) Charitable', value: 315920 },
  { name: 'b) Owner Tax', value: 22450 },
  { name: 'c) Excess Comp', value: 248190 },
  { name: 'd) Personal Exp', value: 39863 },
  { name: 'e) Prof. Fees', value: 68200 },
  { name: 'f) Exec. Search', value: 52000 },
  { name: 'g) Relocation', value: 41500 },
  { name: 'h) Severance', value: 92300 },
  { name: 'i) Above-Mkt Rent', value: -168900 },
  { name: 'j) Owner Comp Norm', value: -425900 },
  { name: 'Adj. EBITDA', value: 0, isEnd: true },
]

const DEMO_NET_DEBT = [
  { rowKey: 'reported_debt', label: 'Reported Debt (Term Loan + Revolver)', value: 2850000 },
  { rowKey: 'capital_leases', label: 'Capital Leases', value: 420000 },
  { rowKey: 'unpaid_accrued_vacation', label: 'Unpaid Accrued Vacation', value: 185000 },
  { rowKey: 'aged_accounts_payable', label: 'Aged Accounts Payable (>90 days)', value: 94000 },
  { rowKey: 'deferred_revenue', label: 'Deferred Revenue', value: 318000 },
  { rowKey: 'customer_deposits', label: 'Customer Deposits', value: 142000 },
]
const DEMO_TOTAL_DEBT = DEMO_NET_DEBT.reduce((s, r) => s + r.value, 0)
const DEMO_CASH = 5310000
const DEMO_NET_DEBT_FINAL = DEMO_TOTAL_DEBT - DEMO_CASH

const DEMO_CUSTOMERS = [
  { name: 'Acme Manufacturing', value: 14205600, pct: 20.8 },
  { name: 'Nexgen Solutions', value: 9561300, pct: 14.0 },
  { name: 'Bridgewater Corp', value: 7512100, pct: 11.0 },
  { name: 'Summit Enterprises', value: 5464700, pct: 8.0 },
  { name: 'Crestwood Partners', value: 4097600, pct: 6.0 },
  { name: 'Lakeside Industries', value: 3414700, pct: 5.0 },
  { name: 'Horizon Group', value: 2731700, pct: 4.0 },
  { name: 'Pacific Dynamics', value: 2048900, pct: 3.0 },
  { name: 'Allied Services', value: 1366000, pct: 2.0 },
  { name: 'Meridian LLC', value: 683000, pct: 1.0 },
  { name: 'All Other', value: 17207400, pct: 25.2 },
]
const DEMO_TOTAL_REV = DEMO_CUSTOMERS.reduce((s, c) => s + c.value, 0)

const DEMO_CASH_CONV = {
  FY22: { ebitda: 7682947, capex: 384147, wc: 192075, taxes: 307317, interest: 61245, fcf: 6738163 },
  TTM:  { ebitda: 8547239, capex: 427361, wc: 256417, taxes: 341889, interest: 65182, fcf: 7456390 },
}

const DEMO_PROOF = [
  { period: 'FY20', gl: 42187456, tax: 42063812, bank: 42075200 },
  { period: 'FY21', gl: 51482903, tax: 51328450, bank: 51349800 },
  { period: 'FY22', gl: 62745381, tax: 62558200, bank: 62580900 },
  { period: 'TTM',  gl: 68293742, tax: 68088300, bank: 68112500 },
]

const WC_DEMO = [
  { rowKey: 'wc_accounts_receivable',  label: 'Accounts Receivable',     current: 1886360,  normalized: 1823947,  adj: -62413,  days: 10.1, targetDays: 9.8  },
  { rowKey: 'wc_inventory',            label: 'Inventory',                current: 1568739,  normalized: 1687429,  adj: 118690,  days: 19.4, targetDays: 20.9 },
  { rowKey: 'wc_prepaid_expenses',     label: 'Prepaid Expenses',         current: 284625,   normalized: 284625,   adj: 0,       days: 0,    targetDays: 0    },
  { rowKey: 'wc_other_current_assets', label: 'Other Current Assets',     current: 124850,   normalized: 124850,   adj: 0,       days: 0,    targetDays: 0    },
  { rowKey: 'wc_accounts_payable',     label: '(Less) Accounts Payable',  current: -1481205, normalized: -1481205, adj: 0,       days: 18.3, targetDays: 18.3 },
  { rowKey: 'wc_accrued_expenses',     label: '(Less) Accrued Expenses',  current: -642180,  normalized: -642180,  adj: 0,       days: 0,    targetDays: 0    },
  { rowKey: 'wc_credit_cards_payable', label: '(Less) Credit Cards',      current: -98420,   normalized: -98420,   adj: 0,       days: 0,    targetDays: 0    },
  { rowKey: 'wc_sales_tax_payable',    label: '(Less) Sales Tax Payable', current: -38240,   normalized: -38240,   adj: 0,       days: 0,    targetDays: 0    },
  { rowKey: 'wc_deferred_revenue',     label: '(Less) Deferred Revenue',  current: -49600,   normalized: -49600,   adj: 0,       days: 0,    targetDays: 0    },
]

const MONTHLY_DEMO = [
  { month: 'Feb', revenue: 5284750, cogs: 2272443, opex: 2431065 },
  { month: 'Mar', revenue: 5692410, cogs: 2447656, opex: 2619509 },
  { month: 'Apr', revenue: 5438920, cogs: 2338735, opex: 2501903 },
  { month: 'May', revenue: 5847630, cogs: 2514481, opex: 2690470 },
  { month: 'Jun', revenue: 5534180, cogs: 2379695, opex: 2545762 },
  { month: 'Jul', revenue: 5621490, cogs: 2417240, opex: 2585885 },
  { month: 'Aug', revenue: 5748320, cogs: 2471778, opex: 2644227 },
  { month: 'Sep', revenue: 5493870, cogs: 2362364, opex: 2527180 },
  { month: 'Oct', revenue: 5962840, cogs: 2564021, opex: 2742906 },
  { month: 'Nov', revenue: 5384620, cogs: 2315387, opex: 2476925 },
  { month: 'Dec', revenue: 5658820, cogs: 2433293, opex: 2603057 },
  { month: 'Jan', revenue: 5662136, cogs: 2434613, opex: 2605351 },
]

const DEMO_RUN_RATE: BridgeItem[] = [
  { name: 'TTM Revenue', value: 68293742, isBase: true },
  { name: 'New Contracts', value: 1850000 },
  { name: 'Full-Year Hire', value: 420000 },
  { name: 'Price Increases', value: 680000 },
  { name: 'Lost Revenue', value: -944000 },
  { name: 'Pro Forma Rev', value: 0, isEnd: true },
]
const DEMO_PRO_FORMA_REV = 70299742
const DEMO_PRO_FORMA_EBITDA = 8843000

const ADJ_ROWS = [
  { rowKey: 'adj_charitable_contributions',    label: 'a) Charitable Contributions',     fy20: 15000,     fy21: 248500,  fy22: 287350,  ttm: 315920  },
  { rowKey: 'adj_owner_tax_legal',             label: 'b) Owner Tax & Legal',             fy20: 18500,     fy21: 18500,   fy22: 21200,   ttm: 22450   },
  { rowKey: 'adj_excess_owner_comp',           label: 'c) Excess Owner Compensation',     fy20: 185200,    fy21: 208750,  fy22: 236280,  ttm: 248190  },
  { rowKey: 'adj_personal_expenses',           label: 'd) Personal Expenses',             fy20: 26574,     fy21: 31850,   fy22: 35894,   ttm: 39863   },
  { rowKey: 'adj_professional_fees_onetime',   label: 'e) Professional Fees - One-time',  fy20: 0,         fy21: 52400,   fy22: 58900,   ttm: 68200   },
  { rowKey: 'adj_executive_search',            label: 'f) Executive Search Fees',         fy20: 0,         fy21: 65000,   fy22: 48000,   ttm: 52000   },
  { rowKey: 'adj_facility_relocation',         label: 'g) Facility Relocation',           fy20: 0,         fy21: 0,       fy22: 38750,   ttm: 41500   },
  { rowKey: 'adj_severance',                   label: 'h) Severance',                     fy20: 0,         fy21: 0,       fy22: 85400,   ttm: 92300   },
  { rowKey: 'adj_above_market_rent',           label: 'i) Above-Market Rent',             fy20: 0,         fy21: -158714, fy22: -162500, ttm: -168900 },
  { rowKey: 'adj_normalize_owner_comp',        label: 'j) Normalize Owner Comp',          fy20: -244500,   fy21: -382300, fy22: -404000, ttm: -425900 },
]

// ─── Design system atoms ──────────────────────────────────────────────────────

function MajorDivider() {
  return <div className="border-t-2 border-[#1E3A5F]/20 my-16" />
}

function Divider() {
  return <div className="border-t border-gray-100 my-10" />
}

function SectionHeader({ children }: { children: React.ReactNode }) {
  return (
    <h2 className="text-lg font-semibold border-l-4 border-[#1E3A5F] pl-4 mb-3">
      {children}
    </h2>
  )
}

function SectionNumber({ children }: { children: React.ReactNode }) {
  return (
    <p className="text-xs font-bold uppercase tracking-widest text-[#1E3A5F] mb-1">
      {children}
    </p>
  )
}

function Narrative({ children }: { children: React.ReactNode }) {
  return (
    <p className="text-sm text-muted-foreground leading-relaxed mb-6">
      {children}
    </p>
  )
}

function StatCard({ label, value, sub }: { label: string; value: string; sub?: string }) {
  return (
    <div className="bg-slate-50 rounded p-4">
      <p className="text-xs text-muted-foreground mb-1">{label}</p>
      <p className="text-2xl font-bold text-foreground">{value}</p>
      {sub && <p className="text-xs text-muted-foreground mt-1">{sub}</p>}
    </div>
  )
}

// ─── Tooltip components ───────────────────────────────────────────────────────

interface WaterfallTooltipProps {
  active?: boolean
  payload?: Array<{ payload: WaterfallPoint }>
  label?: string
}

function WaterfallTooltip({ active, payload, label }: WaterfallTooltipProps) {
  if (!active || !payload?.length) return null
  const d = payload[0].payload
  return (
    <div className="bg-white border border-gray-200 rounded px-3 py-2 text-xs shadow">
      <div className="font-semibold mb-1">{label}</div>
      <div>
        {d.isBase || d.isEnd
          ? fmtFull(d.range[1])
          : `${d.rawValue >= 0 ? '+' : ''}${fmtFull(d.rawValue)}`}
      </div>
    </div>
  )
}

interface ConcentrationTooltipProps {
  active?: boolean
  payload?: Array<{ value: number; payload: typeof DEMO_CUSTOMERS[number] }>
  label?: string
}

function ConcentrationTooltip({ active, payload, label }: ConcentrationTooltipProps) {
  if (!active || !payload?.length) return null
  const d = payload[0]
  return (
    <div className="bg-white border border-gray-200 rounded px-3 py-2 text-xs shadow">
      <div className="font-semibold mb-1">{label}</div>
      <div>{fmtFull(d.value)}</div>
      <div className="text-muted-foreground">{fmtPct(d.payload.pct)} of total</div>
    </div>
  )
}

function getConcentrationColor(pct: number) {
  if (pct >= 15) return '#ef4444'
  if (pct >= 10) return '#f59e0b'
  return '#1E3A5F'
}

// ─── ToC ──────────────────────────────────────────────────────────────────────

const TOC_ITEMS = [
  { id: 'sec-exec',     label: 'I. Executive Summary' },
  { id: 'sec-scope',    label: 'II. Scope & Procedures' },
  { id: 'sec-qoe',      label: 'III. Quality of Earnings' },
  { id: 'sec-revenue',  label: 'IV. Quality of Revenue' },
  { id: 'sec-assets',   label: 'V. Quality of Net Assets' },
  { id: 'sec-cashflow', label: 'VI. Cash Flow & CapEx' },
  { id: 'sec-appendix', label: 'VII. Appendices' },
]

// ─── Section I: Executive Summary ────────────────────────────────────────────

function ExecSummarySection({
  liveCells,
}: {
  liveCells: LiveCell[]
}) {
  const periods = ['FY20', 'FY21', 'FY22', 'TTM']
  const ttmRev = getLive(liveCells, 'overview', 'total_revenue', 'TTM') ?? OVERVIEW_DEMO.revenue[3]
  const ttmEbitda = getLive(liveCells, 'qoe', 'diligence_adjusted_ebitda', 'TTM') ?? OVERVIEW_DEMO.adjEbitda[3]
  const ttmMargin = ttmRev > 0 ? (ttmEbitda / ttmRev) * 100 : OVERVIEW_DEMO.margin[3]
  const netDebt = getLive(liveCells, 'net-debt', 'net_debt', 'TTM') ?? DEMO_NET_DEBT_FINAL
  const isNetCash = netDebt < 0

  // Revenue CAGR FY20→FY22 (3-year)
  const rev20 = getLive(liveCells, 'overview', 'total_revenue', 'FY20') ?? OVERVIEW_DEMO.revenue[0]
  const rev22 = getLive(liveCells, 'overview', 'total_revenue', 'FY22') ?? OVERVIEW_DEMO.revenue[2]
  const cagr = rev20 > 0 ? (Math.pow(rev22 / rev20, 1 / 2) - 1) * 100 : 22.0

  const chartData = periods.map((p, i) => {
    const rev = getLive(liveCells, 'overview', 'total_revenue', p) ?? OVERVIEW_DEMO.revenue[i]
    const ebitda = getLive(liveCells, 'qoe', 'diligence_adjusted_ebitda', p) ?? OVERVIEW_DEMO.adjEbitda[i]
    const margin = rev > 0 ? (ebitda / rev) * 100 : OVERVIEW_DEMO.margin[i]
    return { period: p, revenue: rev / 1e6, margin }
  })

  const topCustPct = DEMO_CUSTOMERS[0].pct
  const maxProofVar = Math.max(...DEMO_PROOF.map((p) => Math.abs(((p.gl - p.tax) / p.gl) * 100)))

  return (
    <div>
      <SectionNumber>Section I</SectionNumber>
      <SectionHeader>Executive Summary</SectionHeader>
      <Narrative>
        This Quality of Earnings report covers the operations of the subject company for the periods
        FY2020 through trailing twelve months (TTM). The analysis is prepared for internal diligence
        use and presents normalized earnings, quality of revenue, net asset analysis, and cash flow
        conversion. All figures in USD unless otherwise noted.
      </Narrative>

      {/* Deal Highlight Stats */}
      <div className="grid grid-cols-4 gap-4 mb-8">
        <StatCard label="TTM Revenue" value={fmt(ttmRev)} sub="Trailing 12 Months" />
        <StatCard label="TTM Adj. EBITDA" value={fmt(ttmEbitda)} sub="Diligence-normalized" />
        <StatCard label="EBITDA Margin" value={fmtPct(ttmMargin)} sub="TTM normalized" />
        <StatCard
          label={isNetCash ? 'Net Cash Position' : 'Net Debt'}
          value={fmt(Math.abs(netDebt))}
          sub={isNetCash ? 'Net cash at close' : 'Debt at close'}
        />
      </div>

      {/* Revenue + Margin Chart */}
      <div className="h-56 mb-8">
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={chartData} margin={{ top: 10, right: 40, left: 20, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis yAxisId="left" tickFormatter={(v) => `$${v.toFixed(0)}M`} tick={{ fontSize: 10 }} width={50} />
            <YAxis yAxisId="right" orientation="right" tickFormatter={(v) => `${v.toFixed(1)}%`} tick={{ fontSize: 10 }} width={45} domain={[0, 20]} />
            <Tooltip
              formatter={(v: unknown, name: unknown) => [
                name === 'margin'
                  ? `${typeof v === 'number' ? v.toFixed(1) : v}%`
                  : `$${typeof v === 'number' ? v.toFixed(2) : v}M`,
                name === 'margin' ? 'Adj. EBITDA Margin' : 'Revenue',
              ]}
            />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v) => v === 'margin' ? 'Adj. EBITDA Margin %' : 'Revenue ($M)'} />
            <Bar yAxisId="left" dataKey="revenue" name="revenue" fill="#1E3A5F" />
            <Line yAxisId="right" type="monotone" dataKey="margin" name="margin" stroke="#10b981" strokeWidth={2} dot={{ r: 4 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      {/* Key Findings */}
      <div className="space-y-2">
        <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Key Findings</p>
        <div className="flex items-start gap-2 text-sm">
          <TrendingUp className="h-4 w-4 text-emerald-500 mt-0.5 shrink-0" />
          <span><em>Revenue CAGR (FY20–FY22): <strong>{fmtPct(cagr)}</strong> — consistent organic growth with no M&A contribution.</em></span>
        </div>
        <div className="flex items-start gap-2 text-sm">
          <CheckCircle className="h-4 w-4 text-emerald-500 mt-0.5 shrink-0" />
          <span><em>Three-way revenue match variance: <strong>{fmtPct(maxProofVar)}</strong> — all periods within acceptable threshold of 0.5%.</em></span>
        </div>
        <div className="flex items-start gap-2 text-sm">
          <AlertTriangle className="h-4 w-4 text-amber-500 mt-0.5 shrink-0" />
          <span><em>Top customer concentration: <strong>{fmtPct(topCustPct)}</strong> of TTM revenue — recommend confirming contract renewals prior to close.</em></span>
        </div>
        <div className="flex items-start gap-2 text-sm">
          <CheckCircle className="h-4 w-4 text-emerald-500 mt-0.5 shrink-0" />
          <span><em>Net {isNetCash ? 'cash' : 'debt'} position of <strong>{fmt(Math.abs(netDebt))}</strong> — {isNetCash ? 'partial offset to purchase price' : 'buyer to assume at close'}.</em></span>
        </div>
      </div>
    </div>
  )
}

// ─── Section II: Scope & Procedures ──────────────────────────────────────────

function ScopeSection() {
  return (
    <div>
      <SectionNumber>Section II</SectionNumber>
      <SectionHeader>Scope &amp; Procedures</SectionHeader>
      <Narrative>
        This engagement was conducted in accordance with Quality of Earnings diligence standards.
        The procedures performed were agreed upon between the buyer and their advisors and do not
        constitute an audit, review, or compilation. Accordingly, we do not express an opinion or
        any other form of assurance on the financial statements or the information contained herein.
      </Narrative>

      <table className="w-full text-sm mb-6">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold w-1/4">Source Type</th>
            <th className="text-left py-2 font-semibold w-1/2">Description</th>
            <th className="text-left py-2 font-semibold w-1/4">Periods Provided</th>
          </tr>
        </thead>
        <tbody>
          {[
            { type: 'General Ledger', desc: 'Trial balance and detailed GL exports by account, monthly', periods: 'FY2020 – TTM' },
            { type: 'Tax Returns', desc: 'Federal Form 1120 and state income tax returns', periods: 'FY2020 – FY2022' },
            { type: 'Bank Statements', desc: 'Primary operating and payroll account statements', periods: 'FY2020 – TTM' },
            { type: 'Payroll Records', desc: 'ADP payroll registers, W-2 summary by employee', periods: 'FY2021 – TTM' },
            { type: 'Customer Data', desc: 'CRM export: revenue by customer, contract dates, renewal status', periods: 'TTM' },
            { type: 'Fixed Assets', desc: 'Fixed asset register, depreciation schedules, capex invoices', periods: 'FY2020 – TTM' },
          ].map((row) => (
            <tr key={row.type} className="border-b border-gray-100">
              <td className="py-1.5 font-medium">{row.type}</td>
              <td className="py-1.5 text-muted-foreground">{row.desc}</td>
              <td className="py-1.5 text-muted-foreground">{row.periods}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <Narrative>
        Our testing coverage included detailed transaction-level testing for items exceeding $10,000.
        We tested approximately 78% of total disbursements and 100% of revenue by source across all
        periods. No reliance was placed on management representations without independent corroboration.
      </Narrative>
    </div>
  )
}

// ─── Section III: Quality of Earnings ────────────────────────────────────────

function QoeSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}) {
  const periods = ['FY20', 'FY21', 'FY22', 'TTM']
  const periodKeys = ['fy20', 'fy21', 'fy22', 'ttm'] as const

  const liveEbitda = getLive(liveCells, 'qoe', 'ebitda_as_defined', 'TTM')
  const liveAdjEbitda = getLive(liveCells, 'qoe', 'diligence_adjusted_ebitda', 'TTM')
  const ttmEbitda = liveEbitda ?? 7960816
  const ttmAdjEbitda = liveAdjEbitda ?? 8547239
  const totalAdj = ttmAdjEbitda - ttmEbitda

  const waterfallData = buildWaterfallPoints(DEMO_BRIDGE_ITEMS)
  const wfYMin = ttmEbitda * 0.97
  const wfYMax = Math.max(...waterfallData.map((p) => p.range[1])) * 1.02
  const getBarColor = (entry: WaterfallPoint) => {
    if (entry.isBase || entry.isEnd) return '#1E3A5F'
    return entry.rawValue >= 0 ? '#10b981' : '#ef4444'
  }

  // Adj Schedule per-period values
  const ebitdaAsDefinedByPeriod = [4630517, 6177949, 7217823, 7960816]

  const totalAdjByPeriod = periods.map((p, i) => {
    return ADJ_ROWS.reduce((sum, row) => {
      const live = getLive(liveCells, 'qoe', row.rowKey, p)
      return sum + (live ?? row[periodKeys[i]])
    }, 0)
  })

  const adjEbitdaByPeriod = periods.map((p, i) => {
    const live = getLive(liveCells, 'qoe', 'diligence_adjusted_ebitda', p)
    return live ?? OVERVIEW_DEMO.adjEbitda[i]
  })

  // Margin analysis chart
  const marginChartData = periods.map((p, i) => {
    const unadj = ebitdaAsDefinedByPeriod[i]
    const adj = adjEbitdaByPeriod[i]
    const rev = getLive(liveCells, 'overview', 'total_revenue', p) ?? OVERVIEW_DEMO.revenue[i]
    return {
      period: p,
      unadjMargin: rev > 0 ? (unadj / rev) * 100 : 0,
      adjMargin: rev > 0 ? (adj / rev) * 100 : 0,
      unadjEbitda: unadj / 1e6,
      adjEbitda: adj / 1e6,
    }
  })

  return (
    <div>
      <SectionNumber>Section III</SectionNumber>
      <SectionHeader>Quality of Earnings</SectionHeader>
      <Narrative>
        TTM Revenue of {fmt(getLive(liveCells, 'overview', 'total_revenue', 'TTM') ?? OVERVIEW_DEMO.revenue[3])} and
        Diligence-Adjusted EBITDA of {fmt(ttmAdjEbitda)} ({fmtPct(ttmAdjEbitda / (getLive(liveCells, 'overview', 'total_revenue', 'TTM') ?? OVERVIEW_DEMO.revenue[3]) * 100)} margin)
        reflect 10 discrete adjustments totaling {fmt(totalAdj)} net. All add-backs are supported
        by GL-level documentation and independently verified against source records.
      </Narrative>

      {/* EBITDA Waterfall */}
      <div className="h-72 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={waterfallData} margin={{ top: 10, right: 20, left: 30, bottom: 40 }} barCategoryGap="20%">
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="name" tick={{ fontSize: 10 }} angle={-35} textAnchor="end" interval={0} />
            <YAxis domain={[wfYMin, wfYMax]} tickFormatter={fmt} tick={{ fontSize: 10 }} width={60} />
            <Tooltip content={<WaterfallTooltip />} />
            <Bar dataKey="range">
              {waterfallData.map((entry, idx) => (
                <Cell key={idx} fill={getBarColor(entry)} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <Divider />

      {/* Full Adjustment Schedule */}
      <SectionHeader>Full Adjustment Schedule — All Periods</SectionHeader>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-200">
              <th className="text-left py-2 font-semibold">Adjustment</th>
              {periods.map((p) => (
                <th key={p} className="text-right py-2 font-semibold px-2">{p}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr className="border-b border-gray-100 font-semibold">
              <td className="py-1.5">EBITDA, as Defined</td>
              {periods.map((p, i) => (
                <td key={p} className="text-right py-1.5 px-2">
                  <AuditableCell
                    value={fmtFull(getLive(liveCells, 'qoe', 'ebitda_as_defined', p) ?? ebitdaAsDefinedByPeriod[i])}
                    source="General Ledger — calculated per QoE methodology"
                    sourceRef={getSrcRef(liveCells, 'qoe', 'ebitda_as_defined', p)}
                    cellId={getCellId(liveCells, 'qoe', 'ebitda_as_defined', p)}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    flags={getCellFlags(liveCells, 'qoe', 'ebitda_as_defined', p)}
                    onViewSource={onViewSource}
                    onSave={onCellSave}
                    onReference={onCellReference}
                    period={p}
                  />
                </td>
              ))}
            </tr>
            {ADJ_ROWS.map((row) => (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className="py-1.5 text-muted-foreground pl-4">{row.label}</td>
                {periods.map((p, i) => {
                  const live = getLive(liveCells, 'qoe', row.rowKey, p)
                  const val = live ?? row[periodKeys[i]]
                  return (
                    <td key={p} className={`text-right py-1.5 px-2 ${val < 0 ? 'text-red-600' : val > 0 ? 'text-emerald-600' : 'text-muted-foreground'}`}>
                      <AuditableCell
                        value={fmtFull(val)}
                        source={`GL — diligence adjustment ${row.label}`}
                        sourceRef={getSrcRef(liveCells, 'qoe', row.rowKey, p)}
                        cellId={getCellId(liveCells, 'qoe', row.rowKey, p)}
                        workbookId={isDemoWorkbook ? undefined : workbookId}
                        flags={getCellFlags(liveCells, 'qoe', row.rowKey, p)}
                        onViewSource={onViewSource}
                        onSave={onCellSave}
                        onReference={onCellReference}
                        period={p}
                      />
                    </td>
                  )
                })}
              </tr>
            ))}
            <tr className="border-b border-gray-200 font-semibold text-muted-foreground">
              <td className="py-1.5 pl-4">Total Net Adjustments</td>
              {totalAdjByPeriod.map((total, i) => (
                <td key={i} className="text-right py-1.5 px-2">{fmt(total)}</td>
              ))}
            </tr>
            <tr className="font-bold">
              <td className="py-2">Diligence-Adjusted EBITDA</td>
              {periods.map((p, i) => (
                <td key={p} className="text-right py-2 px-2">
                  <AuditableCell
                    value={fmtFull(adjEbitdaByPeriod[i])}
                    source="Final normalized run-rate EBITDA — buyer perspective"
                    sourceRef={getSrcRef(liveCells, 'qoe', 'diligence_adjusted_ebitda', p)}
                    cellId={getCellId(liveCells, 'qoe', 'diligence_adjusted_ebitda', p)}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    flags={getCellFlags(liveCells, 'qoe', 'diligence_adjusted_ebitda', p)}
                    onViewSource={onViewSource}
                    onSave={onCellSave}
                    onReference={onCellReference}
                    period={p}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>

      <Divider />

      {/* Margin Analysis */}
      <SectionHeader>Margin Analysis — Unadjusted vs. Adjusted EBITDA</SectionHeader>
      <div className="h-56 mb-4">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={marginChartData} margin={{ top: 10, right: 20, left: 20, bottom: 5 }} barCategoryGap="30%" barGap={4}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v.toFixed(1)}M`} tick={{ fontSize: 10 }} width={50} />
            <Tooltip formatter={(v: unknown, name: unknown) => [`$${typeof v === 'number' ? v.toFixed(2) : v}M`, name === 'unadjEbitda' ? 'Unadj. EBITDA' : 'Adj. EBITDA']} />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v) => v === 'unadjEbitda' ? 'Unadj. EBITDA' : 'Adj. EBITDA'} />
            <Bar dataKey="unadjEbitda" name="unadjEbitda" fill="#94a3b8" />
            <Bar dataKey="adjEbitda" name="adjEbitda" fill="#1E3A5F" />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <Narrative>
        Adj. EBITDA margins of {fmtPct(marginChartData[3].adjMargin)} (TTM) vs. {fmtPct(marginChartData[3].unadjMargin)} unadjusted — representing
        {' '}{fmtPct(marginChartData[3].adjMargin - marginChartData[3].unadjMargin)} of normalization. Margins have been broadly stable across all four periods,
        indicating consistent operating leverage and no one-time uplift in the reported TTM.
      </Narrative>

      <Divider />

      {/* Run-Rate & Pro-Forma */}
      <RunRateSubSection liveCells={liveCells} />
    </div>
  )
}

// ─── Section III Sub: Run-Rate & Pro-Forma ────────────────────────────────────

function RunRateSubSection({ liveCells }: { liveCells: LiveCell[] }) {
  const rrWaterfallData = buildWaterfallPoints(DEMO_RUN_RATE)
  const rrYMin = 68293742 * 0.97
  const rrYMax = Math.max(...rrWaterfallData.map((p) => p.range[1])) * 1.02
  const getBarColor = (entry: WaterfallPoint) => {
    if (entry.isBase || entry.isEnd) return '#1E3A5F'
    return entry.rawValue >= 0 ? '#10b981' : '#ef4444'
  }

  const proFormaRev = getLive(liveCells, 'run-rate', 'pro_forma_revenue', 'TTM') ?? DEMO_PRO_FORMA_REV
  const proFormaEbitda = getLive(liveCells, 'run-rate', 'pro_forma_ebitda', 'TTM') ?? DEMO_PRO_FORMA_EBITDA
  const proFormaMargin = proFormaRev > 0 ? (proFormaEbitda / proFormaRev) * 100 : 0

  return (
    <div>
      <SectionHeader>Run-Rate &amp; Pro-Forma Revenue</SectionHeader>
      <Narrative>
        Pro-forma revenue reflects the annualized run-rate adjusted for known contract wins,
        full-year headcount effects, executed price increases, and identified churn. No
        incremental margin expansion is assumed; pro-forma EBITDA margin is held at TTM rates.{' '}
        <TrendingUp className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>
          Pro-forma revenue of {fmt(proFormaRev)} represents{' '}
          {fmtPct(((proFormaRev - 68293742) / 68293742) * 100)} uplift above TTM.
        </em>
      </Narrative>

      <div className="h-64 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={rrWaterfallData} margin={{ top: 10, right: 20, left: 30, bottom: 40 }} barCategoryGap="20%">
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="name" tick={{ fontSize: 10 }} angle={-35} textAnchor="end" interval={0} />
            <YAxis domain={[rrYMin, rrYMax]} tickFormatter={fmt} tick={{ fontSize: 10 }} width={60} />
            <Tooltip content={<WaterfallTooltip />} />
            <Bar dataKey="range">
              {rrWaterfallData.map((entry, idx) => (
                <Cell key={idx} fill={getBarColor(entry)} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Metric</th>
            <th className="text-right py-2 font-semibold">TTM Actual</th>
            <th className="text-right py-2 font-semibold">Pro Forma</th>
            <th className="text-right py-2 font-semibold">Change</th>
          </tr>
        </thead>
        <tbody>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground">Revenue</td>
            <td className="text-right py-1.5">{fmt(68293742)}</td>
            <td className="text-right py-1.5">{fmt(proFormaRev)}</td>
            <td className="text-right py-1.5 text-emerald-600">+{fmtPct(((proFormaRev - 68293742) / 68293742) * 100)}</td>
          </tr>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground">Adj. EBITDA</td>
            <td className="text-right py-1.5">{fmt(8547239)}</td>
            <td className="text-right py-1.5">{fmt(proFormaEbitda)}</td>
            <td className="text-right py-1.5 text-emerald-600">+{fmtPct(((proFormaEbitda - 8547239) / 8547239) * 100)}</td>
          </tr>
          <tr>
            <td className="py-2 text-muted-foreground">EBITDA Margin</td>
            <td className="text-right py-2">{fmtPct(12.5)}</td>
            <td className="text-right py-2">{fmtPct(proFormaMargin)}</td>
            <td className="text-right py-2 text-muted-foreground">
              {proFormaMargin >= 12.5 ? '+' : ''}{fmtPct(proFormaMargin - 12.5)}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── Section IV: Quality of Revenue ──────────────────────────────────────────

function RevenueSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}) {
  const maxVar = Math.max(...DEMO_PROOF.map((p) => Math.abs(((p.gl - p.tax) / p.gl) * 100)))
  const topCustomer = DEMO_CUSTOMERS[0]
  const top3Pct = DEMO_CUSTOMERS.slice(0, 3).reduce((s, c) => s + c.pct, 0)

  const proofChartData = DEMO_PROOF.map((p) => {
    const liveGl = getLive(liveCells, 'proof-revenue', 'gl_revenue', p.period)
    const liveTax = getLive(liveCells, 'proof-revenue', 'tax_return_revenue', p.period)
    const liveBank = getLive(liveCells, 'proof-revenue', 'bank_deposit_revenue', p.period)
    return {
      period: p.period,
      gl: (liveGl ?? p.gl) / 1e6,
      tax: (liveTax ?? p.tax) / 1e6,
      bank: (liveBank ?? p.bank) / 1e6,
    }
  })

  return (
    <div>
      <SectionNumber>Section IV</SectionNumber>
      <SectionHeader>Quality of Revenue</SectionHeader>

      {/* Proof of Revenue */}
      <Narrative>
        Revenue is independently corroborated across three sources: the general ledger, federal
        tax returns (Form 1120), and bank deposit summaries. Variances exceeding 0.5% are flagged.{' '}
        <CheckCircle className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>All three sources agree within {fmtPct(maxVar)} — no material revenue recognition inconsistencies identified.</em>
      </Narrative>

      <div className="h-56 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={proofChartData} margin={{ top: 10, right: 20, left: 30, bottom: 5 }} barCategoryGap="25%" barGap={2}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v.toFixed(0)}M`} tick={{ fontSize: 10 }} width={50} />
            <Tooltip formatter={(v: unknown) => [`$${typeof v === 'number' ? v.toFixed(2) : v}M`]} />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
            <Bar dataKey="gl" name="GL Revenue" fill="#1E3A5F" />
            <Bar dataKey="tax" name="Tax Return" fill="#3b82f6" />
            <Bar dataKey="bank" name="Bank Deposits" fill="#10b981" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm mb-4">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Source</th>
            {DEMO_PROOF.map((p) => (
              <th key={p.period} className="text-right py-2 font-semibold">{p.period}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {[
            { rowKey: 'gl_revenue', label: 'GL Revenue', demoVals: DEMO_PROOF.map((p) => p.gl) },
            { rowKey: 'tax_return_revenue', label: 'Tax Return Revenue', demoVals: DEMO_PROOF.map((p) => p.tax) },
            { rowKey: 'bank_deposit_revenue', label: 'Bank Deposit Revenue', demoVals: DEMO_PROOF.map((p) => p.bank) },
          ].map((row) => (
            <tr key={row.rowKey} className="border-b border-gray-100">
              <td className="py-1.5 text-muted-foreground pl-4">{row.label}</td>
              {DEMO_PROOF.map((p, idx) => {
                const live = getLive(liveCells, 'proof-revenue', row.rowKey, p.period)
                const val = live ?? row.demoVals[idx]
                return (
                  <td key={p.period} className="text-right py-1.5">
                    <AuditableCell
                      value={fmtFull(val)}
                      source={`${row.label} — ${p.period}`}
                      sourceRef={getSrcRef(liveCells, 'proof-revenue', row.rowKey, p.period)}
                      cellId={getCellId(liveCells, 'proof-revenue', row.rowKey, p.period)}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                      onSave={onCellSave}
                      onReference={onCellReference}
                      period={p.period}
                    />
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>

      <Divider />

      {/* Customer Concentration */}
      <SectionHeader>Customer Concentration</SectionHeader>
      <Narrative>
        Revenue concentration analysis covers TTM revenue of {fmt(DEMO_TOTAL_REV)} across{' '}
        {DEMO_CUSTOMERS.length - 1} named customers plus a long tail.{' '}
        <AlertTriangle className="inline h-3.5 w-3.5 text-amber-500 ml-0.5 mr-0.5" />
        <em>
          Top customer ({topCustomer.name}) represents {fmtPct(topCustomer.pct)} of TTM revenue.
          Top 3 customers represent {fmtPct(top3Pct)}. Recommend confirming retention rates and
          contract renewal dates prior to close.
        </em>
      </Narrative>

      <div className="h-72 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            layout="vertical"
            data={DEMO_CUSTOMERS.filter((c) => c.name !== 'All Other')}
            margin={{ top: 5, right: 80, left: 120, bottom: 5 }}
            barCategoryGap="20%"
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
            <XAxis type="number" tickFormatter={(v) => fmt(v)} tick={{ fontSize: 10 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={115} />
            <Tooltip content={<ConcentrationTooltip />} />
            <Bar dataKey="value" name="Revenue">
              {DEMO_CUSTOMERS.filter((c) => c.name !== 'All Other').map((entry, idx) => (
                <Cell key={idx} fill={getConcentrationColor(entry.pct)} />
              ))}
              <LabelList
                dataKey="pct"
                position="right"
                formatter={(v: unknown) => typeof v === 'number' ? `${v.toFixed(1)}%` : ''}
                style={{ fontSize: 10, fill: '#666' }}
              />
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Customer</th>
            <th className="text-right py-2 font-semibold">TTM Revenue</th>
            <th className="text-right py-2 font-semibold">% of Total</th>
          </tr>
        </thead>
        <tbody>
          {DEMO_CUSTOMERS.map((cust, idx) => {
            const rowKey = idx < 10 ? `customer_${idx + 1}` : 'all_other_customers'
            const live = getLive(liveCells, 'customer-concentration', rowKey, 'TTM')
            const val = live ?? cust.value
            const pct = DEMO_TOTAL_REV > 0 ? (val / DEMO_TOTAL_REV) * 100 : cust.pct
            const isLast = idx === DEMO_CUSTOMERS.length - 1
            return (
              <tr key={idx} className={`border-b border-gray-100 ${isLast ? 'font-semibold' : ''}`}>
                <td className="py-1.5 text-muted-foreground">{cust.name}</td>
                <td className="text-right py-1.5">
                  <AuditableCell
                    value={fmtFull(val)}
                    source="Customer revenue schedule — TTM"
                    sourceRef={getSrcRef(liveCells, 'customer-concentration', rowKey, 'TTM')}
                    cellId={getCellId(liveCells, 'customer-concentration', rowKey, 'TTM')}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    onViewSource={onViewSource}
                    onSave={onCellSave}
                    onReference={onCellReference}
                    period="TTM"
                  />
                </td>
                <td className={`text-right py-1.5 ${pct >= 15 ? 'text-red-600 font-semibold' : pct >= 10 ? 'text-amber-600 font-semibold' : 'text-muted-foreground'}`}>
                  {fmtPct(pct)}
                </td>
              </tr>
            )
          })}
          <tr className="font-bold">
            <td className="py-2">Total Revenue</td>
            <td className="text-right py-2">{fmtFull(DEMO_TOTAL_REV)}</td>
            <td className="text-right py-2">100.0%</td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── Section V: Quality of Net Assets ────────────────────────────────────────

function NetAssetsSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}) {
  const wcCurrent = WC_DEMO.reduce((s, r) => s + r.current, 0)
  const wcNormalized = WC_DEMO.reduce((s, r) => s + r.normalized, 0)

  const totalDebt = getLive(liveCells, 'net-debt', 'total_debt_like_items', 'TTM') ?? DEMO_TOTAL_DEBT
  const cash = getLive(liveCells, 'net-debt', 'cash_and_equivalents', 'TTM') ?? DEMO_CASH
  const netDebt = getLive(liveCells, 'net-debt', 'net_debt', 'TTM') ?? DEMO_NET_DEBT_FINAL
  const isNetCash = netDebt < 0

  return (
    <div>
      <SectionNumber>Section V</SectionNumber>
      <SectionHeader>Quality of Net Assets</SectionHeader>

      {/* Working Capital */}
      <Narrative>
        Working capital is assessed on a normalized basis, adjusting for seasonality and
        accounting policy differences. Target net working capital (NWC) peg is established
        based on the trailing 12-month average, excluding non-recurring items.
      </Narrative>

      {/* DSO / DIO / DPO inline metrics */}
      <div className="flex gap-8 mb-6">
        <div>
          <p className="text-xs text-muted-foreground">Days Sales Outstanding</p>
          <p className="text-xl font-bold">10.1 days</p>
        </div>
        <div>
          <p className="text-xs text-muted-foreground">Days Inventory Outstanding</p>
          <p className="text-xl font-bold">19.4 days</p>
        </div>
        <div>
          <p className="text-xs text-muted-foreground">Days Payable Outstanding</p>
          <p className="text-xl font-bold">18.3 days</p>
        </div>
      </div>

      <table className="w-full text-sm mb-6">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Item</th>
            <th className="text-right py-2 font-semibold">Current</th>
            <th className="text-right py-2 font-semibold">Normalized</th>
            <th className="text-right py-2 font-semibold">Adjustment</th>
            <th className="text-right py-2 font-semibold">Days</th>
            <th className="text-right py-2 font-semibold">Target Days</th>
          </tr>
        </thead>
        <tbody>
          {WC_DEMO.map((row) => {
            const live = getLive(liveCells, 'working-capital', row.rowKey, 'TTM')
            const val = live ?? row.current
            return (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className="py-1.5 text-muted-foreground">{row.label}</td>
                <td className="text-right py-1.5">
                  <AuditableCell
                    value={fmtFull(val)}
                    source="Balance sheet — working capital analysis"
                    sourceRef={getSrcRef(liveCells, 'working-capital', row.rowKey, 'TTM')}
                    cellId={getCellId(liveCells, 'working-capital', row.rowKey, 'TTM')}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    onViewSource={onViewSource}
                    onSave={onCellSave}
                    onReference={onCellReference}
                    period="TTM"
                  />
                </td>
                <td className="text-right py-1.5 text-muted-foreground">{fmtFull(row.normalized)}</td>
                <td className={`text-right py-1.5 ${row.adj < 0 ? 'text-red-600' : row.adj > 0 ? 'text-emerald-600' : 'text-muted-foreground'}`}>
                  {row.adj !== 0 ? fmtFull(row.adj) : '—'}
                </td>
                <td className="text-right py-1.5 text-muted-foreground">{row.days > 0 ? `${row.days}` : '—'}</td>
                <td className="text-right py-1.5 text-muted-foreground">{row.targetDays > 0 ? `${row.targetDays}` : '—'}</td>
              </tr>
            )
          })}
          <tr className="font-bold border-t border-gray-200">
            <td className="py-2">Net Working Capital</td>
            <td className="text-right py-2">{fmtFull(wcCurrent)}</td>
            <td className="text-right py-2">{fmtFull(wcNormalized)}</td>
            <td className={`text-right py-2 ${wcNormalized - wcCurrent < 0 ? 'text-red-600' : 'text-emerald-600'}`}>
              {fmtFull(wcNormalized - wcCurrent)}
            </td>
            <td className="text-right py-2">—</td>
            <td className="text-right py-2">—</td>
          </tr>
        </tbody>
      </table>

      <Divider />

      {/* Net Debt */}
      <SectionHeader>Net Debt &amp; Debt-Like Items</SectionHeader>
      <Narrative>
        Net debt is calculated as total financial obligations (reported debt plus debt-like items)
        less cash and cash equivalents at closing balance sheet date.{' '}
        {isNetCash && (
          <>
            <CheckCircle className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
            <em>Company is in a net cash position of {fmt(Math.abs(netDebt))}, expected to partially offset purchase price.</em>
          </>
        )}
      </Narrative>

      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Item</th>
            <th className="text-right py-2 font-semibold">Amount (Closing)</th>
          </tr>
        </thead>
        <tbody>
          {DEMO_NET_DEBT.map((row) => {
            const live = getLive(liveCells, 'net-debt', row.rowKey, 'TTM')
            const val = live ?? row.value
            return (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className="py-1.5 text-muted-foreground pl-4">{row.label}</td>
                <td className="text-right py-1.5">
                  <AuditableCell
                    value={fmtFull(val)}
                    source="Balance sheet / closing debt schedule"
                    sourceRef={getSrcRef(liveCells, 'net-debt', row.rowKey, 'TTM')}
                    cellId={getCellId(liveCells, 'net-debt', row.rowKey, 'TTM')}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    onViewSource={onViewSource}
                    onSave={onCellSave}
                    onReference={onCellReference}
                    period="TTM"
                  />
                </td>
              </tr>
            )
          })}
          <tr className="border-b border-gray-100 font-semibold">
            <td className="py-1.5">Total Debt &amp; Debt-Like Items</td>
            <td className="text-right py-1.5">
              <AuditableCell
                value={fmtFull(totalDebt)}
                source="Sum of debt and debt-like items"
                sourceRef={getSrcRef(liveCells, 'net-debt', 'total_debt_like_items', 'TTM')}
                cellId={getCellId(liveCells, 'net-debt', 'total_debt_like_items', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
                onSave={onCellSave}
                onReference={onCellReference}
                period="TTM"
              />
            </td>
          </tr>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground pl-4">Less: Cash &amp; Cash Equivalents</td>
            <td className="text-right py-1.5 text-red-600">
              <AuditableCell
                value={`(${fmtFull(cash)})`}
                source="Bank statements — closing balance"
                sourceRef={getSrcRef(liveCells, 'net-debt', 'cash_and_equivalents', 'TTM')}
                cellId={getCellId(liveCells, 'net-debt', 'cash_and_equivalents', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
                onSave={onCellSave}
                onReference={onCellReference}
                period="TTM"
              />
            </td>
          </tr>
          <tr>
            <td className="py-2 font-bold">Net {isNetCash ? 'Cash' : 'Debt'}</td>
            <td className={`text-right py-2 font-bold ${isNetCash ? 'text-emerald-600' : 'text-red-600'}`}>
              <AuditableCell
                value={isNetCash ? `(${fmtFull(Math.abs(netDebt))})` : fmtFull(netDebt)}
                source="Net debt = total obligations less cash"
                sourceRef={getSrcRef(liveCells, 'net-debt', 'net_debt', 'TTM')}
                cellId={getCellId(liveCells, 'net-debt', 'net_debt', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
                onSave={onCellSave}
                onReference={onCellReference}
                period="TTM"
              />
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── Section VI: Cash Flow & CapEx ────────────────────────────────────────────

function CashFlowSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}) {
  const ttmFcf = getLive(liveCells, 'cash-conversion', 'free_cash_flow', 'TTM') ?? DEMO_CASH_CONV.TTM.fcf
  const ttmEbitda = getLive(liveCells, 'cash-conversion', 'ebitda', 'TTM') ?? DEMO_CASH_CONV.TTM.ebitda
  const conversionRate = ttmEbitda > 0 ? (ttmFcf / ttmEbitda) * 100 : 0

  const cashConvChartData = [
    { period: 'FY22', ebitda: DEMO_CASH_CONV.FY22.ebitda / 1e6, fcf: DEMO_CASH_CONV.FY22.fcf / 1e6 },
    { period: 'TTM',  ebitda: ttmEbitda / 1e6, fcf: ttmFcf / 1e6 },
  ]

  const cashConvRows = [
    { rowKey: 'ebitda', label: 'Adjusted EBITDA', fy22: DEMO_CASH_CONV.FY22.ebitda, ttm: DEMO_CASH_CONV.TTM.ebitda, isBold: false },
    { rowKey: 'capex', label: 'Less: Capital Expenditures', fy22: -DEMO_CASH_CONV.FY22.capex, ttm: -DEMO_CASH_CONV.TTM.capex, isBold: false },
    { rowKey: 'change_in_working_capital', label: 'Less: Change in Working Capital', fy22: -DEMO_CASH_CONV.FY22.wc, ttm: -DEMO_CASH_CONV.TTM.wc, isBold: false },
    { rowKey: 'taxes_paid_cash', label: 'Less: Taxes Paid (Cash)', fy22: -DEMO_CASH_CONV.FY22.taxes, ttm: -DEMO_CASH_CONV.TTM.taxes, isBold: false },
    { rowKey: 'interest_paid_cash', label: 'Less: Interest Paid (Cash)', fy22: -DEMO_CASH_CONV.FY22.interest, ttm: -DEMO_CASH_CONV.TTM.interest, isBold: false },
    { rowKey: 'free_cash_flow', label: 'Free Cash Flow', fy22: DEMO_CASH_CONV.FY22.fcf, ttm: DEMO_CASH_CONV.TTM.fcf, isBold: true },
    { rowKey: 'fcf_to_ebitda_ratio', label: 'FCF / EBITDA Conversion', fy22: (DEMO_CASH_CONV.FY22.fcf / DEMO_CASH_CONV.FY22.ebitda) * 100, ttm: conversionRate, isBold: true, isPercent: true },
  ]

  return (
    <div>
      <SectionNumber>Section VI</SectionNumber>
      <SectionHeader>Cash Flow &amp; CapEx</SectionHeader>
      <Narrative>
        Free cash flow is calculated as Adj. EBITDA less capex, working capital changes,
        cash taxes, and cash interest. High FCF conversion (≥85%) indicates strong cash
        generation with low reinvestment requirements.{' '}
        <TrendingUp className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>
          TTM FCF conversion of {fmtPct(conversionRate)} — driven by low capex intensity
          (~{fmtPct((DEMO_CASH_CONV.TTM.capex / DEMO_CASH_CONV.TTM.ebitda) * 100)} of EBITDA).
        </em>
      </Narrative>

      <div className="h-56 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={cashConvChartData} margin={{ top: 10, right: 20, left: 30, bottom: 5 }} barCategoryGap="30%" barGap={4}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v.toFixed(1)}M`} tick={{ fontSize: 10 }} width={55} />
            <Tooltip formatter={(v: unknown, name: unknown) => [`$${typeof v === 'number' ? v.toFixed(2) : v}M`, name === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow']} />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v) => v === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow'} />
            <Bar dataKey="ebitda" name="ebitda" fill="#1E3A5F" />
            <Bar dataKey="fcf" name="fcf" fill="#10b981" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm mb-8">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Item</th>
            <th className="text-right py-2 font-semibold">FY22</th>
            <th className="text-right py-2 font-semibold">TTM</th>
          </tr>
        </thead>
        <tbody>
          {cashConvRows.map((row) => {
            const liveFy22 = getLive(liveCells, 'cash-conversion', row.rowKey, 'FY22')
            const liveTtm = getLive(liveCells, 'cash-conversion', row.rowKey, 'TTM')
            const fy22Val = liveFy22 ?? row.fy22
            const ttmVal = liveTtm ?? row.ttm
            return (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className={`py-1.5 ${row.isBold ? 'font-bold' : 'text-muted-foreground pl-4'}`}>{row.label}</td>
                <td className={`text-right py-1.5 ${row.isBold ? 'font-bold' : ''}`}>
                  {row.isPercent ? fmtPct(fy22Val) : (
                    <AuditableCell
                      value={fmtFull(fy22Val)}
                      source="Cash flow analysis — FY22"
                      sourceRef={getSrcRef(liveCells, 'cash-conversion', row.rowKey, 'FY22')}
                      cellId={getCellId(liveCells, 'cash-conversion', row.rowKey, 'FY22')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                      onSave={onCellSave}
                      onReference={onCellReference}
                      period="FY22"
                    />
                  )}
                </td>
                <td className={`text-right py-1.5 ${row.isBold ? 'font-bold' : ''}`}>
                  {row.isPercent ? fmtPct(ttmVal) : (
                    <AuditableCell
                      value={fmtFull(ttmVal)}
                      source="Cash flow analysis — TTM"
                      sourceRef={getSrcRef(liveCells, 'cash-conversion', row.rowKey, 'TTM')}
                      cellId={getCellId(liveCells, 'cash-conversion', row.rowKey, 'TTM')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                      onSave={onCellSave}
                      onReference={onCellReference}
                      period="TTM"
                    />
                  )}
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>

      <Divider />

      {/* CapEx Discussion */}
      <SectionHeader>Capital Expenditure Analysis</SectionHeader>
      <Narrative>
        Maintenance CapEx of approximately {fmt(DEMO_CASH_CONV.TTM.capex)} ({fmtPct((DEMO_CASH_CONV.TTM.capex / DEMO_CASH_CONV.TTM.ebitda) * 100)} of
        Adj. EBITDA) consists primarily of equipment replacement and technology refresh. No
        significant growth CapEx was identified in the TTM period. Management has indicated
        no major planned capital projects in the near-term horizon. The business operates with
        an asset-light model; property is leased. CapEx is expected to remain at maintenance
        levels of 4–6% of EBITDA absent strategic investment.
      </Narrative>
    </div>
  )
}

// ─── Section VII: Appendices ──────────────────────────────────────────────────

function AppendicesSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
  onCellSave?: () => void
  onCellReference?: (ctx: { label: string; period: string; displayValue: string }) => void
}) {
  return (
    <div>
      <SectionNumber>Section VII</SectionNumber>
      <SectionHeader>Appendices</SectionHeader>

      {/* Monthly Income Statement */}
      <p className="text-sm font-semibold mb-3">Appendix A — Monthly Income Statement (TTM)</p>
      <div className="overflow-x-auto mb-6">
        <table className="w-full text-xs">
          <thead>
            <tr className="border-b border-gray-200">
              <th className="text-left py-2 font-semibold">($000s)</th>
              {MONTHLY_DEMO.map((m) => (
                <th key={m.month} className="text-right py-2 font-semibold px-1">{m.month}</th>
              ))}
              <th className="text-right py-2 font-semibold px-1">Total</th>
            </tr>
          </thead>
          <tbody>
            {/* Revenue row */}
            <tr className="border-b border-gray-100 font-semibold">
              <td className="py-1.5">Revenue</td>
              {MONTHLY_DEMO.map((m, mi) => {
                const live = getLive(liveCells, 'margins-month', `revenue_${m.month.toLowerCase()}`, 'TTM')
                const val = live ?? m.revenue
                return (
                  <td key={m.month} className="text-right py-1 px-1">
                    <AuditableCell
                      value={`$${Math.round(val / 1000)}K`}
                      source={`Revenue — ${m.month}`}
                      sourceRef={getSrcRef(liveCells, 'margins-month', `revenue_${m.month.toLowerCase()}`, 'TTM')}
                      cellId={getCellId(liveCells, 'margins-month', `revenue_${m.month.toLowerCase()}`, 'TTM')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                      onSave={onCellSave}
                      onReference={onCellReference}
                      period="TTM"
                    />
                  </td>
                )
              })}
              <td className="text-right py-1 px-1 font-bold">
                ${Math.round(MONTHLY_DEMO.reduce((s, m) => s + m.revenue, 0) / 1000)}K
              </td>
            </tr>
            {/* COGS row */}
            <tr className="border-b border-gray-100">
              <td className="py-1.5 text-muted-foreground pl-2">COGS</td>
              {MONTHLY_DEMO.map((m) => {
                const live = getLive(liveCells, 'margins-month', `cogs_${m.month.toLowerCase()}`, 'TTM')
                const val = live ?? m.cogs
                return (
                  <td key={m.month} className="text-right py-1 px-1 text-muted-foreground">
                    <AuditableCell
                      value={`$${Math.round(val / 1000)}K`}
                      source={`COGS — ${m.month}`}
                      sourceRef={getSrcRef(liveCells, 'margins-month', `cogs_${m.month.toLowerCase()}`, 'TTM')}
                      cellId={getCellId(liveCells, 'margins-month', `cogs_${m.month.toLowerCase()}`, 'TTM')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                      onSave={onCellSave}
                      onReference={onCellReference}
                      period="TTM"
                    />
                  </td>
                )
              })}
              <td className="text-right py-1 px-1 text-muted-foreground">
                ${Math.round(MONTHLY_DEMO.reduce((s, m) => s + m.cogs, 0) / 1000)}K
              </td>
            </tr>
            {/* Gross Profit */}
            <tr className="border-b border-gray-100 font-semibold">
              <td className="py-1.5">Gross Profit</td>
              {MONTHLY_DEMO.map((m) => {
                const gp = m.revenue - m.cogs
                return (
                  <td key={m.month} className="text-right py-1 px-1">${Math.round(gp / 1000)}K</td>
                )
              })}
              <td className="text-right py-1 px-1">
                ${Math.round(MONTHLY_DEMO.reduce((s, m) => s + m.revenue - m.cogs, 0) / 1000)}K
              </td>
            </tr>
            {/* OpEx */}
            <tr className="border-b border-gray-100">
              <td className="py-1.5 text-muted-foreground pl-2">OpEx</td>
              {MONTHLY_DEMO.map((m) => (
                <td key={m.month} className="text-right py-1 px-1 text-muted-foreground">${Math.round(m.opex / 1000)}K</td>
              ))}
              <td className="text-right py-1 px-1 text-muted-foreground">
                ${Math.round(MONTHLY_DEMO.reduce((s, m) => s + m.opex, 0) / 1000)}K
              </td>
            </tr>
            {/* EBITDA */}
            <tr className="font-bold">
              <td className="py-1.5">EBITDA</td>
              {MONTHLY_DEMO.map((m) => {
                const ebitda = m.revenue - m.cogs - m.opex
                return (
                  <td key={m.month} className={`text-right py-1 px-1 ${ebitda < 0 ? 'text-red-600' : ''}`}>
                    ${Math.round(ebitda / 1000)}K
                  </td>
                )
              })}
              <td className="text-right py-1 px-1">
                ${Math.round(MONTHLY_DEMO.reduce((s, m) => s + m.revenue - m.cogs - m.opex, 0) / 1000)}K
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <Narrative>
        <strong>Audit Trail Note:</strong> All cell-level edits in this workbook are recorded in an
        immutable audit log. Each entry captures: the original value, the revised value, the
        user who made the change, and a timestamp. The audit log is accessible via the Audit
        tab and can be exported alongside the financial model. Changes made in any tab —
        including this Complete QoE view — are reflected immediately across all other sections
        of the workbook through the shared live data layer.
      </Narrative>
    </div>
  )
}

// ─── Main Component ───────────────────────────────────────────────────────────

export const CompleteQoeSection = memo(function CompleteQoeSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
  onCellSave,
  onCellReference,
}: Props) {
  const sectionRefs = useRef<Record<string, HTMLDivElement | null>>({})
  const [activeAnchor, setActiveAnchor] = useState('sec-exec')

  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((e) => {
          if (e.isIntersecting) setActiveAnchor(e.target.id)
        })
      },
      { threshold: 0.2 }
    )
    Object.values(sectionRefs.current).forEach((el) => el && observer.observe(el))
    return () => observer.disconnect()
  }, [])

  const scrollTo = (id: string) => {
    const el = document.getElementById(id)
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  const setRef = (id: string) => (el: HTMLDivElement | null) => {
    sectionRefs.current[id] = el
  }

  const sharedProps = { liveCells, workbookId, isDemoWorkbook, onViewSource, onCellSave, onCellReference }

  return (
    <div className="flex gap-10 max-w-6xl">
      {/* Sticky Table of Contents */}
      <div className="w-44 shrink-0">
        <div className="sticky top-8 pt-1">
          <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-3">Contents</p>
          {TOC_ITEMS.map((item) => (
            <button
              key={item.id}
              onClick={() => scrollTo(item.id)}
              className={
                activeAnchor === item.id
                  ? 'block w-full text-left text-xs py-1.5 border-l-2 border-[#1E3A5F] pl-2 text-[#1E3A5F] font-semibold'
                  : 'block w-full text-left text-xs py-1.5 pl-3 text-muted-foreground hover:text-foreground'
              }
            >
              {item.label}
            </button>
          ))}
        </div>
      </div>

      {/* Document */}
      <div className="flex-1 min-w-0">
        <h1 className="text-2xl font-bold mb-1">Quality of Earnings Report</h1>
        <p className="text-sm text-muted-foreground mb-10">Confidential — Prepared for Internal Diligence Use</p>

        <div id="sec-exec" ref={setRef('sec-exec')}>
          <ExecSummarySection liveCells={liveCells} />
        </div>

        <MajorDivider />

        <div id="sec-scope" ref={setRef('sec-scope')}>
          <ScopeSection />
        </div>

        <MajorDivider />

        <div id="sec-qoe" ref={setRef('sec-qoe')}>
          <QoeSection {...sharedProps} />
        </div>

        <MajorDivider />

        <div id="sec-revenue" ref={setRef('sec-revenue')}>
          <RevenueSection {...sharedProps} />
        </div>

        <MajorDivider />

        <div id="sec-assets" ref={setRef('sec-assets')}>
          <NetAssetsSection {...sharedProps} />
        </div>

        <MajorDivider />

        <div id="sec-cashflow" ref={setRef('sec-cashflow')}>
          <CashFlowSection {...sharedProps} />
        </div>

        <MajorDivider />

        <div id="sec-appendix" ref={setRef('sec-appendix')}>
          <AppendicesSection {...sharedProps} />
        </div>
      </div>
    </div>
  )
})
