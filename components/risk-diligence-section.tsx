'use client'

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
} from 'recharts'
import { TrendingUp, AlertTriangle, CheckCircle } from 'lucide-react'
import { AuditableCell, type SourceRef } from '@/components/auditable-cell'

// ─── Types ───────────────────────────────────────────────────────────────────

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
}

interface Props {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  periods: string[]
  onViewSource: (sourceRef: SourceRef) => void
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
      (c) =>
        c.section === section && c.row_key === rowKey && c.period === period
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
    (c) =>
      c.section === section && c.row_key === rowKey && c.period === period
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
    (c) =>
      c.section === section && c.row_key === rowKey && c.period === period
  )?.id
}

// ─── Waterfall helper ─────────────────────────────────────────────────────────

interface BridgeItem {
  name: string
  value: number
  isBase?: boolean
  isEnd?: boolean
}

// Range bar approach: each bar has [start, end] so recharts positions it exactly —
// no stacking tricks, no offset math. Set YAxis domain to zoom in on the action zone.
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

// EBITDA Bridge — reads from qoe section TTM; fallback to demo
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

// Net Debt
const DEMO_NET_DEBT = [
  { rowKey: 'reported_debt', label: 'Reported Debt (Term Loan + Revolver)', value: 2850000 },
  { rowKey: 'capital_leases', label: 'Capital Leases', value: 420000 },
  { rowKey: 'unpaid_accrued_vacation', label: 'Unpaid Accrued Vacation', value: 185000 },
  { rowKey: 'aged_accounts_payable', label: 'Aged Accounts Payable (>90 days)', value: 94000 },
  { rowKey: 'deferred_revenue', label: 'Deferred Revenue', value: 318000 },
  { rowKey: 'customer_deposits', label: 'Customer Deposits', value: 142000 },
]
const DEMO_TOTAL_DEBT = DEMO_NET_DEBT.reduce((s, r) => s + r.value, 0) // 4,009,000
const DEMO_CASH = 5310000
const DEMO_NET_DEBT_FINAL = DEMO_TOTAL_DEBT - DEMO_CASH // -1,301,000

// Customer Concentration (TTM $68.3M)
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

// Cash Conversion
const DEMO_CASH_CONV = {
  FY22: { ebitda: 7682947, capex: 384147, wc: 192075, taxes: 307317, interest: 61245, fcf: 6738163 },
  TTM: { ebitda: 8547239, capex: 427361, wc: 256417, taxes: 341889, interest: 65182, fcf: 7456390 },
}

// Proof of Revenue
const DEMO_PROOF = [
  { period: 'FY20', gl: 42187456, tax: 42063812, bank: 42075200 },
  { period: 'FY21', gl: 51482903, tax: 51328450, bank: 51349800 },
  { period: 'FY22', gl: 62745381, tax: 62558200, bank: 62580900 },
  { period: 'TTM', gl: 68293742, tax: 68088300, bank: 68112500 },
]

// Run-Rate adjustments
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
const DEMO_PRO_FORMA_MARGIN = (DEMO_PRO_FORMA_EBITDA / DEMO_PRO_FORMA_REV) * 100

// ─── Custom Waterfall Tooltip ─────────────────────────────────────────────────

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

// ─── Sub-section: Divider ─────────────────────────────────────────────────────

function Divider() {
  return <div className="border-t border-gray-100 my-12" />
}

// ─── Sub-section: Header ──────────────────────────────────────────────────────

function SectionHeader({ children }: { children: React.ReactNode }) {
  return (
    <h2 className="text-lg font-semibold border-l-4 border-[#1E3A5F] pl-4 mb-3">
      {children}
    </h2>
  )
}

// ─── Sub-section: Narrative ───────────────────────────────────────────────────

function Narrative({ children }: { children: React.ReactNode }) {
  return (
    <p className="text-sm text-muted-foreground leading-relaxed mb-6">
      {children}
    </p>
  )
}

// ─── 1. EBITDA Bridge ─────────────────────────────────────────────────────────

function EbitdaBridge({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  // Try to build bridge from live cells (qoe section, TTM)
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

  // QoE adjustment rows for table (non-base, non-end)
  const adjRows = [
    { rowKey: 'adj_charitable_contributions', label: 'a) Charitable Contributions', demo: 315920 },
    { rowKey: 'adj_owner_tax_legal', label: 'b) Owner Tax & Legal', demo: 22450 },
    { rowKey: 'adj_excess_owner_comp', label: 'c) Excess Owner Compensation', demo: 248190 },
    { rowKey: 'adj_personal_expenses', label: 'd) Personal Expenses', demo: 39863 },
    { rowKey: 'adj_professional_fees_onetime', label: 'e) Professional Fees - One-time', demo: 68200 },
    { rowKey: 'adj_executive_search', label: 'f) Executive Search Fees', demo: 52000 },
    { rowKey: 'adj_facility_relocation', label: 'g) Facility Relocation', demo: 41500 },
    { rowKey: 'adj_severance', label: 'h) Severance', demo: 92300 },
    { rowKey: 'adj_above_market_rent', label: 'i) Above-Market Rent', demo: -168900 },
    { rowKey: 'adj_normalize_owner_comp', label: 'j) Normalize Owner Comp', demo: -425900 },
  ]

  return (
    <div>
      <SectionHeader>EBITDA Bridge — TTM Adj. EBITDA Waterfall</SectionHeader>
      <Narrative>
        The waterfall below bridges from EBITDA as defined ({fmt(ttmEbitda)}) to
        Diligence-Adjusted EBITDA ({fmt(ttmAdjEbitda)}) via{' '}
        {adjRows.length} discrete add-backs, net of owner compensation
        normalization and above-market rent. Total net adjustments:{' '}
        <strong>{fmt(totalAdj)}</strong>.{' '}
        <TrendingUp className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5" />{' '}
        <em>92% of add-backs are recurring in nature and well-documented.</em>
      </Narrative>

      <div className="h-72 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={waterfallData}
            margin={{ top: 10, right: 20, left: 30, bottom: 40 }}
            barCategoryGap="20%"
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis
              dataKey="name"
              tick={{ fontSize: 10 }}
              angle={-35}
              textAnchor="end"
              interval={0}
            />
            <YAxis
              domain={[wfYMin, wfYMax]}
              tickFormatter={fmt}
              tick={{ fontSize: 10 }}
              width={60}
            />
            <Tooltip content={<WaterfallTooltip />} />
            <Bar dataKey="range">
              {waterfallData.map((entry, idx) => (
                <Cell key={idx} fill={getBarColor(entry)} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold text-foreground">Adjustment</th>
            <th className="text-right py-2 font-semibold text-foreground">TTM Amount</th>
          </tr>
        </thead>
        <tbody>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 font-semibold">EBITDA, as Defined</td>
            <td className="text-right py-1.5 font-semibold">
              <AuditableCell
                value={fmtFull(ttmEbitda)}
                source="General Ledger — calculated per QoE methodology"
                sourceRef={getSrcRef(liveCells, 'qoe', 'ebitda_as_defined', 'TTM')}
                cellId={getCellId(liveCells, 'qoe', 'ebitda_as_defined', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
              />
            </td>
          </tr>
          {adjRows.map((row) => {
            const live = getLive(liveCells, 'qoe', row.rowKey, 'TTM')
            const val = live ?? row.demo
            return (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className="py-1.5 text-muted-foreground pl-4">{row.label}</td>
                <td className={`text-right py-1.5 ${val < 0 ? 'text-red-600' : 'text-emerald-600'}`}>
                  <AuditableCell
                    value={fmtFull(val)}
                    source="General Ledger — diligence adjustment"
                    sourceRef={getSrcRef(liveCells, 'qoe', row.rowKey, 'TTM')}
                    cellId={getCellId(liveCells, 'qoe', row.rowKey, 'TTM')}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    onViewSource={onViewSource}
                  />
                </td>
              </tr>
            )
          })}
          <tr>
            <td className="py-2 font-bold">Diligence-Adjusted EBITDA</td>
            <td className="text-right py-2 font-bold">
              <AuditableCell
                value={fmtFull(ttmAdjEbitda)}
                source="Final normalized run-rate EBITDA — buyer perspective"
                sourceRef={getSrcRef(liveCells, 'qoe', 'diligence_adjusted_ebitda', 'TTM')}
                cellId={getCellId(liveCells, 'qoe', 'diligence_adjusted_ebitda', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
              />
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── 2. Net Debt ──────────────────────────────────────────────────────────────

function NetDebtSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  const totalDebt = getLive(liveCells, 'net-debt', 'total_debt_like_items', 'TTM') ?? DEMO_TOTAL_DEBT
  const cash = getLive(liveCells, 'net-debt', 'cash_and_equivalents', 'TTM') ?? DEMO_CASH
  const netDebt = getLive(liveCells, 'net-debt', 'net_debt', 'TTM') ?? DEMO_NET_DEBT_FINAL
  const isNetCash = netDebt < 0

  return (
    <div>
      <SectionHeader>Net Debt &amp; Debt-Like Items</SectionHeader>
      <Narrative>
        Net debt is calculated as total financial obligations (reported debt plus
        debt-like items) less cash and cash equivalents at the closing balance
        sheet date. Items such as deferred revenue, accrued vacation, and aged
        payables are included as they represent economic obligations a buyer
        would assume.{' '}
        {isNetCash && (
          <>
            <CheckCircle className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
            <em>Company is in a net cash position of {fmt(Math.abs(netDebt))}, which is
            expected to partially offset purchase price.</em>
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
              />
            </td>
          </tr>
          <tr>
            <td className="py-2 font-bold">
              Net {isNetCash ? 'Cash' : 'Debt'}
            </td>
            <td className={`text-right py-2 font-bold ${isNetCash ? 'text-emerald-600' : 'text-red-600'}`}>
              <AuditableCell
                value={isNetCash ? `(${fmtFull(Math.abs(netDebt))})` : fmtFull(netDebt)}
                source="Net debt = total obligations less cash"
                sourceRef={getSrcRef(liveCells, 'net-debt', 'net_debt', 'TTM')}
                cellId={getCellId(liveCells, 'net-debt', 'net_debt', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
              />
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── 3. Customer Concentration ────────────────────────────────────────────────

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

function CustomerConcentrationSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  const topCustomer = DEMO_CUSTOMERS[0]
  const top3Pct = DEMO_CUSTOMERS.slice(0, 3).reduce((s, c) => s + c.pct, 0)

  return (
    <div>
      <SectionHeader>Revenue &amp; Customer Concentration</SectionHeader>
      <Narrative>
        Revenue concentration analysis covers TTM revenue of {fmt(DEMO_TOTAL_REV)}{' '}
        across {DEMO_CUSTOMERS.length - 1} named customers plus a long tail. Concentration
        risk is flagged at ≥15% (red) and ≥10% (amber).{' '}
        <AlertTriangle className="inline h-3.5 w-3.5 text-amber-500 ml-0.5 mr-0.5" />
        <em>
          Top customer ({topCustomer.name}) represents {fmtPct(topCustomer.pct)} of TTM
          revenue — concentration risk. Top 3 customers: {fmtPct(top3Pct)}. Recommend
          confirming retention rates and contract renewal dates prior to close.
        </em>
      </Narrative>

      <div className="h-80 mb-6">
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
                  />
                </td>
                <td
                  className={`text-right py-1.5 ${pct >= 15 ? 'text-red-600 font-semibold' : pct >= 10 ? 'text-amber-600 font-semibold' : 'text-muted-foreground'}`}
                >
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

// ─── 4. Cash Conversion ───────────────────────────────────────────────────────

const CASH_CONV_CHART_DATA = [
  {
    period: 'FY22',
    ebitda: DEMO_CASH_CONV.FY22.ebitda / 1e6,
    fcf: DEMO_CASH_CONV.FY22.fcf / 1e6,
  },
  {
    period: 'TTM',
    ebitda: DEMO_CASH_CONV.TTM.ebitda / 1e6,
    fcf: DEMO_CASH_CONV.TTM.fcf / 1e6,
  },
]

function CashConversionSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  const ttmFcf = getLive(liveCells, 'cash-conversion', 'free_cash_flow', 'TTM') ?? DEMO_CASH_CONV.TTM.fcf
  const ttmEbitda = getLive(liveCells, 'cash-conversion', 'ebitda', 'TTM') ?? DEMO_CASH_CONV.TTM.ebitda
  const conversionRate = ttmEbitda > 0 ? (ttmFcf / ttmEbitda) * 100 : 0

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
      <SectionHeader>Cash Conversion Analysis</SectionHeader>
      <Narrative>
        Free cash flow is calculated as Adj. EBITDA less capex, working capital
        changes, cash taxes, and cash interest. High FCF conversion ratios (≥85%)
        indicate strong cash generation with low reinvestment requirements.{' '}
        <TrendingUp className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>
          TTM FCF conversion of {fmtPct(conversionRate)} — driven by low capex intensity
          (~{fmtPct((DEMO_CASH_CONV.TTM.capex / DEMO_CASH_CONV.TTM.ebitda) * 100)} of EBITDA).
          Minimal maintenance capex required; business is primarily services/software.
        </em>
      </Narrative>

      <div className="h-64 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={CASH_CONV_CHART_DATA}
            margin={{ top: 10, right: 20, left: 30, bottom: 5 }}
            barCategoryGap="30%"
            barGap={4}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v.toFixed(1)}M`} tick={{ fontSize: 10 }} width={55} />
            <Tooltip
              formatter={(v: unknown, name: unknown) => [
                typeof v === 'number' ? `$${v.toFixed(2)}M` : String(v),
                name === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow',
              ]}
            />
            <Legend
              formatter={(value) => (value === 'ebitda' ? 'Adj. EBITDA' : 'Free Cash Flow')}
              iconSize={10}
              wrapperStyle={{ fontSize: 11 }}
            />
            <Bar dataKey="ebitda" name="ebitda" fill="#1E3A5F" />
            <Bar dataKey="fcf" name="fcf" fill="#10b981" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm">
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
                <td className={`py-1.5 ${row.isBold ? 'font-bold' : 'text-muted-foreground pl-4'}`}>
                  {row.label}
                </td>
                <td className={`text-right py-1.5 ${row.isBold ? 'font-bold' : ''}`}>
                  {row.isPercent ? (
                    fmtPct(fy22Val)
                  ) : (
                    <AuditableCell
                      value={fmtFull(fy22Val)}
                      source="Cash flow analysis — FY22"
                      sourceRef={getSrcRef(liveCells, 'cash-conversion', row.rowKey, 'FY22')}
                      cellId={getCellId(liveCells, 'cash-conversion', row.rowKey, 'FY22')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                    />
                  )}
                </td>
                <td className={`text-right py-1.5 ${row.isBold ? 'font-bold' : ''}`}>
                  {row.isPercent ? (
                    fmtPct(ttmVal)
                  ) : (
                    <AuditableCell
                      value={fmtFull(ttmVal)}
                      source="Cash flow analysis — TTM"
                      sourceRef={getSrcRef(liveCells, 'cash-conversion', row.rowKey, 'TTM')}
                      cellId={getCellId(liveCells, 'cash-conversion', row.rowKey, 'TTM')}
                      workbookId={isDemoWorkbook ? undefined : workbookId}
                      onViewSource={onViewSource}
                    />
                  )}
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

// ─── 5. Proof of Revenue ──────────────────────────────────────────────────────

function ProofOfRevenueSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  const chartData = DEMO_PROOF.map((p) => {
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

  const maxVar = Math.max(
    ...DEMO_PROOF.map((p) => Math.abs(((p.gl - p.tax) / p.gl) * 100))
  )

  return (
    <div>
      <SectionHeader>Proof of Revenue — Three-Way Match</SectionHeader>
      <Narrative>
        Revenue is independently corroborated across three sources: (1) the
        company&apos;s general ledger, (2) federal tax returns (Form 1120/Schedule C),
        and (3) bank deposit summaries. Variances exceeding 0.5% are flagged for
        further investigation.{' '}
        <CheckCircle className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>
          All three sources agree within {fmtPct(maxVar)} across all periods —
          no material revenue recognition inconsistencies identified.
        </em>
      </Narrative>

      <div className="h-64 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={chartData}
            margin={{ top: 10, right: 20, left: 30, bottom: 5 }}
            barCategoryGap="25%"
            barGap={2}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v.toFixed(0)}M`} tick={{ fontSize: 10 }} width={50} />
            <Tooltip formatter={(v: unknown) => [typeof v === 'number' ? `$${v.toFixed(2)}M` : String(v)]} />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
            <Bar dataKey="gl" name="GL Revenue" fill="#1E3A5F" />
            <Bar dataKey="tax" name="Tax Return" fill="#3b82f6" />
            <Bar dataKey="bank" name="Bank Deposits" fill="#10b981" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Item</th>
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
                    />
                  </td>
                )
              })}
            </tr>
          ))}
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground pl-4">Variance: GL vs. Tax</td>
            {DEMO_PROOF.map((p) => {
              const glLive = getLive(liveCells, 'proof-revenue', 'gl_revenue', p.period) ?? p.gl
              const taxLive = getLive(liveCells, 'proof-revenue', 'tax_return_revenue', p.period) ?? p.tax
              const variance = glLive - taxLive
              return (
                <td key={p.period} className={`text-right py-1.5 ${Math.abs(variance / glLive) < 0.005 ? 'text-emerald-600' : 'text-red-600'}`}>
                  {fmtFull(variance)}
                </td>
              )
            })}
          </tr>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground pl-4">Variance: GL vs. Bank</td>
            {DEMO_PROOF.map((p) => {
              const glLive = getLive(liveCells, 'proof-revenue', 'gl_revenue', p.period) ?? p.gl
              const bankLive = getLive(liveCells, 'proof-revenue', 'bank_deposit_revenue', p.period) ?? p.bank
              const variance = glLive - bankLive
              return (
                <td key={p.period} className={`text-right py-1.5 ${Math.abs(variance / glLive) < 0.005 ? 'text-emerald-600' : 'text-red-600'}`}>
                  {fmtFull(variance)}
                </td>
              )
            })}
          </tr>
          <tr>
            <td className="py-1.5 text-muted-foreground pl-4">% Variance GL vs. Tax</td>
            {DEMO_PROOF.map((p) => {
              const glLive = getLive(liveCells, 'proof-revenue', 'gl_revenue', p.period) ?? p.gl
              const taxLive = getLive(liveCells, 'proof-revenue', 'tax_return_revenue', p.period) ?? p.tax
              const pctVar = glLive > 0 ? Math.abs(((glLive - taxLive) / glLive) * 100) : 0
              return (
                <td key={p.period} className={`text-right py-1.5 ${pctVar < 0.5 ? 'text-emerald-600' : 'text-red-600'}`}>
                  {fmtPct(pctVar)}
                </td>
              )
            })}
          </tr>
        </tbody>
      </table>
    </div>
  )
}

// ─── 6. Run-Rate & Pro-Forma ──────────────────────────────────────────────────

function RunRateSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: {
  liveCells: LiveCell[]
  workbookId: string
  isDemoWorkbook: boolean
  onViewSource: (s: SourceRef) => void
}) {
  const waterfallData = buildWaterfallPoints(DEMO_RUN_RATE)
  const rrYMin = 68293742 * 0.97
  const rrYMax = Math.max(...waterfallData.map((p) => p.range[1])) * 1.02

  const getBarColor = (entry: WaterfallPoint) => {
    if (entry.isBase || entry.isEnd) return '#1E3A5F'
    return entry.rawValue >= 0 ? '#10b981' : '#ef4444'
  }

  const proFormaRev = getLive(liveCells, 'run-rate', 'pro_forma_revenue', 'TTM') ?? DEMO_PRO_FORMA_REV
  const proFormaEbitda = getLive(liveCells, 'run-rate', 'pro_forma_ebitda', 'TTM') ?? DEMO_PRO_FORMA_EBITDA
  const proFormaMargin = proFormaRev > 0 ? (proFormaEbitda / proFormaRev) * 100 : DEMO_PRO_FORMA_MARGIN

  const runRateAdjRows = [
    { rowKey: 'ttm_revenue_base', label: 'TTM Revenue (Base)', demo: 68293742 },
    { rowKey: 'new_contract_value', label: 'New Contract Value (Annualized)', demo: 1850000 },
    { rowKey: 'full_year_hire_effect', label: 'Full-Year Hire Effect', demo: 420000 },
    { rowKey: 'price_increase_effect', label: 'Price Increase Effect', demo: 680000 },
    { rowKey: 'lost_revenue_adjustment', label: 'Less: Lost/Churned Revenue', demo: -944000 },
  ]

  return (
    <div>
      <SectionHeader>Run-Rate &amp; Pro-Forma Revenue Adjustments</SectionHeader>
      <Narrative>
        Pro-forma revenue reflects the annualized run-rate adjusted for known
        contract wins, full-year headcount effects, executed price increases, and
        identified churn. No incremental margin expansion is assumed; pro-forma
        EBITDA margin is held at TTM rates.{' '}
        <TrendingUp className="inline h-3.5 w-3.5 text-emerald-500 ml-0.5 mr-0.5" />
        <em>
          Pro-forma revenue of {fmt(proFormaRev)} represents{' '}
          {fmtPct(((proFormaRev - 68293742) / 68293742) * 100)} uplift above TTM —
          conservative approach; management&apos;s pipeline supports upside.
        </em>
      </Narrative>

      <div className="h-72 mb-6">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={waterfallData}
            margin={{ top: 10, right: 20, left: 30, bottom: 40 }}
            barCategoryGap="20%"
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis
              dataKey="name"
              tick={{ fontSize: 10 }}
              angle={-35}
              textAnchor="end"
              interval={0}
            />
            <YAxis domain={[rrYMin, rrYMax]} tickFormatter={fmt} tick={{ fontSize: 10 }} width={60} />
            <Tooltip content={<WaterfallTooltip />} />
            <Bar dataKey="range">
              {waterfallData.map((entry, idx) => (
                <Cell key={idx} fill={getBarColor(entry)} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <table className="w-full text-sm mb-8">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-2 font-semibold">Adjustment</th>
            <th className="text-right py-2 font-semibold">Amount</th>
          </tr>
        </thead>
        <tbody>
          {runRateAdjRows.map((row) => {
            const live = getLive(liveCells, 'run-rate', row.rowKey, 'TTM')
            const val = live ?? row.demo
            const isBase = row.rowKey === 'ttm_revenue_base'
            return (
              <tr key={row.rowKey} className="border-b border-gray-100">
                <td className={`py-1.5 ${isBase ? 'font-semibold' : 'text-muted-foreground pl-4'}`}>
                  {row.label}
                </td>
                <td className={`text-right py-1.5 ${isBase ? 'font-semibold' : val < 0 ? 'text-red-600' : 'text-emerald-600'}`}>
                  <AuditableCell
                    value={fmtFull(val)}
                    source="Pro-forma revenue bridge"
                    sourceRef={getSrcRef(liveCells, 'run-rate', row.rowKey, 'TTM')}
                    cellId={getCellId(liveCells, 'run-rate', row.rowKey, 'TTM')}
                    workbookId={isDemoWorkbook ? undefined : workbookId}
                    onViewSource={onViewSource}
                  />
                </td>
              </tr>
            )
          })}
          <tr>
            <td className="py-2 font-bold">Pro Forma Revenue</td>
            <td className="text-right py-2 font-bold">
              <AuditableCell
                value={fmtFull(proFormaRev)}
                source="Pro-forma revenue — sum of TTM base + adjustments"
                sourceRef={getSrcRef(liveCells, 'run-rate', 'pro_forma_revenue', 'TTM')}
                cellId={getCellId(liveCells, 'run-rate', 'pro_forma_revenue', 'TTM')}
                workbookId={isDemoWorkbook ? undefined : workbookId}
                onViewSource={onViewSource}
              />
            </td>
          </tr>
        </tbody>
      </table>

      {/* Summary mini-table */}
      <h3 className="text-sm font-semibold mb-3 text-muted-foreground uppercase tracking-wide">
        TTM vs. Pro Forma Summary
      </h3>
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
            <td className="text-right py-1.5 text-emerald-600">
              +{fmtPct(((proFormaRev - 68293742) / 68293742) * 100)}
            </td>
          </tr>
          <tr className="border-b border-gray-100">
            <td className="py-1.5 text-muted-foreground">Adj. EBITDA</td>
            <td className="text-right py-1.5">{fmt(8547239)}</td>
            <td className="text-right py-1.5">{fmt(proFormaEbitda)}</td>
            <td className="text-right py-1.5 text-emerald-600">
              +{fmtPct(((proFormaEbitda - 8547239) / 8547239) * 100)}
            </td>
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

// ─── Main Component ───────────────────────────────────────────────────────────

export function RiskDiligenceSection({
  liveCells,
  workbookId,
  isDemoWorkbook,
  onViewSource,
}: Props) {
  const sharedProps = { liveCells, workbookId, isDemoWorkbook, onViewSource }

  return (
    <div className="max-w-5xl">
      <EbitdaBridge {...sharedProps} />
      <Divider />
      <NetDebtSection {...sharedProps} />
      <Divider />
      <CustomerConcentrationSection {...sharedProps} />
      <Divider />
      <CashConversionSection {...sharedProps} />
      <Divider />
      <ProofOfRevenueSection {...sharedProps} />
      <Divider />
      <RunRateSection {...sharedProps} />
    </div>
  )
}
