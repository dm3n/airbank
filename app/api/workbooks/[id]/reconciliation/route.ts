/**
 * GET /api/workbooks/[id]/reconciliation
 *
 * Three-way match engine — Prompt 4 of the agent plan.
 * Compares GL vs Tax Return vs Bank Deposits for each available period,
 * returns PASS / WARN / FAIL per reconciliation leg with variance detail.
 */
import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

export const dynamic = 'force-dynamic'

type MatchStatus = 'PASS' | 'WARN' | 'FAIL' | 'UNAVAILABLE'

interface Discrepancy {
  period: string
  field: string
  valueA: number
  valueB: number
  variance: number
  variancePct: number
  description: string
}

interface ReconciliationLeg {
  type: 'GL_vs_Tax' | 'GL_vs_Bank' | 'Bank_vs_Tax'
  labelA: string
  labelB: string
  status: MatchStatus
  overallVariancePct: number | null
  discrepancies: Discrepancy[]
  periodsAvailable: string[]
  confidence: number
}

function reconcilePair(
  aVals: Record<string, number | null>,
  bVals: Record<string, number | null>,
  labelA: string,
  labelB: string,
  type: ReconciliationLeg['type'],
  warnThreshold = 0.02,
  failThreshold = 0.05
): ReconciliationLeg {
  const periods = Object.keys(aVals).filter(p => aVals[p] !== null && bVals[p] !== null)

  if (periods.length === 0) {
    return { type, labelA, labelB, status: 'UNAVAILABLE', overallVariancePct: null, discrepancies: [], periodsAvailable: [], confidence: 0 }
  }

  const discrepancies: Discrepancy[] = []
  let maxVariancePct = 0

  for (const period of periods) {
    const a = aVals[period]!
    const b = bVals[period]!
    if (a === 0 && b === 0) continue
    const variance = a - b
    const base = Math.max(Math.abs(a), Math.abs(b), 1)
    const variancePct = Math.abs(variance) / base

    if (variancePct > maxVariancePct) maxVariancePct = variancePct

    if (variancePct > warnThreshold) {
      discrepancies.push({
        period,
        field: 'Revenue',
        valueA: a,
        valueB: b,
        variance,
        variancePct,
        description: `${labelA} ${period}: $${(a / 1e6).toFixed(3)}M vs ${labelB}: $${(b / 1e6).toFixed(3)}M — ${(variancePct * 100).toFixed(2)}% variance`,
      })
    }
  }

  const status: MatchStatus =
    maxVariancePct >= failThreshold ? 'FAIL' :
    maxVariancePct >= warnThreshold ? 'WARN' : 'PASS'

  const confidence = periods.length >= 3 ? 0.95 : periods.length === 2 ? 0.80 : 0.60

  return { type, labelA, labelB, status, overallVariancePct: maxVariancePct, discrepancies, periodsAvailable: periods, confidence }
}

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()

  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, periods')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const periods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']

  // Pull proof-of-revenue cells (three-way match section)
  const { data: cells } = await serviceClient
    .from('workbook_cells')
    .select('row_key, period, raw_value')
    .eq('workbook_id', workbookId)
    .eq('section', 'proof-revenue')
    .in('row_key', ['gl_revenue', 'tax_return_revenue', 'bank_deposit_revenue'])

  // Build lookup maps per row_key → period → value
  const byKey: Record<string, Record<string, number | null>> = {
    gl_revenue: {},
    tax_return_revenue: {},
    bank_deposit_revenue: {},
  }

  for (const period of periods) {
    byKey.gl_revenue[period] = null
    byKey.tax_return_revenue[period] = null
    byKey.bank_deposit_revenue[period] = null
  }

  for (const cell of cells ?? []) {
    if (byKey[cell.row_key]) {
      byKey[cell.row_key][cell.period] = cell.raw_value
    }
  }

  // Also pull income-statement total revenue as fallback for GL
  const { data: incCells } = await serviceClient
    .from('workbook_cells')
    .select('row_key, period, raw_value')
    .eq('workbook_id', workbookId)
    .eq('section', 'income-statement')
    .eq('row_key', 'total_net_revenue')

  for (const cell of incCells ?? []) {
    if (byKey.gl_revenue[cell.period] === null && cell.raw_value !== null) {
      byKey.gl_revenue[cell.period] = cell.raw_value
    }
  }

  // Run three legs
  const glVsTax = reconcilePair(byKey.gl_revenue, byKey.tax_return_revenue, 'GL Revenue', 'Tax Return Revenue', 'GL_vs_Tax')
  const glVsBank = reconcilePair(byKey.gl_revenue, byKey.bank_deposit_revenue, 'GL Revenue', 'Bank Deposits', 'GL_vs_Bank')
  const bankVsTax = reconcilePair(byKey.bank_deposit_revenue, byKey.tax_return_revenue, 'Bank Deposits', 'Tax Return Revenue', 'Bank_vs_Tax')

  const legs = [glVsTax, glVsBank, bankVsTax]
  const availableLegs = legs.filter(l => l.status !== 'UNAVAILABLE')

  const overallStatus: MatchStatus =
    availableLegs.length === 0 ? 'UNAVAILABLE' :
    availableLegs.some(l => l.status === 'FAIL') ? 'FAIL' :
    availableLegs.some(l => l.status === 'WARN') ? 'WARN' : 'PASS'

  const totalDiscrepancies = legs.flatMap(l => l.discrepancies)

  return NextResponse.json({
    workbook_id: workbookId,
    status: overallStatus,
    legs,
    total_discrepancies: totalDiscrepancies.length,
    critical_count: totalDiscrepancies.filter(d => d.variancePct >= 0.05).length,
    warn_count: totalDiscrepancies.filter(d => d.variancePct >= 0.02 && d.variancePct < 0.05).length,
    generated_at: new Date().toISOString(),
  })
}
