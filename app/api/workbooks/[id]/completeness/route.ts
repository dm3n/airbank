/**
 * GET /api/workbooks/[id]/completeness
 *
 * Returns an overall completeness % and a per-section breakdown.
 * A cell is considered "populated" if raw_value IS NOT NULL.
 * Expected cells = requiredRows × periods (fetched from workbook).
 * Formula-only rows (isCalculated=true with a formula) are excluded —
 * they're auto-filled by the analysis pipeline and don't count against completeness.
 */
import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { SECTION_CONFIGS } from '@/lib/section-prompts'

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params

  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()

  // Verify ownership
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, periods')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const workbookPeriods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']

  // Fetch all non-null cells for this workbook
  const { data: cells } = await serviceClient
    .from('workbook_cells')
    .select('section, row_key, period, raw_value')
    .eq('workbook_id', workbookId)
    .not('raw_value', 'is', null)

  // Build a set of populated (section, row_key, period) triples
  type Key = string
  const populated = new Set<Key>()
  for (const c of cells ?? []) {
    populated.add(`${c.section}::${c.row_key}::${c.period}`)
  }

  let totalExpected = 0
  let totalPopulated = 0

  const sections = SECTION_CONFIGS
    // Skip overview tab (removed from UI) and monthly sections (they have override periods)
    .filter(s => s.key !== 'overview' && !s.overridePeriods)
    .map(section => {
      // Input rows only — formula rows are auto-calculated, not user-input
      const inputRows = section.requiredRows.filter(r => !r.formula && !r.isCalculated)
      const periods = workbookPeriods

      const expected = inputRows.length * periods.length
      let sectionPopulated = 0

      for (const row of inputRows) {
        for (const period of periods) {
          if (populated.has(`${section.key}::${row.rowKey}::${period}`)) {
            sectionPopulated++
          }
        }
      }

      totalExpected += expected
      totalPopulated += sectionPopulated

      return {
        key: section.key,
        displayName: section.displayName,
        expected,
        populated: sectionPopulated,
        pct: expected > 0 ? Math.round((sectionPopulated / expected) * 100) : 100,
      }
    })

  const overall = totalExpected > 0 ? Math.round((totalPopulated / totalExpected) * 100) : 0

  return NextResponse.json({
    overall,
    populated: totalPopulated,
    expected: totalExpected,
    sections,
  })
}
