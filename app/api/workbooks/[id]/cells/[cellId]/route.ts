import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { resolveFormulas, CellValue } from '@/lib/formula-resolver'

export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string; cellId: string }> }
) {
  const { id: workbookId, cellId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json()
  const { raw_value, note } = body

  if (raw_value === undefined) {
    return NextResponse.json({ error: 'raw_value is required' }, { status: 400 })
  }

  const serviceClient = createSupabaseServiceClient()

  // Verify ownership via workbook
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Not authorized' }, { status: 403 })

  // Read old value
  const { data: cell, error: cellError } = await serviceClient
    .from('workbook_cells')
    .select('*')
    .eq('id', cellId)
    .eq('workbook_id', workbookId)
    .single()

  if (cellError || !cell) return NextResponse.json({ error: 'Cell not found' }, { status: 404 })

  const newValue = typeof raw_value === 'string' ? parseFloat(raw_value) : raw_value

  const formatDisplay = (v: number | null, valueType?: string): string => {
    if (v === null || v === undefined) return ''
    if (valueType === 'percent') return `${v.toFixed(1)}%`
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
    }).format(v)
  }

  // Update the edited cell
  const { data: updatedCell, error: updateError } = await serviceClient
    .from('workbook_cells')
    .update({
      raw_value: newValue,
      display_value: formatDisplay(newValue),
      is_overridden: true,
    })
    .eq('id', cellId)
    .select()
    .single()

  if (updateError) return NextResponse.json({ error: updateError.message }, { status: 500 })

  // Insert audit entry
  await serviceClient.from('audit_entries').insert({
    cell_id: cellId,
    workbook_id: workbookId,
    old_value: cell.raw_value,
    new_value: newValue,
    note: note ?? null,
    edited_by: user.id,
  })

  // ── Formula waterfall recalculation ──────────────────────────────
  const section: string = cell.section
  if (section) {
    // Fetch all cells for this section
    const { data: sectionCells } = await serviceClient
      .from('workbook_cells')
      .select('id, row_key, period, raw_value')
      .eq('workbook_id', workbookId)
      .eq('section', section)

    if (sectionCells && sectionCells.length > 0) {
      // Merge the just-saved value into the list before resolving
      const cellList: CellValue[] = sectionCells.map(c => ({
        rowKey: c.row_key,
        period: c.period,
        rawValue: c.id === cellId ? newValue : (c.raw_value as number | null),
      }))

      const recalculated = resolveFormulas(section, cellList)

      if (recalculated.length > 0) {
        // Fetch the section config to get valueType for display formatting
        const { getSectionConfig } = await import('@/lib/section-prompts')
        const config = getSectionConfig(section)
        const rowTypeMap = new Map(config?.requiredRows.map(r => [r.rowKey, r.valueType]) ?? [])

        // Build a lookup from (row_key, period) → cell id
        const idLookup = new Map(sectionCells.map(c => [`${c.row_key}::${c.period}`, c.id]))

        for (const rc of recalculated) {
          const existingId = idLookup.get(`${rc.rowKey}::${rc.period}`)
          const display = formatDisplay(rc.rawValue, rowTypeMap.get(rc.rowKey))

          if (existingId) {
            await serviceClient
              .from('workbook_cells')
              .update({ raw_value: rc.rawValue, display_value: display })
              .eq('id', existingId)
          } else {
            await serviceClient
              .from('workbook_cells')
              .insert({
                workbook_id: workbookId,
                section,
                row_key: rc.rowKey,
                period: rc.period,
                raw_value: rc.rawValue,
                display_value: display,
                is_overridden: false,
              })
          }
        }
      }
    }
  }

  return NextResponse.json(updatedCell)
}
