import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

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

  // Format display value
  const displayValue =
    newValue !== null
      ? new Intl.NumberFormat('en-US', {
          style: 'currency',
          currency: 'USD',
          minimumFractionDigits: 0,
        }).format(newValue)
      : ''

  // Update cell
  const { data: updatedCell, error: updateError } = await serviceClient
    .from('workbook_cells')
    .update({
      raw_value: newValue,
      display_value: displayValue,
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

  return NextResponse.json(updatedCell)
}
