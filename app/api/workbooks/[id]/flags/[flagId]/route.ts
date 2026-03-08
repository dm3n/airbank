import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient } from '@/lib/supabase-server'

export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string; flagId: string }> }
) {
  const { id: workbookId, flagId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json().catch(() => ({}))
  const updates: Record<string, unknown> = {}

  if (body.resolved === true) {
    updates.resolved_at = new Date().toISOString()
    updates.resolved_by = user.id
  } else if (body.resolved === false) {
    updates.resolved_at = null
    updates.resolved_by = null
  }
  if (body.title !== undefined) updates.title = body.title
  if (body.body !== undefined) updates.body = body.body

  const { data, error } = await supabase
    .from('cell_flags')
    .update(updates)
    .eq('id', flagId)
    .eq('workbook_id', workbookId)
    .select()
    .single()

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string; flagId: string }> }
) {
  const { id: workbookId, flagId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { error } = await supabase
    .from('cell_flags')
    .delete()
    .eq('id', flagId)
    .eq('workbook_id', workbookId)

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json({ success: true })
}
