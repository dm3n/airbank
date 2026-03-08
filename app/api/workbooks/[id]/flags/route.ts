import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient } from '@/lib/supabase-server'

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { data, error } = await supabase
    .from('cell_flags')
    .select('*, comments:flag_comments(count)')
    .eq('workbook_id', workbookId)
    .order('created_at', { ascending: false })

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })

  // Flatten comment count
  const flags = (data ?? []).map((f) => ({
    ...f,
    comment_count: Array.isArray(f.comments) ? (f.comments[0] as { count: number })?.count ?? 0 : 0,
    comments: undefined,
  }))

  return NextResponse.json(flags)
}

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json().catch(() => ({}))
  const { section, row_key, period, flag_type = 'needs_review', severity = 'warning', title, body: flagBody, cell_id } = body

  if (!section || !row_key || !title) {
    return NextResponse.json({ error: 'section, row_key, and title are required' }, { status: 400 })
  }

  const { data: flag, error } = await supabase
    .from('cell_flags')
    .insert({
      workbook_id: workbookId,
      cell_id: cell_id ?? null,
      section,
      row_key,
      period: period ?? null,
      flag_type,
      severity,
      title,
      body: flagBody ?? null,
      created_by_ai: false,
    })
    .select()
    .single()

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(flag, { status: 201 })
}
