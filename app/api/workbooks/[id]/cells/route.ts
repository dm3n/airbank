import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient } from '@/lib/supabase-server'

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { searchParams } = new URL(req.url)
  const section = searchParams.get('section')

  let query = supabase
    .from('workbook_cells')
    .select('*, source_document:documents(id,file_name,doc_type), flags:cell_flags(id,flag_type,severity,title,resolved_at,created_by_ai)')
    .eq('workbook_id', workbookId)

  if (section) query = query.eq('section', section)

  const { data, error } = await query.order('section').order('row_key').order('period')

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}
