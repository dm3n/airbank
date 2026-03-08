import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { data, error } = await supabase
    .from('missing_data_requests')
    .select('*')
    .eq('workbook_id', workbookId)
    .eq('status', 'open')
    .order('section')

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}

export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json()
  const { requestId, status } = body

  if (!requestId || !['resolved', 'skipped'].includes(status)) {
    return NextResponse.json({ error: 'Invalid requestId or status' }, { status: 400 })
  }

  const serviceClient = createSupabaseServiceClient()

  // Verify ownership
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Not authorized' }, { status: 403 })

  const { data, error } = await serviceClient
    .from('missing_data_requests')
    .update({ status })
    .eq('id', requestId)
    .eq('workbook_id', workbookId)
    .select()
    .single()

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}
