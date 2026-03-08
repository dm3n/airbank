import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

export async function GET() {
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()

  if (authError || !user) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
  }

  const { data, error } = await supabase
    .from('workbooks')
    .select('id, company_name, status, periods, created_at, updated_at')
    .order('created_at', { ascending: false })

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}

export async function POST(req: NextRequest) {
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()

  if (authError || !user) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
  }

  const body = await req.json()
  const { company_name, periods } = body

  if (!company_name?.trim()) {
    return NextResponse.json({ error: 'company_name is required' }, { status: 400 })
  }

  const serviceClient = createSupabaseServiceClient()

  // Create workbook row
  const { data: workbook, error: wbError } = await serviceClient
    .from('workbooks')
    .insert({
      company_name: company_name.trim(),
      status: 'uploading',
      periods: periods ?? ['FY20', 'FY21', 'FY22', 'TTM'],
      created_by: user.id,
    })
    .select()
    .single()

  if (wbError) return NextResponse.json({ error: wbError.message }, { status: 500 })

  // RAG corpus is created inline by the analyze route (not as a background task here).
  // This ensures corpus creation errors surface via SSE and are always retryable.

  return NextResponse.json(workbook, { status: 201 })
}
