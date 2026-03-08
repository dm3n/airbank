import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { deleteRagCorpus } from '@/lib/vertex-rag'

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { data, error } = await supabase
    .from('workbooks')
    .select('*, documents(*), rag_corpora(*)')
    .eq('id', id)
    .single()

  if (error) return NextResponse.json({ error: error.message }, { status: 404 })
  return NextResponse.json(data)
}

export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json()
  const allowedFields = ['company_name', 'status', 'missing_fields', 'periods']
  const updates: Record<string, unknown> = {}
  for (const field of allowedFields) {
    if (field in body) updates[field] = body[field]
  }

  const serviceClient = createSupabaseServiceClient()
  const { data, error } = await serviceClient
    .from('workbooks')
    .update(updates)
    .eq('id', id)
    .eq('created_by', user.id)
    .select()
    .single()

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()

  // Fetch corpus name before deleting the row so we can clean it up in Vertex AI
  const { data: corpus } = await serviceClient
    .from('rag_corpora')
    .select('corpus_name')
    .eq('workbook_id', id)
    .single()

  // Delete the workbook row — DB cascades to documents, cells, missing_data_requests, etc.
  const { error } = await serviceClient
    .from('workbooks')
    .delete()
    .eq('id', id)
    .eq('created_by', user.id)

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })

  // Best-effort: delete the Vertex AI RAG corpus so it doesn't orphan on GCP
  if (corpus?.corpus_name) {
    deleteRagCorpus(corpus.corpus_name) // intentionally not awaited
  }

  return new NextResponse(null, { status: 204 })
}
