import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { deleteFromGCS } from '@/lib/gcs'

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string; docId: string }> }
) {
  const { id: workbookId, docId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()

  // Verify ownership via workbook
  const { data: doc } = await serviceClient
    .from('documents')
    .select('id, storage_path, gcs_uri, workbook_id')
    .eq('id', docId)
    .eq('workbook_id', workbookId)
    .single()

  if (!doc) return NextResponse.json({ error: 'Document not found' }, { status: 404 })

  // Verify user owns the workbook
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Workbook not found' }, { status: 404 })

  // Delete from GCS if present
  if (doc.gcs_uri) {
    try {
      await deleteFromGCS(doc.gcs_uri)
    } catch (err) {
      console.error('GCS delete failed (continuing):', err)
    }
  }

  // Delete from Supabase Storage if present
  if (doc.storage_path) {
    await serviceClient.storage.from('qoe-documents').remove([doc.storage_path])
  }

  // Delete from DB
  const { error: dbError } = await serviceClient.from('documents').delete().eq('id', docId)
  if (dbError) return NextResponse.json({ error: dbError.message }, { status: 500 })

  return NextResponse.json({ success: true })
}
