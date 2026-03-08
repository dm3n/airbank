import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ docId: string }> }
) {
  const { docId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()

  // Get document and verify user owns the workbook
  const { data: doc, error: docError } = await serviceClient
    .from('documents')
    .select('id, file_name, storage_path, workbook_id, workbooks!inner(created_by)')
    .eq('id', docId)
    .single()

  if (docError || !doc) return NextResponse.json({ error: 'Document not found' }, { status: 404 })

  const workbookCreatedBy = (doc.workbooks as unknown as { created_by: string } | null)?.created_by
  if (workbookCreatedBy !== user.id) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 })
  }

  if (!doc.storage_path) {
    return NextResponse.json({ error: 'Document not yet stored' }, { status: 409 })
  }

  const { data: signedUrlData, error: signedUrlError } = await serviceClient.storage
    .from('qoe-documents')
    .createSignedUrl(doc.storage_path, 3600) // 1 hour

  if (signedUrlError || !signedUrlData?.signedUrl) {
    return NextResponse.json({ error: 'Failed to generate signed URL' }, { status: 500 })
  }

  return NextResponse.json({
    signedUrl: signedUrlData.signedUrl,
    fileName: doc.file_name,
    expiresIn: 3600,
  })
}
