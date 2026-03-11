import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { uploadToGCS } from '@/lib/gcs'
import { importRagFile } from '@/lib/vertex-rag'

// 2 minutes: GCS upload for large financial PDFs (20-50 MB) can take 60-90 s.
export const maxDuration = 120

const DOC_TYPE_PATTERNS: Record<string, string[]> = {
  general_ledger: ['general ledger', 'generalledger', 'gl_', '_gl ', 'ledger'],
  bank_statements: ['bank statement', 'bankstatement', 'bank_stmt', 'statement'],
  trial_balance: ['trial balance', 'trialbalance', 'trial_bal', 'tb_'],
  financials: ['financial statement', 'financials', 'p&l', 'pnl', 'income statement', 'balance sheet'],
}

function detectDocType(fileName: string): string {
  const lower = fileName.toLowerCase()
  for (const [type, patterns] of Object.entries(DOC_TYPE_PATTERNS)) {
    if (patterns.some((p) => lower.includes(p))) return type
  }
  return 'other'
}

const REQUIRED_DOC_TYPES = ['general_ledger', 'bank_statements', 'trial_balance', 'financials']

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { data, error } = await supabase
    .from('documents')
    .select('*')
    .eq('workbook_id', id)
    .order('created_at', { ascending: true })

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json(data)
}

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  // Verify workbook ownership
  const { data: workbook, error: wbError } = await supabase
    .from('workbooks')
    .select('id, status')
    .eq('id', workbookId)
    .single()

  if (wbError || !workbook) return NextResponse.json({ error: 'Workbook not found' }, { status: 404 })

  const serviceClient = createSupabaseServiceClient()

  const formData = await req.formData()
  const file = formData.get('file') as File | null
  if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 })

  const fileBuffer = Buffer.from(await file.arrayBuffer())
  const fileName = file.name
  const contentType = file.type || 'application/octet-stream'
  const docType = detectDocType(fileName)

  // 1. Create document row (pending)
  const { data: doc, error: docError } = await serviceClient
    .from('documents')
    .insert({
      workbook_id: workbookId,
      file_name: fileName,
      doc_type: docType,
      ingestion_status: 'uploading',
      file_size: fileBuffer.length,
      content_type: contentType,
    })
    .select()
    .single()

  if (docError) return NextResponse.json({ error: docError.message }, { status: 500 })

  const storagePath = `${workbookId}/${doc.id}/${fileName}`
  const gcsPath = `workbooks/${workbookId}/${doc.id}/${fileName}`

  // 2. Upload to Supabase Storage
  const { error: storageError } = await serviceClient.storage
    .from('qoe-documents')
    .upload(storagePath, fileBuffer, { contentType, upsert: true })

  if (storageError) {
    await serviceClient.from('documents').update({ ingestion_status: 'error' }).eq('id', doc.id)
    return NextResponse.json({ error: storageError.message }, { status: 500 })
  }

  // 3. Upload to GCS — synchronous before returning so gcs_uri is guaranteed set.
  //    This is critical: the analyze route's pre-flight re-import step queries
  //    for docs with gcs_uri set but no rag_file_id. If GCS upload were async,
  //    it might not complete before the corpus is created, causing docs to be
  //    permanently skipped from RAG.
  await serviceClient
    .from('documents')
    .update({ ingestion_status: 'ingesting', storage_path: storagePath })
    .eq('id', doc.id)

  let gcsUri: string | null = null
  try {
    gcsUri = await uploadToGCS(fileBuffer, gcsPath, contentType)
    await serviceClient.from('documents').update({ gcs_uri: gcsUri }).eq('id', doc.id)
  } catch (err) {
    console.error('GCS upload failed:', err)
    await serviceClient.from('documents').update({ ingestion_status: 'error' }).eq('id', doc.id)
    const msg = err instanceof Error ? err.message : String(err)
    return NextResponse.json(
      { error: `File saved but failed to reach AI pipeline: ${msg}. Please try uploading again.` },
      { status: 500 }
    )
  }

  // 4. Import into RAG corpus if already ready — async (corpus import takes 30-120s)
  //    If corpus not ready yet, the analyze route's pre-flight step will import it.
  ;(async () => {
    try {
      const { data: corpus } = await serviceClient
        .from('rag_corpora')
        .select('corpus_name')
        .eq('workbook_id', workbookId)
        .single()

      if (corpus?.corpus_name) {
        const ragFileId = await importRagFile(corpus.corpus_name, gcsUri!)
        await serviceClient
          .from('documents')
          .update({ rag_file_id: ragFileId, ingestion_status: 'ready' })
          .eq('id', doc.id)
      } else {
        // Corpus not ready yet — mark ready so the analyze pre-flight picks it up
        await serviceClient
          .from('documents')
          .update({ ingestion_status: 'ready' })
          .eq('id', doc.id)
      }
    } catch (err) {
      console.error('RAG import error (will retry at analysis time):', err)
      // Mark ready so analyze pre-flight can attempt the import
      await serviceClient.from('documents').update({ ingestion_status: 'ready' }).eq('id', doc.id)
    }
  })()

  return NextResponse.json({ ...doc, storage_path: storagePath, gcs_uri: gcsUri }, { status: 201 })
}
