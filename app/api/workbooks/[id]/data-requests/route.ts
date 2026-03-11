/**
 * Seller Data Request Engine — Prompt 9 of the agent plan.
 *
 * GET  /api/workbooks/[id]/data-requests
 *   Returns all open missing-data requests, optionally with AI-generated
 *   email text ready to send to the seller.
 *
 * POST /api/workbooks/[id]/data-requests
 *   Generates an AI-written specific data request email from the open
 *   missing_data_requests records, using Gemini. Stores it and returns it.
 *   Optionally accepts { requestIds: string[] } to limit scope.
 */
import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { GoogleAuth } from 'google-auth-library'

export const dynamic = 'force-dynamic'

const MODEL = 'gemini-2.0-flash-001'

async function getAccessToken() {
  const credJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON
  const auth = credJson
    ? new GoogleAuth({ credentials: JSON.parse(credJson), scopes: ['https://www.googleapis.com/auth/cloud-platform'] })
    : new GoogleAuth({ scopes: ['https://www.googleapis.com/auth/cloud-platform'] })
  const client = await auth.getClient()
  const tokenResponse = await (client as { getAccessToken: () => Promise<{ token: string | null }> }).getAccessToken()
  if (!tokenResponse.token) throw new Error('Could not obtain GCP access token')
  return tokenResponse.token
}

async function generateDataRequestEmail(
  companyName: string,
  missingItems: { section: string; field_key: string; period: string | null; reason: string | null; suggested_doc: string | null }[]
): Promise<string> {
  const project = process.env.GOOGLE_CLOUD_PROJECT
  const location = process.env.GOOGLE_CLOUD_LOCATION ?? 'us-central1'
  const token = await getAccessToken()

  const itemsList = missingItems
    .map((m, i) => {
      const period = m.period ? ` for period ${m.period}` : ''
      const doc = m.suggested_doc ? ` (suggested document: ${m.suggested_doc})` : ''
      const reason = m.reason ? ` — ${m.reason}` : ''
      return `${i + 1}. ${m.field_key.replace(/_/g, ' ')}${period}${doc}${reason}`
    })
    .join('\n')

  const prompt = `You are a financial diligence analyst at an M&A advisory firm conducting Quality of Earnings analysis for ${companyName}.

Write a concise, professional data request email to the seller requesting the following missing financial data items. Be specific about exactly what is needed and why. Do not be generic. Reference specific periods, document types, and account names where possible.

Missing items:
${itemsList}

Format:
- Subject line first (prefixed with "Subject: ")
- Then the email body
- Keep it under 300 words
- Professional but direct tone
- End with a request for a specific response deadline (5 business days)`

  const res = await fetch(
    `https://${location}-aiplatform.googleapis.com/v1/projects/${project}/locations/${location}/publishers/google/models/${MODEL}:generateContent`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.3, maxOutputTokens: 512 },
      }),
    }
  )

  if (!res.ok) throw new Error(`Gemini error: ${await res.text()}`)
  const json = await res.json() as { candidates?: { content?: { parts?: { text?: string }[] } }[] }
  return json.candidates?.[0]?.content?.parts?.[0]?.text ?? 'Could not generate email.'
}

// ── GET ───────────────────────────────────────────────────────────────────────

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, company_name')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()
  if (!workbook) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const { withEmail } = Object.fromEntries(new URL(req.url).searchParams)

  const { data: missing, error } = await serviceClient
    .from('missing_data_requests')
    .select('*')
    .eq('workbook_id', workbookId)
    .eq('status', 'open')
    .order('section')

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })

  let emailDraft: string | null = null
  if (withEmail === 'true' && missing && missing.length > 0) {
    try {
      emailDraft = await generateDataRequestEmail(workbook.company_name ?? 'the company', missing)
    } catch (e) {
      emailDraft = `Error generating email: ${String(e)}`
    }
  }

  return NextResponse.json({ missing, count: missing?.length ?? 0, email_draft: emailDraft })
}

// ── POST ──────────────────────────────────────────────────────────────────────

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const serviceClient = createSupabaseServiceClient()
  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, company_name')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()
  if (!workbook) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const body = await req.json().catch(() => ({})) as { requestIds?: string[] }

  let query = serviceClient
    .from('missing_data_requests')
    .select('*')
    .eq('workbook_id', workbookId)
    .eq('status', 'open')

  if (body.requestIds?.length) {
    query = query.in('id', body.requestIds)
  }

  const { data: missing, error } = await query.order('section')
  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  if (!missing || missing.length === 0) {
    return NextResponse.json({ message: 'No open missing data requests found.', email_draft: null })
  }

  try {
    const emailDraft = await generateDataRequestEmail(workbook.company_name ?? 'the company', missing)

    // Parse subject from the draft
    const lines = emailDraft.split('\n').filter(l => l.trim())
    const subjectLine = lines.find(l => l.toLowerCase().startsWith('subject:'))
    const subject = subjectLine ? subjectLine.replace(/^subject:\s*/i, '').trim() : 'Data Request — Financial Diligence'
    const body = subjectLine ? lines.slice(lines.indexOf(subjectLine) + 1).join('\n').trim() : emailDraft

    return NextResponse.json({
      subject,
      body,
      email_draft: emailDraft,
      items_count: missing.length,
      items: missing.map(m => ({ id: m.id, section: m.section, field_key: m.field_key, period: m.period })),
    })
  } catch (e) {
    return NextResponse.json({ error: String(e) }, { status: 500 })
  }
}
