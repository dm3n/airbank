import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient } from '@/lib/supabase-server'
import { chatWithWorkbook, type ChatMessage, type WorkbookContext } from '@/lib/gemini'

export const dynamic = 'force-dynamic'
export const maxDuration = 60

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json().catch(() => ({}))
  const { message, history = [], cellRef } = body as {
    message: string
    history: ChatMessage[]
    cellRef?: { label: string; period: string; displayValue: string }
  }

  if (!message?.trim()) {
    return NextResponse.json({ error: 'message is required' }, { status: 400 })
  }

  // ── Fetch workbook ──────────────────────────────────────────────────────────
  const { data: workbook, error: wbErr } = await supabase
    .from('workbooks')
    .select('id, company_name, periods')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (wbErr || !workbook) {
    return NextResponse.json({ error: 'Workbook not found' }, { status: 404 })
  }

  // ── Fetch cells ─────────────────────────────────────────────────────────────
  const { data: cells } = await supabase
    .from('workbook_cells')
    .select('section, row_key, period, raw_value, display_value, confidence')
    .eq('workbook_id', workbookId)
    .order('section')
    .order('row_key')
    .order('period')

  // ── Fetch flags (all, including resolved for full context) ──────────────────
  const { data: flags } = await supabase
    .from('cell_flags')
    .select('id, section, row_key, period, flag_type, severity, title, body, created_by_ai, resolved_at')
    .eq('workbook_id', workbookId)
    .order('created_at', { ascending: false })

  const ctx: WorkbookContext = {
    companyName: workbook.company_name ?? 'Unknown Company',
    periods: workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM'],
    cells: cells ?? [],
    flags: flags ?? [],
  }

  try {
    const response = await chatWithWorkbook(ctx, history, message.trim(), cellRef)
    return NextResponse.json({ response })
  } catch (err) {
    console.error('[chat] Gemini error:', err)
    return NextResponse.json(
      { error: 'AI is unavailable right now. Please try again.' },
      { status: 503 }
    )
  }
}
