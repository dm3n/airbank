/**
 * Gemini extraction — direct REST API for full timeout + retry control.
 *
 * Accuracy priorities (KPMG / Deloitte standard):
 *  1. Unit-scale detection  — "(in thousands)" tables must be ×1,000
 *  2. Parenthetical negatives — (1,234) = -1,234 in accounting notation
 *  3. Period mapping          — fiscal year labels → FY20/FY21/FY22/TTM
 *  4. Verbatim source excerpt — every number is human-verifiable
 *  5. Conservative missing    — only flag if truly absent from all chunks
 */

import type { RagChunk } from './vertex-rag'
import type { SectionConfig } from './section-prompts'

export interface ExtractedCell {
  row_key: string
  period: string
  raw_value: number | null
  display_value: string
  source_page: number | null
  source_excerpt: string
  source_filename: string
  confidence: number
  is_calculated: boolean
}

export interface MissingField {
  row_key: string
  period: string
  reason: string
  suggested_doc: string
}

export interface ExtractedFlag {
  row_key: string
  period: string
  flag_type: 'low_confidence' | 'discrepancy' | 'needs_review' | 'ai_note'
  severity: 'info' | 'warning' | 'critical'
  title: string
  body: string
}

// ── GCP auth ──────────────────────────────────────────────────────────────────

async function getAccessToken(): Promise<string> {
  const { GoogleAuth } = await import('google-auth-library')
  const credentialsJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON
  const credentials = credentialsJson ? JSON.parse(credentialsJson) : undefined
  const auth = new GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/cloud-platform'],
  })
  const client = await auth.getClient()
  const tokenResponse = await client.getAccessToken()
  if (!tokenResponse.token) throw new Error('Failed to get GCP access token')
  return tokenResponse.token
}

const LOCATION = () => process.env.GOOGLE_CLOUD_LOCATION || 'us-west1'
const PROJECT  = () => {
  const p = process.env.GOOGLE_CLOUD_PROJECT
  if (!p) throw new Error('GOOGLE_CLOUD_PROJECT env var not set')
  return p
}
const MODEL = 'gemini-2.0-flash-001'

// ── System instruction (loaded once, reused across all calls) ─────────────────

const SYSTEM_INSTRUCTION = `\
You are a senior financial diligence analyst at a Big-4 accounting firm.
Your task is to extract numeric data from uploaded financial documents into a
Quality of Earnings (QoE) workbook with audit-quality precision.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 1 — UNIT SCALE (most common extraction error — do this first)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Before reading any numbers, scan document headers, footnotes, and column
labels for unit indicators and apply a multiplier to every value:

  "(in thousands)" / "($ thousands)" / "(000s)" / "$ in 000s"  → ×1,000
  "(in millions)"  / "($ millions)"  / "(MM)"                  → ×1,000,000
  "(in billions)"                                               → ×1,000,000,000
  "$" prefix on individual cells / no scale indicator           → ×1 (raw dollars)

Apply the detected scale consistently to EVERY numeric value in that document.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 2 — NEGATIVE NUMBER FORMATS (accounting convention)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  (1,234,567)   → -1234567   parentheses always mean negative
  -1,234,567    → -1234567   explicit minus
  Allowances, accumulated depreciation, contra-revenue → typically negative

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 3 — PERIOD MAPPING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  "Year ended December 31, 2022" / "FY2022" / "12 months ended Dec 2022" → FY22
  "Year ended December 31, 2021"                                          → FY21
  "Year ended December 31, 2020"                                          → FY20
  "Trailing twelve months" / "LTM" / "TTM" / most recent 12 months       → TTM
  Fiscal years not ending in December: use the calendar year of the end date
  Monthly columns (Jan 2023, Feb 2023 …)                                  → Jan, Feb …

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 4 — CONFIDENCE CALIBRATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  0.95–1.00  Exact value verbatim, clearly labelled, single unambiguous source
  0.80–0.94  Strong match: minor label difference, multi-doc corroboration
  0.60–0.79  Reasonable inference: value derived from adjacent data
  0.40–0.59  Estimate / extrapolation — explain basis in source_excerpt
  NEVER guess; prefer missing over fabricating a number.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 5 — SOURCE EXCERPT (mandatory for every cell)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Copy the EXACT sentence or table row from the document containing the value.
Maximum 300 characters. This enables human reviewers to verify every number.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DO NOT mark a field as MISSING if:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  • The value is legitimately zero → use raw_value: 0, confidence: 0.90
  • The value can be derived from other rows you already extracted
  • The label differs slightly but the economic meaning is identical
  • You haven't checked every chunk — search ALL provided chunks first

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CALCULATED ROWS (is_calculated: true)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  If the document states the subtotal explicitly → extract it (confidence ≥ 0.90)
  Otherwise → compute from your already-extracted component cells and set
  is_calculated: true, confidence: 0.85 in the output.
`

// ── REST call with per-attempt timeout + exponential retry ────────────────────

async function callGemini(
  prompt: string,
  signal?: AbortSignal,
  attempt = 0,
): Promise<string> {
  const MAX_ATTEMPTS = 3
  const CALL_TIMEOUT_MS = 55_000  // 55 s — leaves headroom within a 60 s section budget

  if (signal?.aborted) throw new Error('Aborted before Gemini call')

  const token = await getAccessToken()
  const url =
    `https://${LOCATION()}-aiplatform.googleapis.com/v1/projects/${PROJECT()}/locations/${LOCATION()}/publishers/google/models/${MODEL}:generateContent`

  // Combine caller abort + our per-call timeout into one signal
  const callAc = new AbortController()
  const timer = setTimeout(() => callAc.abort(), CALL_TIMEOUT_MS)
  const onParentAbort = () => callAc.abort()
  signal?.addEventListener('abort', onParentAbort, { once: true })

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: SYSTEM_INSTRUCTION }] },
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0,       // fully deterministic
          maxOutputTokens: 8192,
          responseMimeType: 'application/json',
        },
      }),
      signal: callAc.signal,
    })

    if (!res.ok) {
      const errBody = await res.text().catch(() => '')
      throw new Error(`Gemini HTTP ${res.status}: ${errBody.slice(0, 300)}`)
    }

    const data = await res.json()
    const text: string = data?.candidates?.[0]?.content?.parts?.[0]?.text ?? ''
    if (!text) throw new Error('Gemini returned an empty response')
    return text

  } catch (err) {
    clearTimeout(timer)
    signal?.removeEventListener('abort', onParentAbort)

    // Caller cancelled — propagate immediately, no retry
    if (signal?.aborted) throw err

    if (attempt < MAX_ATTEMPTS - 1) {
      const backoff = (attempt + 1) * 3000
      console.warn(`Gemini attempt ${attempt + 1} failed, retrying in ${backoff}ms:`, String(err).slice(0, 200))
      await new Promise(r => setTimeout(r, backoff))
      return callGemini(prompt, signal, attempt + 1)
    }
    throw err
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', onParentAbort)
  }
}

// ── Main extraction function ──────────────────────────────────────────────────

export async function analyzeSection(
  ragChunks: RagChunk[],
  section: SectionConfig,
  periods: string[],
  signal?: AbortSignal,
): Promise<{ cells: ExtractedCell[]; missing: MissingField[]; flags: ExtractedFlag[] }> {

  if (ragChunks.length === 0) {
    const missing: MissingField[] = section.requiredRows.flatMap(row =>
      periods.map(period => ({
        row_key: row.rowKey,
        period,
        reason: 'No relevant chunks retrieved — document may not contain this section',
        suggested_doc: 'General Ledger or Financial Statements',
      }))
    )
    return { cells: [], missing, flags: [] }
  }

  // Group chunks by source document so the model sees clear provenance
  const chunksByDoc = new Map<string, RagChunk[]>()
  for (const chunk of ragChunks) {
    const name = chunk.documentName || chunk.sourceUri?.split('/').pop() || 'Unknown document'
    if (!chunksByDoc.has(name)) chunksByDoc.set(name, [])
    chunksByDoc.get(name)!.push(chunk)
  }

  const contextText = [...chunksByDoc.entries()]
    .map(([docName, chunks]) => {
      const body = chunks
        .map((c, i) =>
          `  [Chunk ${i + 1} | Page ${c.pageNumber ?? '?'} | Relevance ${c.score.toFixed(2)}]\n  ${c.text}`)
        .join('\n\n')
      return `╔══ Document: ${docName} ══╗\n${body}`
    })
    .join('\n\n')

  const requiredRowsJson = JSON.stringify(
    section.requiredRows.map(r => ({
      row_key: r.rowKey,
      label: r.label,
      value_type: r.valueType,   // 'currency' | 'percent' | 'count' | 'text' | 'date'
      is_calculated: r.isCalculated,
    })),
    null, 2
  )

  const prompt = `\
## Extract: "${section.displayName}" section

## Required periods: ${periods.join(', ')}

## Required rows — extract every row_key × period combination:
${requiredRowsJson}

## Source documents (${ragChunks.length} chunks across ${chunksByDoc.size} file(s)):
${contextText}

## Extraction checklist:
1. Detect unit scale (thousands / millions / raw) from each document BEFORE reading values.
2. Apply parenthetical-negative rule: (x) → negative.
3. Map document period labels to the required period keys (FY20 / FY21 / FY22 / TTM).
4. Search ALL chunks before marking anything as missing.
5. For value_type "percent": output as a percentage number (e.g., 12.5 means 12.5%, not 0.125).
6. For value_type "currency": output raw integer dollars after unit-scale conversion.
7. Calculated rows (is_calculated: true) → extract stated total if present, else compute it.
8. For rank-ordered generic row keys (vendor_1…vendor_N, product_line_1…product_line_N,
   customer_1…customer_N): set display_value to "ActualName: $X,XXX,XXX" — include the
   actual vendor / product / customer name from the document so it is preserved in the UI.
9. In the "flags" array, include an entry for any cell where:
   - Confidence < 0.65 (flag_type: "low_confidence", severity: "warning")
   - Period-over-period variance >50% without an obvious explanation (flag_type: "discrepancy", severity: "warning")
   - A calculated row does not reconcile with a stated total (flag_type: "discrepancy", severity: "critical")
   - A one-time or non-recurring item is detected (flag_type: "ai_note", severity: "info")
   If none of these apply, output an empty flags array.

## Output — strict JSON, no markdown, no code fences:
{
  "cells": [
    {
      "row_key": "net_income",
      "period": "FY22",
      "raw_value": 7120305,
      "display_value": "$7,120,305",
      "source_page": 4,
      "source_excerpt": "Net income for the year ended December 31, 2022: $7,120,305",
      "source_filename": "audited_financials_2022.pdf",
      "confidence": 0.97,
      "is_calculated": false
    }
  ],
  "missing": [
    {
      "row_key": "bank_deposits",
      "period": "FY21",
      "reason": "January 2021 bank statement not present in uploaded documents",
      "suggested_doc": "Bank Statements — 2021"
    }
  ],
  "flags": [
    {
      "row_key": "net_income",
      "period": "FY21",
      "flag_type": "low_confidence",
      "severity": "warning",
      "title": "Low confidence — Net Income FY21",
      "body": "Value could not be directly verified from source; derived from adjacent data."
    }
  ]
}`

  const rawText = await callGemini(prompt, signal)

  // Strip any accidental markdown fences the model may add
  const cleaned = rawText
    .replace(/^```(?:json)?\s*/im, '')
    .replace(/\s*```\s*$/im, '')
    .trim()

  let parsed: { cells: ExtractedCell[]; missing: MissingField[]; flags: ExtractedFlag[] }
  try {
    parsed = JSON.parse(cleaned)
  } catch (parseErr) {
    const preview = cleaned.slice(0, 400)
    console.error(`[gemini] JSON parse error — section "${section.key}":`, preview)
    throw new Error(`Gemini returned unparseable JSON for section "${section.key}": ${preview}`)
  }

  // Validate + coerce every cell so the DB never receives bad types
  // Allow workbook-configured periods AND month names (for margins-by-month section)
  const MONTH_NAMES = new Set(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])
  const validPeriods = new Set([...periods, ...MONTH_NAMES])

  const cells: ExtractedCell[] = (parsed.cells ?? [])
    .filter(Boolean)
    .map(c => ({
      row_key: String(c.row_key ?? ''),
      period: String(c.period ?? ''),
      raw_value: typeof c.raw_value === 'number' && isFinite(c.raw_value)
        ? Math.round(c.raw_value * 100) / 100   // max 2 decimal places
        : null,
      display_value: String(c.display_value ?? ''),
      source_page: typeof c.source_page === 'number' ? Math.floor(c.source_page) : null,
      source_excerpt: String(c.source_excerpt ?? '').slice(0, 500),
      source_filename: String(c.source_filename ?? ''),
      confidence: Math.min(1, Math.max(0, Number(c.confidence ?? 0.5))),
      is_calculated: Boolean(c.is_calculated),
    }))
    .filter(c => c.row_key && c.period && validPeriods.has(c.period))

  const missing: MissingField[] = (parsed.missing ?? [])
    .filter(Boolean)
    .map(m => ({
      row_key: String(m.row_key ?? ''),
      period: String(m.period ?? ''),
      reason: String(m.reason ?? ''),
      suggested_doc: String(m.suggested_doc ?? ''),
    }))
    .filter(m => m.row_key && m.period && validPeriods.has(m.period))

  const VALID_FLAG_TYPES = new Set(['low_confidence', 'discrepancy', 'needs_review', 'ai_note'])
  const VALID_SEVERITIES = new Set(['info', 'warning', 'critical'])

  const flags: ExtractedFlag[] = (parsed.flags ?? [])
    .filter(Boolean)
    .map(f => ({
      row_key: String(f.row_key ?? ''),
      period: String(f.period ?? ''),
      flag_type: VALID_FLAG_TYPES.has(f.flag_type) ? f.flag_type : 'needs_review',
      severity: VALID_SEVERITIES.has(f.severity) ? f.severity : 'warning',
      title: String(f.title ?? '').slice(0, 200),
      body: String(f.body ?? '').slice(0, 1000),
    }))
    .filter(f => f.row_key && f.period && f.title && validPeriods.has(f.period))

  return { cells, missing, flags }
}

// ── Workbook AI Chat ───────────────────────────────────────────────────────────

export interface ChatMessage {
  role: 'user' | 'model'
  content: string
}

export interface WorkbookContext {
  companyName: string
  periods: string[]
  cells: Array<{
    section: string
    row_key: string
    period: string
    raw_value: number | null
    display_value: string | null
    confidence: number | null
  }>
  flags: Array<{
    id: string
    section: string
    row_key: string
    period: string | null
    flag_type: string
    severity: string
    title: string
    body: string | null
    created_by_ai: boolean
    resolved_at: string | null
  }>
}

function buildWorkbookSystemPrompt(ctx: WorkbookContext): string {
  // Organise cells by section → period table
  const sections = new Map<string, Map<string, Map<string, string>>>()
  for (const c of ctx.cells) {
    if (!sections.has(c.section)) sections.set(c.section, new Map())
    const sec = sections.get(c.section)!
    if (!sec.has(c.row_key)) sec.set(c.row_key, new Map())
    sec.get(c.row_key)!.set(c.period, c.display_value ?? (c.raw_value != null ? String(c.raw_value) : '—'))
  }

  const periodList = ctx.periods.join(', ')

  let financialSummary = ''
  for (const [sectionKey, rows] of sections) {
    financialSummary += `\n### ${sectionKey.toUpperCase().replace(/-/g, ' ')}\n`
    for (const [rowKey, periods] of rows) {
      const vals = ctx.periods.map(p => `${p}: ${periods.get(p) ?? '—'}`).join('  |  ')
      financialSummary += `  ${rowKey.replace(/_/g, ' ')}: ${vals}\n`
    }
  }

  const openFlags = ctx.flags.filter(f => !f.resolved_at)
  const flagsSummary = openFlags.length === 0
    ? 'No open flags.'
    : openFlags.map((f, i) =>
        `${i + 1}. [${f.severity.toUpperCase()}] ${f.title}` +
        (f.section ? ` — ${f.section}` : '') +
        (f.period ? ` (${f.period})` : '') +
        (f.body ? `\n   ${f.body}` : '')
      ).join('\n')

  return `You are Workbook AI, an expert Quality of Earnings (QoE) analyst embedded inside a professional financial diligence platform.

COMPANY: ${ctx.companyName}
PERIODS ANALYSED: ${periodList}

─── FINANCIAL DATA ───────────────────────────────────────────────
${financialSummary}

─── OPEN FLAGS (${openFlags.length}) ────────────────────────────────────────
${flagsSummary}

─── YOUR ROLE ────────────────────────────────────────────────────
You have deep knowledge of this workbook. Your responsibilities:
1. Answer precise questions about any metric — cite section, period, and value
2. Walk the analyst through open flags one by one, explain the issue, and suggest resolution steps
3. Spot discrepancies, unusual trends, or missing data across sections
4. Explain QoE adjustments, normalisation logic, and their impact on EBITDA
5. When asked to "walk me through" or "review" the workbook, go flag by flag with analysis
6. Keep responses concise, professional, and actionable — bullet points preferred for lists
7. If a specific metric is referenced, lead with that metric's data before broader context
8. Never fabricate numbers — only cite values present in the financial data above`
}

export async function chatWithWorkbook(
  ctx: WorkbookContext,
  history: ChatMessage[],
  userMessage: string,
  cellRef?: { label: string; period: string; displayValue: string },
): Promise<string> {
  const token = await getAccessToken()
  const url =
    `https://${LOCATION()}-aiplatform.googleapis.com/v1/projects/${PROJECT()}/locations/${LOCATION()}/publishers/google/models/${MODEL}:generateContent`

  // Prepend cell reference context to the user message if provided
  const messageWithContext = cellRef
    ? `[Referencing metric: ${cellRef.label} | Period: ${cellRef.period} | Value: ${cellRef.displayValue}]\n\n${userMessage}`
    : userMessage

  const contents = [
    ...history.map(m => ({
      role: m.role,
      parts: [{ text: m.content }],
    })),
    { role: 'user', parts: [{ text: messageWithContext }] },
  ]

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: buildWorkbookSystemPrompt(ctx) }] },
      contents,
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 2048,
      },
    }),
  })

  if (!res.ok) {
    const errBody = await res.text().catch(() => '')
    throw new Error(`Gemini chat HTTP ${res.status}: ${errBody.slice(0, 300)}`)
  }

  const data = await res.json()
  const text: string = data?.candidates?.[0]?.content?.parts?.[0]?.text ?? ''
  if (!text) throw new Error('Gemini returned an empty chat response')
  return text
}
