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
import type { SectionConfig, RowDef } from './section-prompts'

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

// ── System instruction ────────────────────────────────────────────────────────

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
Keep it under 100 characters. This enables human reviewers to verify numbers.

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
  const CALL_TIMEOUT_MS = 55_000

  if (signal?.aborted) throw new Error('Aborted before Gemini call')

  const token = await getAccessToken()
  const url =
    `https://${LOCATION()}-aiplatform.googleapis.com/v1/projects/${PROJECT()}/locations/${LOCATION()}/publishers/google/models/${MODEL}:generateContent`

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
          temperature: 0,
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

// ── Shared context builder ────────────────────────────────────────────────────

function buildContext(ragChunks: RagChunk[]): { contextText: string; docCount: number } {
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

  return { contextText, docCount: chunksByDoc.size }
}

// ── JSON coercion helpers ─────────────────────────────────────────────────────

const MONTH_NAMES = new Set(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])
const VALID_FLAG_TYPES = new Set(['low_confidence', 'discrepancy', 'needs_review', 'ai_note'])
const VALID_SEVERITIES = new Set(['info', 'warning', 'critical'])

function coerceCells(raw: unknown[], validPeriods: Set<string>): ExtractedCell[] {
  return (raw ?? [])
    .filter(Boolean)
    .map((c: unknown) => {
      const cell = c as Record<string, unknown>
      return {
        row_key: String(cell.row_key ?? ''),
        period: String(cell.period ?? ''),
        raw_value: typeof cell.raw_value === 'number' && isFinite(cell.raw_value as number)
          ? Math.round((cell.raw_value as number) * 100) / 100
          : null,
        display_value: String(cell.display_value ?? ''),
        source_page: typeof cell.source_page === 'number' ? Math.floor(cell.source_page as number) : null,
        source_excerpt: String(cell.source_excerpt ?? '').slice(0, 200),
        source_filename: String(cell.source_filename ?? ''),
        confidence: Math.min(1, Math.max(0, Number(cell.confidence ?? 0.5))),
        is_calculated: Boolean(cell.is_calculated),
      }
    })
    .filter(c => c.row_key && c.period && validPeriods.has(c.period))
}

function coerceMissing(raw: unknown[], validPeriods: Set<string>): MissingField[] {
  return (raw ?? [])
    .filter(Boolean)
    .map((m: unknown) => {
      const miss = m as Record<string, unknown>
      return {
        row_key: String(miss.row_key ?? ''),
        period: String(miss.period ?? ''),
        reason: String(miss.reason ?? ''),
        suggested_doc: String(miss.suggested_doc ?? ''),
      }
    })
    .filter(m => m.row_key && m.period && validPeriods.has(m.period))
}

function coerceFlags(raw: unknown[], validPeriods: Set<string>): ExtractedFlag[] {
  return (raw ?? [])
    .filter(Boolean)
    .map((f: unknown) => {
      const flag = f as Record<string, unknown>
      return {
        row_key: String(flag.row_key ?? ''),
        period: String(flag.period ?? ''),
        flag_type: VALID_FLAG_TYPES.has(flag.flag_type as string)
          ? (flag.flag_type as ExtractedFlag['flag_type'])
          : 'needs_review',
        severity: VALID_SEVERITIES.has(flag.severity as string)
          ? (flag.severity as ExtractedFlag['severity'])
          : 'warning',
        title: String(flag.title ?? '').slice(0, 200),
        body: String(flag.body ?? '').slice(0, 1000),
      }
    })
    .filter(f => f.row_key && f.period && f.title && validPeriods.has(f.period))
}

// ── Single batch extraction call ──────────────────────────────────────────────

async function extractBatch(
  contextText: string,
  docCount: number,
  chunkCount: number,
  sectionName: string,
  rows: RowDef[],
  periods: string[],
  validPeriods: Set<string>,
  signal?: AbortSignal,
): Promise<{ cells: ExtractedCell[]; missing: MissingField[]; flags: ExtractedFlag[] }> {

  const requiredRowsJson = JSON.stringify(
    rows.map(r => ({
      row_key: r.rowKey,
      label: r.label,
      value_type: r.valueType,
      is_calculated: r.isCalculated,
    })),
    null, 2
  )

  const prompt = `\
## Extract: "${sectionName}" section

## Required periods: ${periods.join(', ')}

## Required rows — extract every row_key × period combination:
${requiredRowsJson}

## Source documents (${chunkCount} chunks across ${docCount} file(s)):
${contextText}

## Extraction checklist:
1. Detect unit scale (thousands / millions / raw) from each document BEFORE reading values.
2. Apply parenthetical-negative rule: (x) → negative.
3. Map document period labels to the required period keys (${periods.join(' / ')}).
4. Search ALL chunks before marking anything as missing.
5. For value_type "percent": output as a decimal percentage (e.g., 12.5 means 12.5%, not 0.125).
6. For value_type "currency": output raw integer dollars after unit-scale conversion.
7. Calculated rows (is_calculated: true) → extract stated total if present, else compute from components.
8. For rank-ordered generic row keys (vendor_1…N, product_line_1…N, customer_1…N):
   set display_value to "ActualName: $X,XXX,XXX" — include the real name from the document.
9. Keep source_excerpt under 100 characters.
10. In the "flags" array, include an entry for:
    - Confidence < 0.65 → flag_type: "low_confidence", severity: "warning"
    - Period-over-period variance >50% unexplained → flag_type: "discrepancy", severity: "warning"
    - Calculated row does not reconcile with stated total → flag_type: "discrepancy", severity: "critical"
    - One-time or non-recurring item detected → flag_type: "ai_note", severity: "info"

## Output — strict JSON only, no markdown, no code fences:
{
  "cells": [
    {
      "row_key": "net_income",
      "period": "FY22",
      "raw_value": 7120305,
      "display_value": "$7,120,305",
      "source_page": 4,
      "source_excerpt": "Net income FY2022: $7,120,305",
      "source_filename": "audited_financials_2022.pdf",
      "confidence": 0.97,
      "is_calculated": false
    }
  ],
  "missing": [
    {
      "row_key": "bank_deposits",
      "period": "FY21",
      "reason": "2021 bank statements not in uploaded documents",
      "suggested_doc": "Bank Statements 2021"
    }
  ],
  "flags": []
}`

  const rawText = await callGemini(prompt, signal)
  const cleaned = rawText
    .replace(/^```(?:json)?\s*/im, '')
    .replace(/\s*```\s*$/im, '')
    .trim()

  let parsed: { cells: unknown[]; missing: unknown[]; flags: unknown[] }
  try {
    parsed = JSON.parse(cleaned)
  } catch {
    const preview = cleaned.slice(0, 400)
    console.error(`[gemini] JSON parse error — section "${sectionName}":`, preview)
    throw new Error(`Gemini returned unparseable JSON for "${sectionName}": ${preview}`)
  }

  return {
    cells: coerceCells(parsed.cells ?? [], validPeriods),
    missing: coerceMissing(parsed.missing ?? [], validPeriods),
    flags: coerceFlags(parsed.flags ?? [], validPeriods),
  }
}

// ── Main extraction function (with row batching for large sections) ────────────

/**
 * Sections with many rows (e.g., Income Statement: 29 rows × 4 periods = 116 cells)
 * can exceed Gemini's 8192-token output limit. We split into batches of ROW_BATCH_SIZE
 * rows, run them in parallel (same RAG context), then merge.
 */
const ROW_BATCH_SIZE = 12

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

  const validPeriods = new Set([...periods, ...MONTH_NAMES])
  const { contextText, docCount } = buildContext(ragChunks)

  const rows = section.requiredRows

  if (rows.length <= ROW_BATCH_SIZE) {
    // Single call path
    return extractBatch(
      contextText, docCount, ragChunks.length,
      section.displayName, rows, periods, validPeriods, signal,
    )
  }

  // Multi-batch path — split rows into chunks, run in parallel
  const batches: RowDef[][] = []
  for (let i = 0; i < rows.length; i += ROW_BATCH_SIZE) {
    batches.push(rows.slice(i, i + ROW_BATCH_SIZE))
  }

  const results = await Promise.allSettled(
    batches.map((batchRows, idx) =>
      extractBatch(
        contextText, docCount, ragChunks.length,
        `${section.displayName} (part ${idx + 1}/${batches.length})`,
        batchRows, periods, validPeriods, signal,
      )
    )
  )

  const merged: { cells: ExtractedCell[]; missing: MissingField[]; flags: ExtractedFlag[] } = {
    cells: [], missing: [], flags: [],
  }
  for (const result of results) {
    if (result.status === 'fulfilled') {
      merged.cells.push(...result.value.cells)
      merged.missing.push(...result.value.missing)
      merged.flags.push(...result.value.flags)
    } else {
      console.error('[gemini] batch failed:', result.reason)
    }
  }
  return merged
}

// ── Formula-based post-processing ─────────────────────────────────────────────

/**
 * Evaluate a simple arithmetic formula (no parentheses, linear operations).
 * Formula syntax: "rowKey1+rowKey2-rowKey3" etc.
 * Returns null if any referenced row is missing for this period.
 */
function evalFormula(formula: string, valueMap: Map<string, number | null>): number | null {
  // Match operators and row-key tokens (letters, digits, underscores)
  const tokens = formula.match(/[+\-*/]|[\w]+/g) ?? []
  let result: number | null = null
  let op = '+'

  for (const token of tokens) {
    if (token === '+' || token === '-' || token === '*' || token === '/') {
      op = token
      continue
    }
    const val = valueMap.get(token) ?? null
    if (val === null) return null // missing component — cannot compute

    if (result === null) {
      result = val
    } else {
      switch (op) {
        case '+': result += val; break
        case '-': result -= val; break
        case '*': result *= val; break
        case '/': result = val !== 0 ? result / val : null; break
      }
    }
  }
  return result
}

function formatForValueType(value: number, valueType: string): string {
  if (valueType === 'percent') return `${value.toFixed(1)}%`
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 0,
  }).format(value)
}

/**
 * Apply formula definitions to fill in any calculated rows that Gemini missed.
 * Runs up to 3 passes to handle chains (A → B → C).
 * Existing cells are never overwritten — this only adds missing ones.
 */
export function applyFormulas(
  cells: ExtractedCell[],
  rows: RowDef[],
  periods: string[],
): ExtractedCell[] {
  const allCells = [...cells]

  for (let pass = 0; pass < 3; pass++) {
    // Rebuild lookup maps each pass (newly added cells become available)
    const existingKeys = new Set(allCells.map(c => `${c.row_key}:${c.period}`))
    const valueMap = new Map<string, number | null>()
    for (const c of allCells) {
      if (c.raw_value !== null) valueMap.set(`${c.row_key}:${c.period}`, c.raw_value)
    }

    let addedAny = false

    for (const row of rows) {
      if (!row.formula || !row.isCalculated) continue

      for (const period of periods) {
        const key = `${row.rowKey}:${period}`
        if (existingKeys.has(key)) continue

        // Build a period-scoped value map for formula evaluation
        const periodMap = new Map<string, number | null>()
        for (const r of rows) {
          const v = valueMap.get(`${r.rowKey}:${period}`)
          if (v !== undefined) periodMap.set(r.rowKey, v)
        }

        const computed = evalFormula(row.formula, periodMap)
        if (computed === null) continue // components not yet available

        allCells.push({
          row_key: row.rowKey,
          period,
          raw_value: Math.round(computed * 100) / 100,
          display_value: formatForValueType(computed, row.valueType),
          source_page: null,
          source_excerpt: `Computed: ${row.formula}`,
          source_filename: '',
          confidence: 0.90,
          is_calculated: true,
        })
        addedAny = true
      }
    }

    if (!addedAny) break // converged — no new cells this pass
  }

  return allCells
}

// ── Cross-section reconciliation ──────────────────────────────────────────────

interface SectionCell {
  section: string
  row_key: string
  period: string
  raw_value: number | null
}

interface ReconciliationCheck {
  section1: string; key1: string
  section2: string; key2: string
  /** Relative tolerance (e.g. 0.02 = 2%) */
  tolerance: number
  label: string
  severity: ExtractedFlag['severity']
}

const RECONCILIATION_CHECKS: ReconciliationCheck[] = [
  {
    section1: 'overview', key1: 'total_revenue',
    section2: 'income-statement', key2: 'total_net_revenue',
    tolerance: 0.02, label: 'Total Revenue', severity: 'warning',
  },
  {
    section1: 'qoe', key1: 'net_income',
    section2: 'income-statement', key2: 'net_income',
    tolerance: 0.01, label: 'Net Income', severity: 'critical',
  },
  {
    section1: 'qoe', key1: 'ebitda_as_defined',
    section2: 'overview', key2: 'ebitda_as_defined',
    tolerance: 0.01, label: 'EBITDA as Defined', severity: 'warning',
  },
  {
    section1: 'cogs-vendors', key1: 'total_cogs',
    section2: 'income-statement', key2: 'total_cogs',
    tolerance: 0.02, label: 'Total COGS', severity: 'warning',
  },
  {
    section1: 'customer-concentration', key1: 'total_revenue',
    section2: 'overview', key2: 'total_revenue',
    tolerance: 0.02, label: 'Customer Concentration vs. Revenue', severity: 'warning',
  },
  {
    section1: 'proof-revenue', key1: 'gl_revenue',
    section2: 'income-statement', key2: 'total_net_revenue',
    tolerance: 0.05, label: 'GL Revenue vs. P&L Revenue', severity: 'warning',
  },
  {
    // Balance sheet identity: total assets must equal total liabilities + equity
    section1: 'balance-sheet', key1: 'total_assets',
    section2: 'balance-sheet', key2: 'total_liabilities_equity',
    tolerance: 0.005, label: 'Balance Sheet (Assets = L+E)', severity: 'critical',
  },
]

/**
 * Compare key metrics across sections and return discrepancy flags.
 * Called once after all sections have been analysed.
 */
export function buildReconciliationFlags(
  allCells: SectionCell[],
  periods: string[],
): ExtractedFlag[] {
  const flags: ExtractedFlag[] = []

  // Build lookup: section:rowKey:period → raw_value
  const lookup = new Map<string, number | null>()
  for (const c of allCells) {
    lookup.set(`${c.section}:${c.row_key}:${c.period}`, c.raw_value)
  }

  for (const check of RECONCILIATION_CHECKS) {
    for (const period of periods) {
      const v1 = lookup.get(`${check.section1}:${check.key1}:${period}`)
      const v2 = lookup.get(`${check.section2}:${check.key2}:${period}`)

      if (v1 == null || v2 == null) continue // can't compare if either is missing

      const avg = (Math.abs(v1) + Math.abs(v2)) / 2
      if (avg === 0) continue

      const relDiff = Math.abs(v1 - v2) / avg

      if (relDiff > check.tolerance) {
        const fmt = (n: number) => new Intl.NumberFormat('en-US', {
          style: 'currency', currency: 'USD', minimumFractionDigits: 0,
        }).format(n)

        flags.push({
          row_key: check.key1,
          period,
          flag_type: 'discrepancy',
          severity: check.severity,
          title: `Cross-section discrepancy: ${check.label} (${period})`,
          body: `${check.section1}.${check.key1} = ${fmt(v1)} vs ${check.section2}.${check.key2} = ${fmt(v2)} — ${(relDiff * 100).toFixed(1)}% difference (tolerance ${(check.tolerance * 100).toFixed(0)}%).`,
        })
      }
    }
  }

  return flags
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
