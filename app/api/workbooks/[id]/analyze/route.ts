import { NextRequest } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { createRagCorpus, queryRagCorpus, importRagFile } from '@/lib/vertex-rag'
import { analyzeSection } from '@/lib/gemini'
import { SECTION_CONFIGS } from '@/lib/section-prompts'
import type { SectionConfig } from '@/lib/section-prompts'
import type { MissingField } from '@/lib/gemini'
import type { ExtractedFlag } from '@/lib/gemini'

export const dynamic = 'force-dynamic'
/**
 * 10 minutes — covers first-run corpus creation (60-120 s) + pre-flight imports
 * (up to 90 s) + parallel section analysis (4 batches × ~65 s = 260 s).
 * Subsequent runs skip corpus creation and complete in ~4 min.
 */
export const maxDuration = 600

/**
 * Process sections in parallel batches so all 13 sections fit well within budget.
 * BATCH_SIZE=4: ceil(13/4) = 4 batches × ~65 s average = ~260 s for analysis.
 * Per-section budget: 85 s (30 s RAG + 50 s Gemini + 5 s margin).
 */
const BATCH_SIZE = 4
const SECTION_TIMEOUT_MS = 85_000

function sseEvent(data: unknown): string {
  return `data: ${JSON.stringify(data)}\n\n`
}

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()

  if (authError || !user) {
    return new Response('Unauthorized', { status: 401 })
  }

  const serviceClient = createSupabaseServiceClient()

  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, periods, created_by')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return new Response('Workbook not found', { status: 404 })

  const periods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']

  const stream = new ReadableStream({
    async start(controller) {
      const encoder = new TextEncoder()

      const emit = (data: unknown) => {
        try {
          controller.enqueue(encoder.encode(sseEvent(data)))
        } catch {
          // client disconnected — analysis continues server-side
        }
      }

      // Keepalive every 15 s to prevent proxy/browser SSE timeout
      const keepalive = setInterval(() => {
        try {
          controller.enqueue(encoder.encode(': keepalive\n\n'))
        } catch {
          clearInterval(keepalive)
        }
      }, 15_000)

      try {
        // ── Step 1: Ensure RAG corpus exists ──────────────────────────────────
        // If corpus doesn't exist yet (background task hasn't completed or failed),
        // create it inline here so the user always gets real-time progress and
        // errors are surfaced immediately via SSE instead of a silent workbook failure.
        let { data: corpus } = await serviceClient
          .from('rag_corpora')
          .select('corpus_name')
          .eq('workbook_id', workbookId)
          .single()

        if (!corpus?.corpus_name) {
          emit({ type: 'status', message: 'Creating AI workspace (RAG corpus)…' })
          await serviceClient
            .from('workbooks')
            .update({ status: 'analyzing' })
            .eq('id', workbookId)

          try {
            const corpusName = await createRagCorpus(workbookId)
            await serviceClient
              .from('rag_corpora')
              .insert({ workbook_id: workbookId, corpus_name: corpusName })
            corpus = { corpus_name: corpusName }
            emit({ type: 'status', message: 'AI workspace ready.' })
          } catch (corpusErr) {
            const msg = corpusErr instanceof Error ? corpusErr.message : String(corpusErr)
            emit({ type: 'error', error: `Failed to create RAG corpus: ${msg}` })
            await serviceClient
              .from('workbooks')
              .update({ status: 'error' })
              .eq('id', workbookId)
            return
          }
        } else {
          await serviceClient
            .from('workbooks')
            .update({ status: 'analyzing' })
            .eq('id', workbookId)
        }

        const corpusName = corpus.corpus_name

        // ── Step 2: Pre-flight — import docs with gcs_uri but no rag_file_id ──
        // Catches files uploaded while corpus was being created, or whose
        // background RAG import failed. 90 s hard cap so this never starves sections.
        {
          const { data: unimported } = await serviceClient
            .from('documents')
            .select('id, gcs_uri, file_name')
            .eq('workbook_id', workbookId)
            .eq('ingestion_status', 'ready')
            .is('rag_file_id', null)
            .not('gcs_uri', 'is', null)

          if (unimported && unimported.length > 0) {
            emit({ type: 'status', message: `Importing ${unimported.length} document(s) into RAG corpus…` })

            const preflightAc = new AbortController()
            const preflightTimer = setTimeout(() => {
              console.warn('Pre-flight RAG import budget exhausted (90 s) — proceeding to analysis')
              preflightAc.abort()
            }, 90_000)

            try {
              for (const doc of unimported) {
                if (preflightAc.signal.aborted) break
                try {
                  const ragFileId = await importRagFile(
                    corpusName,
                    doc.gcs_uri as string,
                    preflightAc.signal,
                  )
                  await serviceClient
                    .from('documents')
                    .update({ rag_file_id: ragFileId })
                    .eq('id', doc.id)
                } catch (err) {
                  if (!preflightAc.signal.aborted) {
                    console.error(`Pre-flight RAG import failed for ${doc.file_name}:`, err)
                  }
                  // Continue — analysis still works with other docs
                }
              }
            } finally {
              clearTimeout(preflightTimer)
            }
          }
        }

        // Fetch document list once (for source_document_id lookup when persisting cells)
        const { data: docs } = await serviceClient
          .from('documents')
          .select('id, file_name')
          .eq('workbook_id', workbookId)
        const docByName = new Map(docs?.map((d) => [d.file_name, d.id]) ?? [])

        /**
         * Run one section: resume check → RAG → Gemini → DB upsert.
         * Returns missing fields for final roll-up. Never throws.
         */
        const runSection = async (section: SectionConfig): Promise<{ missing: MissingField[]; flags: ExtractedFlag[] }> => {
          // Resume: skip sections already populated from a prior interrupted run
          const { count: existingCount } = await serviceClient
            .from('workbook_cells')
            .select('id', { count: 'exact', head: true })
            .eq('workbook_id', workbookId)
            .eq('section', section.key)

          if (existingCount && existingCount > 0) {
            emit({
              type: 'section_complete',
              section: section.key,
              displayName: section.displayName,
              cells_extracted: existingCount,
              missing_count: 0,
              resumed: true,
            })
            return { missing: [], flags: [] }
          }

          emit({ type: 'section_start', section: section.key, displayName: section.displayName })

          const sectionAc = new AbortController()
          const sectionTimer = setTimeout(() => sectionAc.abort(), SECTION_TIMEOUT_MS)

          try {
            // RAG retrieval
            let ragChunks
            try {
              ragChunks = await queryRagCorpus(
                corpusName,
                section.ragQuery,
                25,
                sectionAc.signal,
              )
            } catch (err) {
              emit({
                type: 'section_error',
                section: section.key,
                displayName: section.displayName,
                error: String(err),
              })
              return { missing: [], flags: [] }
            }

            // Gemini extraction
            let result
            try {
              result = await analyzeSection(ragChunks, section, periods, sectionAc.signal)
            } catch (err) {
              emit({
                type: 'section_error',
                section: section.key,
                displayName: section.displayName,
                error: String(err),
              })
              return { missing: [], flags: [] }
            }

            // Persist cells
            if (result.cells.length > 0) {
              const cellRows = result.cells.map((cell) => ({
                workbook_id: workbookId,
                section: section.key,
                row_key: cell.row_key,
                period: cell.period,
                raw_value: cell.raw_value,
                display_value: cell.display_value,
                is_calculated: cell.is_calculated,
                source_document_id: cell.source_filename
                  ? (docByName.get(cell.source_filename) ?? null)
                  : null,
                source_page: cell.source_page,
                source_excerpt: cell.source_excerpt,
                confidence: cell.confidence,
                is_overridden: false,
              }))

              await serviceClient
                .from('workbook_cells')
                .upsert(cellRows, { onConflict: 'workbook_id,section,row_key,period' })
            }

            // Persist missing-data requests
            if (result.missing.length > 0) {
              await serviceClient
                .from('missing_data_requests')
                .delete()
                .eq('workbook_id', workbookId)
                .eq('section', section.key)
                .eq('status', 'open')

              const missingRows = result.missing.map((m) => ({
                workbook_id: workbookId,
                section: section.key,
                field_key: m.row_key,
                period: m.period,
                reason: m.reason,
                suggested_doc: m.suggested_doc,
                status: 'open',
              }))

              await serviceClient.from('missing_data_requests').insert(missingRows)

              emit({
                type: 'section_missing',
                section: section.key,
                missing: result.missing,
              })
            }

            // Persist AI flags (merge AI prompt flags + auto low-confidence flags)
            const flagsToInsert = [...result.flags]

            // Auto-add low-confidence flags for cells with confidence < 0.65 not already flagged
            const flaggedKeys = new Set(result.flags.map(f => `${f.row_key}:${f.period}`))
            for (const cell of result.cells) {
              if (cell.confidence < 0.65 && !flaggedKeys.has(`${cell.row_key}:${cell.period}`)) {
                flagsToInsert.push({
                  row_key: cell.row_key,
                  period: cell.period,
                  flag_type: 'low_confidence',
                  severity: 'warning',
                  title: `Low confidence — ${cell.row_key} ${cell.period}`,
                  body: `Confidence score: ${Math.round(cell.confidence * 100)}%. Value may need manual verification.`,
                })
              }
            }

            if (flagsToInsert.length > 0) {
              // Look up cell_ids by (workbook_id, section, row_key, period)
              const { data: savedCells } = await serviceClient
                .from('workbook_cells')
                .select('id, row_key, period')
                .eq('workbook_id', workbookId)
                .eq('section', section.key)

              const cellIdMap = new Map(
                (savedCells ?? []).map(c => [`${c.row_key}:${c.period}`, c.id])
              )

              const flagRows = flagsToInsert.map(f => ({
                workbook_id: workbookId,
                cell_id: cellIdMap.get(`${f.row_key}:${f.period}`) ?? null,
                section: section.key,
                row_key: f.row_key,
                period: f.period,
                flag_type: f.flag_type,
                severity: f.severity,
                title: f.title,
                body: f.body,
                created_by_ai: true,
              }))

              const { data: insertedFlags } = await serviceClient
                .from('cell_flags')
                .insert(flagRows)
                .select('id, body')

              // Add initial AI comment for each flag
              if (insertedFlags && insertedFlags.length > 0) {
                const commentRows = insertedFlags
                  .filter(f => f.body)
                  .map(f => ({
                    flag_id: f.id,
                    workbook_id: workbookId,
                    author_id: null,
                    author_name: 'AI',
                    body: f.body,
                  }))
                if (commentRows.length > 0) {
                  await serviceClient.from('flag_comments').insert(commentRows)
                }
              }
            }

            emit({
              type: 'section_complete',
              section: section.key,
              displayName: section.displayName,
              cells_extracted: result.cells.length,
              missing_count: result.missing.length,
            })

            return { missing: result.missing, flags: flagsToInsert }

          } finally {
            clearTimeout(sectionTimer)
          }
        }

        // ── Step 3: Process sections in parallel batches ──────────────────────
        const allMissing: MissingField[] = []

        for (let i = 0; i < SECTION_CONFIGS.length; i += BATCH_SIZE) {
          const batch = SECTION_CONFIGS.slice(i, i + BATCH_SIZE)
          const settled = await Promise.allSettled(batch.map(runSection))
          for (const outcome of settled) {
            if (outcome.status === 'fulfilled') {
              allMissing.push(...outcome.value.missing)
            }
          }
        }

        // ── Final workbook status ─────────────────────────────────────────────
        const finalStatus = allMissing.length > 0 ? 'needs_input' : 'ready'
        await serviceClient
          .from('workbooks')
          .update({ status: finalStatus })
          .eq('id', workbookId)

        emit({ type: 'complete', status: finalStatus, total_missing: allMissing.length })

      } catch (err) {
        emit({ type: 'error', error: String(err) })
        await serviceClient
          .from('workbooks')
          .update({ status: 'error' })
          .eq('id', workbookId)
      } finally {
        clearInterval(keepalive)
        controller.close()
      }
    },
  })

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'X-Accel-Buffering': 'no',
    },
  })
}
