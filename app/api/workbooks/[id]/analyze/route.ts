import { NextRequest } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import { createRagCorpus, queryRagCorpusMulti, importRagFile } from '@/lib/vertex-rag'
import { analyzeSection, applyFormulas, buildReconciliationFlags } from '@/lib/gemini'
import { SECTION_CONFIGS } from '@/lib/section-prompts'
import type { SectionConfig } from '@/lib/section-prompts'
import type { MissingField, ExtractedFlag } from '@/lib/gemini'

export const dynamic = 'force-dynamic'
/**
 * 10 minutes — covers:
 *  - First-run corpus creation (60–120 s)
 *  - Pre-flight doc imports (up to 120 s)
 *  - Section analysis: 4 batches × ~90 s = ~360 s
 *  - Formula post-processing + reconciliation (< 5 s)
 */
export const maxDuration = 600

/**
 * Run 4 sections in parallel per batch.
 * Each section budget: 120 s (multi-query RAG ≤ 30 s parallel + Gemini batches ≤ 80 s).
 */
const BATCH_SIZE = 4
const SECTION_TIMEOUT_MS = 120_000

// ── Low-confidence threshold ───────────────────────────────────────────────────
const LOW_CONFIDENCE_THRESHOLD = 0.65

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

  const workbookPeriods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']

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
        try { controller.enqueue(encoder.encode(': keepalive\n\n')) }
        catch { clearInterval(keepalive) }
      }, 15_000)

      try {
        // ── Step 0: Gate — require at least one uploaded document ─────────────
        const { data: readyDocs, count: readyCount } = await serviceClient
          .from('documents')
          .select('id', { count: 'exact', head: true })
          .eq('workbook_id', workbookId)
          .in('ingestion_status', ['ready', 'ingesting'])

        if (!readyCount || readyCount === 0) {
          emit({
            type: 'error',
            error: 'No documents uploaded. Please upload at least one financial document (e.g. P&L, general ledger, or bank statements) before running analysis.',
          })
          await serviceClient.from('workbooks').update({ status: 'needs_input' }).eq('id', workbookId)
          clearInterval(keepalive)
          controller.close()
          return
        }

        // ── Step 1: Ensure RAG corpus exists ──────────────────────────────────
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

        // ── Step 2: Pre-flight — import any docs not yet in RAG ───────────────
        // Catches docs that:
        //  (a) uploaded while corpus was being created (no rag_file_id, status='ready')
        //  (b) got stuck mid-import — gcs_uri set but status='ingesting', rag_file_id null
        {
          const { data: unimported } = await serviceClient
            .from('documents')
            .select('id, gcs_uri, file_name')
            .eq('workbook_id', workbookId)
            .in('ingestion_status', ['ready', 'ingesting'])
            .is('rag_file_id', null)
            .not('gcs_uri', 'is', null)

          if (unimported && unimported.length > 0) {
            emit({
              type: 'status',
              message: `Importing ${unimported.length} document(s) into AI workspace…`,
            })

            const preflightAc = new AbortController()
            // 120 s hard cap — large PDFs can take up to 90 s each
            const preflightTimer = setTimeout(() => {
              console.warn('Pre-flight RAG import budget exhausted (120 s) — proceeding to analysis')
              preflightAc.abort()
            }, 120_000)

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
                    .update({ rag_file_id: ragFileId, ingestion_status: 'ready' })
                    .eq('id', doc.id)
                  emit({ type: 'status', message: `Imported: ${doc.file_name}` })
                } catch (err) {
                  if (!preflightAc.signal.aborted) {
                    console.error(`Pre-flight RAG import failed for ${doc.file_name}:`, err)
                    // Mark error so it doesn't keep retrying silently
                    await serviceClient
                      .from('documents')
                      .update({ ingestion_status: 'error' })
                      .eq('id', doc.id)
                      .is('rag_file_id', null) // only if still unimported
                  }
                }
              }
            } finally {
              clearTimeout(preflightTimer)
            }
          }
        }

        // Fetch doc list for source_document_id lookup — try exact name and lowercase match
        const { data: docs } = await serviceClient
          .from('documents')
          .select('id, file_name')
          .eq('workbook_id', workbookId)

        // Build both exact and lowercase maps for fuzzy filename matching
        const docByNameExact = new Map((docs ?? []).map(d => [d.file_name, d.id]))
        const docByNameLower = new Map((docs ?? []).map(d => [d.file_name.toLowerCase(), d.id]))

        function resolveDocId(filename: string): string | null {
          if (!filename) return null
          // Try exact match first, then case-insensitive, then basename only
          return (
            docByNameExact.get(filename) ??
            docByNameLower.get(filename.toLowerCase()) ??
            docByNameLower.get(filename.split('/').pop()?.toLowerCase() ?? '') ??
            null
          )
        }

        // Accumulate all extracted cells for cross-section reconciliation
        const allExtractedCells: Array<{
          section: string; row_key: string; period: string; raw_value: number | null
        }> = []

        /**
         * Run one section: resume check → multi-query RAG → Gemini → formulas → DB upsert.
         * Uses the section's overridePeriods if set (e.g. months for margins-by-month).
         * Returns missing fields for final roll-up. Never throws.
         */
        const runSection = async (
          section: SectionConfig,
        ): Promise<{ missing: MissingField[]; flags: ExtractedFlag[] }> => {
          // Resume: skip sections already fully populated from a prior interrupted run
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

          // Use overridePeriods if defined (e.g. month names for margins-by-month)
          const sectionPeriods = section.overridePeriods ?? workbookPeriods

          const sectionAc = new AbortController()
          const sectionTimer = setTimeout(() => sectionAc.abort(), SECTION_TIMEOUT_MS)

          try {
            // ── Multi-query RAG retrieval ────────────────────────────────────
            // Use ragQueries (multiple targeted queries) for higher recall.
            // Falls back to single ragQuery if ragQueries not defined.
            const queries = section.ragQueries?.length
              ? section.ragQueries
              : [section.ragQuery]

            let ragChunks
            try {
              ragChunks = await queryRagCorpusMulti(
                corpusName,
                queries,
                20,          // topK per query
                sectionAc.signal,
                60,          // max total chunks after dedup
              )
            } catch (err) {
              emit({
                type: 'section_error',
                section: section.key,
                displayName: section.displayName,
                error: `RAG retrieval failed: ${String(err)}`,
              })
              return { missing: [], flags: [] }
            }

            // ── Gemini extraction (with automatic row batching) ──────────────
            let result
            try {
              result = await analyzeSection(ragChunks, section, sectionPeriods, sectionAc.signal)
            } catch (err) {
              emit({
                type: 'section_error',
                section: section.key,
                displayName: section.displayName,
                error: `AI extraction failed: ${String(err)}`,
              })
              return { missing: [], flags: [] }
            }

            // ── Formula post-processing ──────────────────────────────────────
            // Fills in calculated rows that Gemini couldn't derive (e.g., missing subtotals).
            result.cells = applyFormulas(result.cells, section.requiredRows, sectionPeriods)

            // ── Persist cells ────────────────────────────────────────────────
            if (result.cells.length > 0) {
              const cellRows = result.cells.map(cell => ({
                workbook_id: workbookId,
                section: section.key,
                row_key: cell.row_key,
                period: cell.period,
                raw_value: cell.raw_value,
                display_value: cell.display_value,
                is_calculated: cell.is_calculated,
                source_document_id: resolveDocId(cell.source_filename),
                source_page: cell.source_page,
                source_excerpt: cell.source_excerpt,
                confidence: cell.confidence,
                is_overridden: false,
              }))

              await serviceClient
                .from('workbook_cells')
                .upsert(cellRows, { onConflict: 'workbook_id,section,row_key,period' })

              // Accumulate for cross-section reconciliation
              for (const c of result.cells) {
                allExtractedCells.push({
                  section: section.key,
                  row_key: c.row_key,
                  period: c.period,
                  raw_value: c.raw_value,
                })
              }
            }

            // ── Persist missing-data requests ────────────────────────────────
            // Remove any row that was actually filled in (formula may have resolved it)
            const extractedKeys = new Set(result.cells.map(c => `${c.row_key}:${c.period}`))
            const genuinelyMissing = result.missing.filter(
              m => !extractedKeys.has(`${m.row_key}:${m.period}`)
            )

            if (genuinelyMissing.length > 0) {
              // Clear old open requests for this section before inserting new ones
              await serviceClient
                .from('missing_data_requests')
                .delete()
                .eq('workbook_id', workbookId)
                .eq('section', section.key)
                .eq('status', 'open')

              const missingRows = genuinelyMissing.map(m => ({
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
                missing: genuinelyMissing,
              })
            }

            // ── Persist flags ────────────────────────────────────────────────
            const flagsToInsert = [...result.flags]

            // Auto-add low-confidence flags for cells not already flagged
            const flaggedKeys = new Set(result.flags.map(f => `${f.row_key}:${f.period}`))
            for (const cell of result.cells) {
              if (
                cell.confidence < LOW_CONFIDENCE_THRESHOLD &&
                !flaggedKeys.has(`${cell.row_key}:${cell.period}`)
              ) {
                flagsToInsert.push({
                  row_key: cell.row_key,
                  period: cell.period,
                  flag_type: 'low_confidence',
                  severity: 'warning',
                  title: `Low confidence — ${cell.row_key} (${cell.period})`,
                  body: `Confidence: ${Math.round(cell.confidence * 100)}%. Manual verification recommended.`,
                })
              }
            }

            if (flagsToInsert.length > 0) {
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
              missing_count: genuinelyMissing.length,
              chunks_used: ragChunks.length,
            })

            return { missing: genuinelyMissing, flags: flagsToInsert }

          } finally {
            clearTimeout(sectionTimer)
          }
        }

        // ── Step 3: Process sections in parallel batches ───────────────────
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

        // ── Step 4: Cross-section reconciliation ───────────────────────────
        // Compare key metrics (revenue, net income, COGS, etc.) across sections
        // and insert discrepancy flags for any values that don't align.
        emit({ type: 'status', message: 'Running cross-section reconciliation…' })
        try {
          const reconFlags = buildReconciliationFlags(allExtractedCells, workbookPeriods)

          if (reconFlags.length > 0) {
            // Look up cell IDs for reconciliation flag attachment
            const { data: allSavedCells } = await serviceClient
              .from('workbook_cells')
              .select('id, section, row_key, period')
              .eq('workbook_id', workbookId)

            const cellIdMap = new Map(
              (allSavedCells ?? []).map(c => [`${c.section}:${c.row_key}:${c.period}`, c.id])
            )

            const reconFlagRows = reconFlags.map(f => ({
              workbook_id: workbookId,
              cell_id: cellIdMap.get(`overview:${f.row_key}:${f.period}`) ?? null,
              section: 'overview',
              row_key: f.row_key,
              period: f.period,
              flag_type: f.flag_type,
              severity: f.severity,
              title: f.title,
              body: f.body,
              created_by_ai: true,
            }))

            const { data: insertedReconFlags } = await serviceClient
              .from('cell_flags')
              .insert(reconFlagRows)
              .select('id, body')

            if (insertedReconFlags && insertedReconFlags.length > 0) {
              const commentRows = insertedReconFlags
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

            emit({
              type: 'status',
              message: `Found ${reconFlags.length} cross-section discrepancy flag(s).`,
            })
          }
        } catch (reconErr) {
          console.error('Reconciliation step failed (non-fatal):', reconErr)
        }

        // ── Validation gate (Prompt 10) ────────────────────────────────────
        // Require ≥60% completeness to reach draft_ready; otherwise needs_more_data.
        let finalStatus = 'ready'
        try {
          const { data: allCells } = await serviceClient
            .from('workbook_cells')
            .select('section, row_key, period, raw_value')
            .eq('workbook_id', workbookId)
            .not('raw_value', 'is', null)

          const inputSections = SECTION_CONFIGS.filter(s => s.key !== 'overview' && !s.overridePeriods)
          let totalExpected = 0
          let totalPopulated = 0
          const populatedKeys = new Set((allCells ?? []).map(c => `${c.section}::${c.row_key}::${c.period}`))

          for (const sec of inputSections) {
            const inputRows = sec.requiredRows.filter(r => !r.formula && !r.isCalculated)
            for (const row of inputRows) {
              for (const period of workbookPeriods) {
                totalExpected++
                if (populatedKeys.has(`${sec.key}::${row.rowKey}::${period}`)) totalPopulated++
              }
            }
          }

          const completenessPercent = totalExpected > 0 ? Math.round((totalPopulated / totalExpected) * 100) : 0

          if (completenessPercent < 60 || allMissing.length > 0) {
            finalStatus = allMissing.length > 0 ? 'needs_input' : 'needs_more_data'
          } else {
            finalStatus = 'ready'
          }

          emit({ type: 'status', message: `Completeness: ${completenessPercent}% — status: ${finalStatus}` })
        } catch {
          finalStatus = allMissing.length > 0 ? 'needs_input' : 'ready'
        }

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
