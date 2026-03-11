/**
 * Vertex AI RAG Engine wrapper.
 * Uses the REST API directly (Node SDK is still in preview with incomplete typings).
 *
 * API reference: https://cloud.google.com/vertex-ai/docs/reference/rest/v1beta1/projects.locations
 * RAG Engine operations use v1beta1.
 */

export interface RagChunk {
  text: string
  sourceUri: string
  pageNumber?: number
  score: number
  documentId?: string
  documentName?: string
}

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

const PROJECT = () => {
  const p = process.env.GOOGLE_CLOUD_PROJECT
  if (!p) throw new Error('GOOGLE_CLOUD_PROJECT env var not set')
  return p
}
// us-central1, us-east1, us-east4 are restricted for new projects by Vertex AI RAG.
// us-west1 is open. Override with GOOGLE_CLOUD_LOCATION env var if needed.
const LOCATION = () => process.env.GOOGLE_CLOUD_LOCATION || 'us-west1'

// v1beta1 base — all RAG Engine operations use this version
const BASE_V1B = () =>
  `https://${LOCATION()}-aiplatform.googleapis.com/v1beta1/projects/${PROJECT()}/locations/${LOCATION()}`

function authHeaders(token: string): Record<string, string> {
  return {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${token}`,
  }
}

/**
 * Create a combined AbortSignal from a timeout (ms) and an optional parent signal.
 * Whichever fires first will abort the combined controller.
 */
function timeoutSignal(ms: number, parent?: AbortSignal): { signal: AbortSignal; clear: () => void } {
  const ac = new AbortController()
  const timer = setTimeout(() => ac.abort(new Error(`RAG request timed out after ${ms}ms`)), ms)
  const onParent = () => ac.abort()
  parent?.addEventListener('abort', onParent, { once: true })
  const clear = () => {
    clearTimeout(timer)
    parent?.removeEventListener('abort', onParent)
  }
  return { signal: ac.signal, clear }
}

/**
 * Delete a RAG corpus. Fire-and-forget — returns as soon as the DELETE
 * request is accepted. We don't need to wait for the LRO to finish.
 */
export async function deleteRagCorpus(corpusName: string): Promise<void> {
  try {
    const token = await getAccessToken()
    const url = `https://${LOCATION()}-aiplatform.googleapis.com/v1beta1/${corpusName}`
    const { signal, clear } = timeoutSignal(20_000)
    await fetch(url, {
      method: 'DELETE',
      headers: authHeaders(token),
      signal,
    }).finally(clear)
    // We intentionally don't throw on non-OK — best-effort cleanup only
  } catch (err) {
    console.warn('deleteRagCorpus (best-effort) failed:', err)
  }
}

/**
 * Create a RAG corpus for a workbook.
 * Returns the full corpus resource name:
 *   projects/.../locations/.../ragCorpora/...
 */
export async function createRagCorpus(workbookId: string): Promise<string> {
  const token = await getAccessToken()
  const url = `${BASE_V1B()}/ragCorpora`
  const body = {
    displayName: `qoe-workbook-${workbookId}`,
    description: `RAG corpus for QoE workbook ${workbookId}`,
  }

  const { signal, clear } = timeoutSignal(30_000)
  const res = await fetch(url, {
    method: 'POST',
    headers: authHeaders(token),
    body: JSON.stringify(body),
    signal,
  }).finally(clear)

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`createRagCorpus failed: ${res.status} ${err}`)
  }

  const operation = await res.json()
  const corpusName = await pollOperation(operation.name, token)
  return corpusName
}

/**
 * Import a GCS file into a RAG corpus.
 * Blocks until the import operation completes (30-120 s for large PDFs).
 * Returns the operation name as an identifier.
 *
 * @param abortSignal  Optional signal from the caller (e.g. pre-flight timeout)
 */
export async function importRagFile(
  corpusName: string,
  gcsUri: string,
  abortSignal?: AbortSignal,
): Promise<string> {
  const token = await getAccessToken()
  const url = `https://${LOCATION()}-aiplatform.googleapis.com/v1beta1/${corpusName}/ragFiles:import`
  const body = {
    importRagFilesConfig: {
      gcsSource: { uris: [gcsUri] },
      ragFileChunkingConfig: {
        // 2048 tokens keeps full financial tables in one chunk.
        // 512 overlap ensures multi-page tables are never split across chunks.
        chunkSize: 2048,
        chunkOverlap: 512,
      },
    },
  }

  const { signal: fetchSig, clear } = timeoutSignal(30_000, abortSignal)
  const res = await fetch(url, {
    method: 'POST',
    headers: authHeaders(token),
    body: JSON.stringify(body),
    signal: fetchSig,
  }).finally(clear)

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`importRagFile failed (${res.status}): ${err.slice(0, 300)}`)
  }

  const operation = await res.json()
  // Poll with a 150 s hard cap per file — large PDFs finish in 30-120 s.
  // Combine with the caller's signal so the pre-flight timeout propagates.
  const { signal: pollSig, clear: clearPoll } = timeoutSignal(150_000, abortSignal)
  try {
    await pollOperation(operation.name, token, pollSig)
  } finally {
    clearPoll()
  }
  return operation.name
}

/**
 * Query the RAG corpus and return ranked chunks.
 * Uses the v1beta1 retrieveContexts endpoint.
 *
 * @param signal  Optional AbortSignal from the caller (section-level timeout)
 * @param topK    Number of chunks to retrieve (default 25 for higher recall)
 */
export async function queryRagCorpus(
  corpusName: string,
  query: string,
  topK = 25,
  signal?: AbortSignal,
): Promise<RagChunk[]> {
  const token = await getAccessToken()
  // Correct endpoint: POST /v1beta1/projects/{proj}/locations/{loc}:retrieveContexts
  const url = `${BASE_V1B()}:retrieveContexts`
  const body = {
    query: {
      text: query,
      similarityTopK: topK,
    },
    ragResources: [
      {
        ragCorpus: corpusName,
      },
    ],
  }

  const { signal: fetchSignal, clear } = timeoutSignal(30_000, signal)
  const res = await fetch(url, {
    method: 'POST',
    headers: authHeaders(token),
    body: JSON.stringify(body),
    signal: fetchSignal,
  }).finally(clear)

  if (!res.ok) {
    const err = await res.text()
    throw new Error(`queryRagCorpus failed: ${res.status} ${err}`)
  }

  const data = await res.json()
  const chunks: RagChunk[] = []

  for (const ctx of data.contexts?.contexts ?? []) {
    chunks.push({
      text: ctx.text ?? '',
      sourceUri: ctx.sourceUri ?? '',
      pageNumber: ctx.pageSpan?.firstPage,
      score: ctx.score ?? 0,
      documentName: ctx.sourceDisplayName ?? ctx.sourceUri?.split('/').pop() ?? '',
    })
  }

  return chunks
}

/**
 * Poll a long-running operation until done.
 * Returns the response resource name (e.g. corpus name after createRagCorpus).
 * Accepts an optional AbortSignal so callers can cancel mid-poll.
 * Uses progressive backoff: 3s → 5s → 8s → 10s (capped) between polls.
 */
async function pollOperation(
  operationName: string,
  token: string,
  abortSignal?: AbortSignal,
): Promise<string> {
  const opUrl = `https://${LOCATION()}-aiplatform.googleapis.com/v1beta1/${operationName}`
  const backoffMs = [3000, 5000, 8000, 10000]
  const MAX_ATTEMPTS = 40 // ~5 min total

  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    if (abortSignal?.aborted) throw new Error('pollOperation aborted')

    await sleep(backoffMs[Math.min(attempt, backoffMs.length - 1)])

    if (abortSignal?.aborted) throw new Error('pollOperation aborted')

    const { signal, clear } = timeoutSignal(15_000, abortSignal)
    const res = await fetch(opUrl, {
      headers: { Authorization: `Bearer ${token}` },
      signal,
    }).finally(clear)

    if (!res.ok) {
      const err = await res.text()
      throw new Error(`pollOperation failed (${res.status}): ${err.slice(0, 300)}`)
    }

    const op = await res.json()
    if (op.done) {
      if (op.error) throw new Error(`LRO failed: ${JSON.stringify(op.error)}`)
      return op.response?.name ?? operationName
    }
  }
  throw new Error(`Operation timed out after ~5 min: ${operationName}`)
}

function sleep(ms: number) {
  return new Promise<void>((r) => setTimeout(r, ms))
}

/**
 * Run multiple RAG queries in parallel and return deduplicated, ranked chunks.
 *
 * Strategy:
 *  1. All queries fire simultaneously (max parallelism).
 *  2. Duplicate chunks (same first 150 chars) are collapsed — keep highest score.
 *  3. Result is sorted by score descending, capped at maxTotal.
 *
 * This dramatically improves recall for sections whose data spans multiple
 * table types (e.g., Income Statement needs revenue, COGS, OpEx, and tax queries).
 */
export async function queryRagCorpusMulti(
  corpusName: string,
  queries: string[],
  topKPerQuery = 20,
  signal?: AbortSignal,
  maxTotal = 60,
): Promise<RagChunk[]> {
  if (queries.length === 0) return []

  const results = await Promise.allSettled(
    queries.map(q => queryRagCorpus(corpusName, q, topKPerQuery, signal))
  )

  // Deduplicate by content fingerprint (first 150 chars), keep highest score
  const seen = new Map<string, RagChunk>()
  for (const result of results) {
    if (result.status !== 'fulfilled') continue
    for (const chunk of result.value) {
      const fingerprint = chunk.text.slice(0, 150).trim()
      const existing = seen.get(fingerprint)
      if (!existing || chunk.score > existing.score) {
        seen.set(fingerprint, chunk)
      }
    }
  }

  return [...seen.values()]
    .sort((a, b) => b.score - a.score)
    .slice(0, maxTotal)
}
