'use client'

import { useState, useCallback, useEffect, useRef } from 'react'
import { useRouter } from 'next/navigation'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Button } from '@/components/ui/button'
import { Progress } from '@/components/ui/progress'
import { Input } from '@/components/ui/input'
import {
  Upload,
  FileText,
  CheckCircle2,
  AlertCircle,
  Loader2,
  X,
  Circle,
} from 'lucide-react'
import Image from 'next/image'
import { MissingDataPrompt } from '@/components/missing-data-prompt'

interface NewWorkbookDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  onWorkbookCreated?: () => void
}

type Step = 'upload' | 'uploading' | 'analyzing' | 'missing' | 'complete'
type DocumentStatus = 'pending' | 'complete' | 'missing'

const connectedIntegrations = [
  { name: 'Sandbox Alpine', status: 'complete' as const },
]

const integrations = [
  { name: 'Dropbox', logo: '/integrations/dropbox.png' },
  { name: 'OneDrive', logo: '/integrations/onedrive.png' },
  { name: 'QuickBooks', logo: '/integrations/quickbooks.png' },
  { name: 'Google Drive', logo: '/integrations/google-drive.png' },
  { name: 'Xero', logo: '/integrations/xero.png' },
  { name: 'Sage', logo: '/integrations/sage.png' },
  { name: 'NetSuite', logo: '/integrations/netsuite.png' },
  { name: 'Stripe', logo: '/integrations/stripe.png' },
  { name: 'Plaid', logo: '/integrations/plaid.png' },
  { name: 'Airtable', logo: '/integrations/airtable.png' },
  { name: 'Excel', logo: '/integrations/excel.png' },
]

const initialRequiredDocs: { id: string; name: string; docType: string; status: DocumentStatus }[] = [
  { id: 'general-ledger', name: 'General Ledger', docType: 'general_ledger', status: 'pending' },
  { id: 'bank-statements', name: 'Bank Statements', docType: 'bank_statements', status: 'pending' },
  { id: 'trial-balance', name: 'Trial Balance', docType: 'trial_balance', status: 'pending' },
  { id: 'financials', name: 'Financial Statements', docType: 'financials', status: 'pending' },
]

interface SseEvent {
  type: string
  section?: string
  displayName?: string
  message?: string      // for 'status' events (corpus creation, pre-flight import)
  cells_extracted?: number
  missing_count?: number
  missing?: unknown[]
  status?: string
  total_missing?: number
  error?: string
  resumed?: boolean
}

interface MissingRequest {
  id: string
  section: string
  field_key: string
  period: string | null
  reason: string | null
  suggested_doc: string | null
  status: string
}

export function NewWorkbookDialog({ open, onOpenChange, onWorkbookCreated }: NewWorkbookDialogProps) {
  const router = useRouter()
  const [step, setStep] = useState<Step>('upload')
  const [companyName, setCompanyName] = useState('')
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([])
  const [documents, setDocuments] = useState(initialRequiredDocs)
  const [isDragging, setIsDragging] = useState(false)
  const [scrollPosition, setScrollPosition] = useState(0)
  const [showIntegrationsMenu, setShowIntegrationsMenu] = useState(false)

  const [workbookId, setWorkbookId] = useState<string | null>(null)
  const [uploadProgress, setUploadProgress] = useState(0)
  const [uploadStatus, setUploadStatus] = useState('')
  const [analyzeLog, setAnalyzeLog] = useState<{ section: string; status: string }[]>([])
  const [analyzeProgress, setAnalyzeProgress] = useState(0)
  const [reconnectMsg, setReconnectMsg] = useState<string | null>(null)
  const [missingRequests, setMissingRequests] = useState<MissingRequest[]>([])
  const [error, setError] = useState<string | null>(null)
  const [analyzeElapsed, setAnalyzeElapsed] = useState(0)

  const abortRef = useRef<AbortController | null>(null)
  const analyzeTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Auto-scroll animation for integrations
  useEffect(() => {
    if (step === 'upload') {
      const interval = setInterval(() => {
        setScrollPosition((prev) => {
          const next = prev + 0.3
          const max = integrations.length * 36
          return next >= max ? next - max : next
        })
      }, 20)
      return () => clearInterval(interval)
    }
  }, [step])

  const resetDialog = useCallback(() => {
    abortRef.current?.abort()
    abortRef.current = null
    if (analyzeTimerRef.current) {
      clearInterval(analyzeTimerRef.current)
      analyzeTimerRef.current = null
    }
    setStep('upload')
    setCompanyName('')
    setUploadedFiles([])
    setDocuments(initialRequiredDocs)
    setWorkbookId(null)
    setUploadProgress(0)
    setUploadStatus('')
    setAnalyzeLog([])
    setAnalyzeProgress(0)
    setReconnectMsg(null)
    setMissingRequests([])
    setError(null)
    setAnalyzeElapsed(0)
  }, [])

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(true)
  }, [])

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
  }, [])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    setUploadedFiles((prev) => [...prev, ...Array.from(e.dataTransfer.files)])
  }, [])

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setUploadedFiles((prev) => [...prev, ...Array.from(e.target.files!)])
    }
  }

  const removeFile = (index: number) => {
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index))
  }

  const handleStart = async () => {
    if (!companyName.trim() || uploadedFiles.length === 0) return

    setError(null)
    setStep('uploading')

    // Create the abort controller up-front so it covers every phase.
    abortRef.current?.abort()
    const abort = new AbortController()
    abortRef.current = abort

    /** Abort-aware sleep — resolves immediately if the dialog is closed. */
    const sleep = (ms: number) =>
      new Promise<void>((resolve) => {
        const id = setTimeout(resolve, ms)
        abort.signal.addEventListener('abort', () => { clearTimeout(id); resolve() }, { once: true })
      })

    try {
      // ── 1. Create workbook ────────────────────────────────────
      setUploadStatus('Creating workbook...')
      const wbRes = await fetch('/api/workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ company_name: companyName }),
        signal: abort.signal,
      })
      if (!wbRes.ok) {
        const err = await wbRes.json().catch(() => ({}))
        throw new Error((err as { error?: string }).error || 'Failed to create workbook')
      }
      const wb = await wbRes.json()
      setWorkbookId(wb.id)

      // ── 2. Upload files ───────────────────────────────────────
      for (let i = 0; i < uploadedFiles.length; i++) {
        if (abort.signal.aborted) return
        const file = uploadedFiles[i]
        setUploadStatus(`Uploading ${file.name} (${i + 1}/${uploadedFiles.length})...`)
        setUploadProgress((i / uploadedFiles.length) * 50)

        const formData = new FormData()
        formData.append('file', file)
        try {
          const docRes = await fetch(`/api/workbooks/${wb.id}/documents`, {
            method: 'POST',
            body: formData,
            signal: abort.signal,
          })
          if (docRes.ok) {
            const doc = await docRes.json()
            setDocuments((prev) =>
              prev.map((d) =>
                d.docType === doc.doc_type ? { ...d, status: 'complete' as DocumentStatus } : d
              )
            )
          }
        } catch {
          if (abort.signal.aborted) return
          // upload failed for this file — continue with others
        }
        setUploadProgress(((i + 1) / uploadedFiles.length) * 50)
      }

      // ── 3: Analyze — single SSE stream handles corpus creation + imports + analysis ──
      setUploadProgress(100)
      setStep('analyzing')
      setAnalyzeProgress(0)
      setAnalyzeElapsed(0)
      if (analyzeTimerRef.current) clearInterval(analyzeTimerRef.current)
      analyzeTimerRef.current = setInterval(() => {
        setAnalyzeElapsed((s) => s + 1)
      }, 1_000)

      // On stream error/disconnect, retry up to 3 times (resumes from completed sections).
      // The server-side resume check skips already-persisted sections.
      let done = false
      let streamErrors = 0

      while (!done && !abort.signal.aborted) {
        try {
          const res = await fetch(`/api/workbooks/${wb.id}/analyze`, { signal: abort.signal })

          if (!res.ok) {
            const body = await res.text()
            throw new Error(body || `HTTP ${res.status}`)
          }
          if (!res.body) throw new Error('No response stream')

          setReconnectMsg(null)
          const reader = res.body.getReader()
          const dec = new TextDecoder()
          let buf = ''

          outer: while (true) {
            const { done: streamDone, value } = await reader.read()
            if (streamDone) break

            buf += dec.decode(value, { stream: true })
            const parts = buf.split('\n\n')
            buf = parts.pop() ?? ''

            for (const part of parts) {
              for (const line of part.split('\n')) {
                if (!line.startsWith('data: ')) continue
                let evt: SseEvent
                try { evt = JSON.parse(line.slice(6)) } catch { continue }

                if (evt.type === 'status') {
                  setReconnectMsg(evt.message ?? null)
                }

                if (evt.type === 'section_start') {
                  setReconnectMsg(null)
                  setAnalyzeLog((prev) => [
                    ...prev,
                    { section: evt.displayName ?? evt.section ?? '', status: 'analyzing' },
                  ])
                }

                if (evt.type === 'section_complete') {
                  const label = evt.displayName ?? evt.section ?? ''
                  const status = evt.resumed
                    ? `${evt.cells_extracted} cells (resumed)`
                    : `${evt.cells_extracted} cells`
                  setAnalyzeLog((prev) => {
                    const exists = prev.some((l) => l.section === label || l.section === evt.section)
                    if (exists) return prev.map((l) =>
                      l.section === label || l.section === evt.section ? { ...l, status } : l
                    )
                    return [...prev, { section: label, status }]
                  })
                  setAnalyzeProgress((prev) => Math.min(prev + 3.5, 95))
                }

                if (evt.type === 'section_error') {
                  const label = evt.displayName ?? evt.section ?? ''
                  setAnalyzeLog((prev) =>
                    prev.map((l) =>
                      l.section === label || l.section === evt.section ? { ...l, status: 'error' } : l
                    )
                  )
                }

                if (evt.type === 'complete') {
                  setAnalyzeProgress(100)
                  if (analyzeTimerRef.current) {
                    clearInterval(analyzeTimerRef.current)
                    analyzeTimerRef.current = null
                  }
                  done = true
                  if (evt.total_missing && evt.total_missing > 0) {
                    const missingRes = await fetch(`/api/workbooks/${wb.id}/missing`)
                    const missing: MissingRequest[] = await missingRes.json()
                    setMissingRequests(missing)
                    setStep('missing')
                  } else {
                    setStep('complete')
                    setTimeout(() => {
                      onWorkbookCreated?.()
                      router.push(`/dashboard/workbook/${wb.id}`)
                      onOpenChange(false)
                      resetDialog()
                    }, 1500)
                  }
                  break outer
                }

                if (evt.type === 'error') {
                  setError(evt.error ?? 'Analysis failed')
                  done = true
                  break outer
                }
              }
            }
          }
        } catch (err) {
          if (abort.signal.aborted) return
          streamErrors++
          if (streamErrors >= 3) {
            // After 3 stream errors navigate to workbook — resume will pick up on next open
            onWorkbookCreated?.()
            router.push(`/dashboard/workbook/${wb.id}`)
            onOpenChange(false)
            resetDialog()
            return
          }
          const retryDelay = streamErrors * 5_000
          setReconnectMsg(`Connection interrupted — retrying in ${retryDelay / 1000}s…`)
          await sleep(retryDelay)
          if (abort.signal.aborted) return
          setReconnectMsg(null)
        }
      }
    } catch (err) {
      if (abort.signal.aborted) return
      setError(err instanceof Error ? err.message : String(err))
      setStep('upload')
    }
  }

  const handleMissingResolved = () => {
    if (!workbookId) return
    setStep('complete')
    onWorkbookCreated?.()
    setTimeout(() => {
      router.push(`/dashboard/workbook/${workbookId}`)
      onOpenChange(false)
      resetDialog()
    }, 1500)
  }

  const handleProceedAnyway = () => {
    if (!workbookId) return
    setStep('complete')
    onWorkbookCreated?.()
    setTimeout(() => {
      router.push(`/dashboard/workbook/${workbookId}`)
      onOpenChange(false)
      resetDialog()
    }, 1500)
  }

  const canStart = companyName.trim().length > 0 && uploadedFiles.length > 0

  return (
    <Dialog
      open={open}
      onOpenChange={(v) => {
        if (!v) resetDialog()
        onOpenChange(v)
      }}
    >
      <DialogContent className="max-w-[90vw] max-h-[95vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-2xl font-semibold">Create New Workbook</DialogTitle>
          <DialogDescription className="text-sm">
            {step === 'upload' && 'Upload your financial documents or connect an integration'}
            {step === 'uploading' && uploadStatus}
            {step === 'analyzing' && 'AI is analyzing your documents section by section…'}
            {step === 'missing' && 'Some data could not be found. Upload additional documents or skip.'}
            {step === 'complete' && 'Your workbook is ready!'}
          </DialogDescription>
        </DialogHeader>

        {error && (
          <div className="flex items-start gap-2 p-3 rounded-md bg-red-50 border border-red-200 text-sm text-red-700">
            <AlertCircle className="h-4 w-4 flex-shrink-0 mt-0.5" />
            <div className="flex-1 space-y-2">
              <span>{error}</span>
              {workbookId && (
                <div>
                  <Button
                    size="sm"
                    variant="outline"
                    className="text-xs border-red-300 text-red-700 hover:bg-red-100"
                    onClick={() => {
                      onWorkbookCreated?.()
                      router.push(`/dashboard/workbook/${workbookId}`)
                      onOpenChange(false)
                      resetDialog()
                    }}
                  >
                    View workbook anyway
                  </Button>
                </div>
              )}
            </div>
            <button onClick={() => setError(null)} className="ml-auto flex-shrink-0">
              <X className="h-4 w-4" />
            </button>
          </div>
        )}

        {/* ── Upload ───────────────────────────────────── */}
        {step === 'upload' && (
          <div className="space-y-8">
            <div className="mb-6">
              <label className="text-sm font-medium mb-2 block">Company Name</label>
              <Input
                placeholder="Enter company name"
                value={companyName}
                onChange={(e) => setCompanyName(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && canStart && handleStart()}
              />
            </div>

            {connectedIntegrations.length > 0 && (
              <div className="flex items-center gap-2 mb-3">
                {connectedIntegrations.map((ci) => (
                  <div key={ci.name} className="flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-blue-50 border border-blue-200">
                    <CheckCircle2 className="h-3.5 w-3.5 text-blue-600" />
                    <span className="text-xs font-medium text-blue-700">{ci.name}</span>
                    <span className="text-[10px] text-blue-500 capitalize">{ci.status}</span>
                  </div>
                ))}
              </div>
            )}

            <div className="relative mt-16">
              <div className="absolute -top-8 left-0 right-0 z-10 flex items-center gap-3 px-4">
                <span className="text-xs text-gray-500 whitespace-nowrap">Integrate with the tools you already use</span>
                <div className="relative flex-1 max-w-md overflow-hidden">
                  <div className="absolute left-0 top-0 bottom-0 w-8 bg-gradient-to-r from-white to-transparent z-10 pointer-events-none" />
                  <div className="absolute right-0 top-0 bottom-0 w-8 bg-gradient-to-l from-white to-transparent z-10 pointer-events-none" />
                  <div
                    className="flex gap-3 py-1"
                    style={{ transform: `translateX(-${scrollPosition}px)`, transition: 'transform 0.05s linear' }}
                  >
                    {[...integrations, ...integrations, ...integrations, ...integrations].map((integration, idx) => (
                      <button
                        key={`${integration.name}-${idx}`}
                        onClick={() => setShowIntegrationsMenu(true)}
                        className="flex-shrink-0 hover:scale-110 transition-transform"
                        title={integration.name}
                      >
                        <div className="relative w-6 h-6">
                          <Image src={integration.logo} alt={integration.name} fill className="object-contain" />
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <label
                htmlFor="file-upload"
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                className={`relative rounded-3xl border-2 border-solid transition-all duration-300 cursor-pointer flex items-center justify-center group ${
                  isDragging ? 'border-blue-400/80 bg-blue-50/60' : 'border-blue-300/90 bg-blue-50/25 hover:bg-blue-50/60 hover:border-blue-400/80'
                }`}
                style={{ minHeight: '280px' }}
              >
                <h3 className="text-base font-semibold text-blue-500/90 group-hover:text-blue-600 transition-colors duration-200">
                  Upload files to this agent
                </h3>
                <input
                  type="file"
                  multiple
                  onChange={handleFileSelect}
                  className="hidden"
                  id="file-upload"
                  accept=".pdf,.xlsx,.xls,.csv"
                />
              </label>
            </div>

            {uploadedFiles.length > 0 && (
              <div>
                <h4 className="text-sm font-medium mb-3">Uploaded Files ({uploadedFiles.length})</h4>
                <div className="space-y-3">
                  {uploadedFiles.map((file, index) => (
                    <div key={index} className="flex items-center justify-between py-2 hover:bg-gray-50 transition-colors rounded-lg px-2">
                      <div className="flex items-center gap-3 flex-1 min-w-0">
                        <FileText className="h-5 w-5 text-gray-400 flex-shrink-0" />
                        <span className="text-sm text-gray-600 truncate">{file.name}</span>
                      </div>
                      <button onClick={() => removeFile(index)} className="hover:bg-gray-100 rounded-full p-1 transition-colors flex-shrink-0">
                        <X className="h-4 w-4 text-gray-400" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div>
              <h4 className="text-sm font-medium mb-3">Required Documents</h4>
              <div className="grid grid-cols-2 gap-2">
                {documents.map((doc) => (
                  <div key={doc.id} className="flex items-center gap-3 p-3 rounded-lg border bg-white">
                    {doc.status === 'complete' ? (
                      <CheckCircle2 className="h-5 w-5 text-green-500 flex-shrink-0" />
                    ) : doc.status === 'missing' ? (
                      <AlertCircle className="h-5 w-5 text-orange-500 flex-shrink-0" />
                    ) : (
                      <Circle className="h-5 w-5 text-gray-300 flex-shrink-0" />
                    )}
                    <span className="text-sm text-gray-700">{doc.name}</span>
                  </div>
                ))}
              </div>
            </div>

            <Button className="w-full" onClick={handleStart} disabled={!canStart}>
              <Upload className="mr-2 h-4 w-4" />
              Upload &amp; Analyze
            </Button>

            {showIntegrationsMenu && (
              <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setShowIntegrationsMenu(false)}>
                <div className="bg-white rounded-2xl p-8 max-w-2xl w-full mx-4 shadow-2xl" onClick={(e) => e.stopPropagation()}>
                  <div className="flex items-center justify-between mb-6">
                    <h3 className="text-xl font-semibold">Connect an Integration</h3>
                    <button onClick={() => setShowIntegrationsMenu(false)} className="text-gray-400 hover:text-gray-600 transition-colors">
                      <X className="h-5 w-5" />
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-3 justify-center">
                    {integrations.map((integration) => (
                      <button
                        key={integration.name}
                        onClick={() => setShowIntegrationsMenu(false)}
                        title={integration.name}
                        className="p-3 rounded-xl border border-gray-200 hover:border-blue-500 hover:bg-blue-50 transition-all hover:scale-110"
                      >
                        <div className="relative w-8 h-8">
                          <Image src={integration.logo} alt={integration.name} fill className="object-contain" />
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── Uploading ────────────────────────────────── */}
        {step === 'uploading' && (
          <div className="pt-2 pb-6">
            <div className="text-center space-y-6">
              <Loader2 className="h-5 w-5 text-blue-600 animate-spin mx-auto mt-8 mb-2" />
              <div>
                <h3 className="text-lg font-semibold mb-2">
                  {uploadStatus.startsWith('Preparing AI workbook') ? '' : 'Uploading Documents'}
                </h3>
                {!uploadStatus.startsWith('Preparing AI workbook') && (
                  <p className="text-sm text-gray-500">{uploadStatus}</p>
                )}
              </div>
              <div className="max-w-md mx-auto mt-16">
                <Progress value={uploadProgress} className="h-2" />
                <p className="text-xs text-gray-500 mt-2">{Math.round(uploadProgress)}% complete</p>
              </div>
              <div className="grid grid-cols-2 gap-2 max-w-md mx-auto">
                {documents.map((doc) => (
                  <div key={doc.id} className="flex items-center gap-2 p-2 rounded-lg border bg-white text-sm">
                    {doc.status === 'complete' ? (
                      <CheckCircle2 className="h-4 w-4 text-green-500 flex-shrink-0" />
                    ) : (
                      <Circle className="h-4 w-4 text-gray-300 flex-shrink-0" />
                    )}
                    {doc.name}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── Analyzing ────────────────────────────────── */}
        {step === 'analyzing' && (
          <div className="py-4 space-y-4">
            {reconnectMsg && (
              <div className="flex items-center gap-2 text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2">
                <Loader2 className="h-3 w-3 animate-spin flex-shrink-0" />
                {reconnectMsg}
              </div>
            )}

            {/* Minimal header row */}
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Loader2 className="h-4 w-4 text-blue-500 animate-spin" />
                <span className="text-sm font-medium">Analyzing with AI</span>
              </div>
              <span className="text-xs text-muted-foreground tabular-nums">
                {Math.round(analyzeProgress)}%
              </span>
            </div>

            <Progress value={analyzeProgress} className="h-1" />

            {/* Section log */}
            {analyzeLog.length > 0 && (
              <div className="rounded-lg border divide-y overflow-hidden">
                {analyzeLog.map((entry, i) => (
                  <div
                    key={i}
                    className="flex items-center justify-between px-3 py-2 text-xs bg-background hover:bg-muted/30 transition-colors"
                  >
                    <div className="flex items-center gap-2">
                      {entry.status === 'analyzing' ? (
                        <Loader2 className="h-3 w-3 text-blue-500 animate-spin flex-shrink-0" />
                      ) : entry.status === 'error' ? (
                        <AlertCircle className="h-3 w-3 text-red-400 flex-shrink-0" />
                      ) : (
                        <CheckCircle2 className="h-3 w-3 text-green-500 flex-shrink-0" />
                      )}
                      <span className="font-medium">{entry.section}</span>
                    </div>
                    <span
                      className={
                        entry.status === 'error'
                          ? 'text-red-400'
                          : entry.status === 'analyzing'
                          ? 'text-blue-500'
                          : 'text-muted-foreground'
                      }
                    >
                      {entry.status === 'analyzing' ? 'analyzing…' : entry.status}
                    </span>
                  </div>
                ))}
              </div>
            )}

            {/* Escape hatch — shown after 30 s so users are never stuck */}
            {analyzeElapsed >= 30 && workbookId && (
              <div className="pt-2 border-t flex items-center justify-between">
                <p className="text-xs text-muted-foreground">
                  {analyzeElapsed >= 90
                    ? 'Analysis is still running in the background.'
                    : 'Taking longer than expected?'}
                </p>
                <Button
                  size="sm"
                  variant="outline"
                  className="text-xs"
                  onClick={() => {
                    abortRef.current?.abort()
                    onWorkbookCreated?.()
                    router.push(`/dashboard/workbook/${workbookId}`)
                    onOpenChange(false)
                    resetDialog()
                  }}
                >
                  View workbook
                </Button>
              </div>
            )}
          </div>
        )}

        {/* ── Missing Data ──────────────────────────────── */}
        {step === 'missing' && workbookId && (
          <div className="py-4 space-y-6">
            <MissingDataPrompt
              workbookId={workbookId}
              requests={missingRequests}
              onResolved={handleMissingResolved}
            />
            <div className="flex justify-end gap-3">
              <Button variant="outline" onClick={handleProceedAnyway}>
                Proceed Anyway
              </Button>
              <Button onClick={handleMissingResolved}>
                Continue to Workbook
              </Button>
            </div>
          </div>
        )}

        {/* ── Complete ─────────────────────────────────── */}
        {step === 'complete' && (
          <div className="py-12">
            <div className="text-center space-y-6">
              <div className="mx-auto w-20 h-20 rounded-full bg-green-500 flex items-center justify-center">
                <CheckCircle2 className="h-12 w-12 text-white" />
              </div>
              <div>
                <h3 className="text-2xl font-bold mb-2">Workbook Created!</h3>
                <p className="text-gray-500">Your QoE workbook is ready. Redirecting you now…</p>
              </div>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  )
}
