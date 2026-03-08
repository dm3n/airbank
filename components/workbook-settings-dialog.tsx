'use client'

import { useState, useCallback, useEffect, useRef } from 'react'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Button } from '@/components/ui/button'
import { Progress } from '@/components/ui/progress'
import {
  Upload,
  FileText,
  CheckCircle2,
  X,
  Loader2,
  RefreshCw,
  AlertCircle,
  Trash2,
} from 'lucide-react'
import Image from 'next/image'

interface Document {
  id: string
  file_name: string
  doc_type: string
  ingestion_status: string
  file_size: number
  created_at: string
}

interface SseEvent {
  type: string
  section?: string
  displayName?: string
  cells_extracted?: number
  status?: string
  total_missing?: number
  error?: string
  resumed?: boolean
}

export interface WorkbookSettingsDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  workbookName: string
  workbookId?: string
  onAnalysisComplete?: () => void
  onDeleted?: () => void
}

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

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

export function WorkbookSettingsDialog({
  open,
  onOpenChange,
  workbookName,
  workbookId,
  onAnalysisComplete,
  onDeleted,
}: WorkbookSettingsDialogProps) {
  const [documents, setDocuments] = useState<Document[]>([])
  const [docsLoading, setDocsLoading] = useState(false)
  const [confirmDelete, setConfirmDelete] = useState(false)
  const [deleting, setDeleting] = useState(false)
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([])
  const [isDragging, setIsDragging] = useState(false)
  const [scrollPosition, setScrollPosition] = useState(0)
  const [showIntegrationsMenu, setShowIntegrationsMenu] = useState(false)

  const [uploading, setUploading] = useState(false)
  const [uploadProgress, setUploadProgress] = useState(0)
  const [uploadStatus, setUploadStatus] = useState('')
  const [analyzing, setAnalyzing] = useState(false)
  const [analyzeProgress, setAnalyzeProgress] = useState(0)
  const [analyzeLog, setAnalyzeLog] = useState<{ section: string; status: string }[]>([])
  const [reconnectMsg, setReconnectMsg] = useState<string | null>(null)
  const [analyzeComplete, setAnalyzeComplete] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const abortRef = useRef<AbortController | null>(null)

  // Fetch documents when dialog opens
  useEffect(() => {
    if (open && workbookId) {
      setDocsLoading(true)
      fetch(`/api/workbooks/${workbookId}/documents`)
        .then((r) => r.json())
        .then((data) => { if (Array.isArray(data)) setDocuments(data) })
        .catch(console.error)
        .finally(() => setDocsLoading(false))
    }
    if (!open) {
      abortRef.current?.abort()
      abortRef.current = null
      setAnalyzing(false)
      setUploading(false)
      setUploadedFiles([])
      setUploadProgress(0)
      setUploadStatus('')
      setAnalyzeLog([])
      setAnalyzeProgress(0)
      setReconnectMsg(null)
      setAnalyzeComplete(false)
      setError(null)
      setConfirmDelete(false)
      setDeleting(false)
    }
  }, [open, workbookId])

  // Auto-scroll animation for integrations
  useEffect(() => {
    const interval = setInterval(() => {
      setScrollPosition((prev) => {
        const next = prev + 0.3
        const max = integrations.length * 36
        return next >= max ? next - max : next
      })
    }, 20)
    return () => clearInterval(interval)
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

  const handleUploadAndAnalyze = async () => {
    if (!workbookId) return
    setError(null)
    setAnalyzeComplete(false)

    // 1. Upload staged files first
    if (uploadedFiles.length > 0) {
      setUploading(true)
      for (let i = 0; i < uploadedFiles.length; i++) {
        const file = uploadedFiles[i]
        setUploadStatus(`Uploading ${file.name} (${i + 1}/${uploadedFiles.length})…`)
        setUploadProgress((i / uploadedFiles.length) * 100)

        const formData = new FormData()
        formData.append('file', file)
        try {
          const res = await fetch(`/api/workbooks/${workbookId}/documents`, {
            method: 'POST',
            body: formData,
          })
          if (res.ok) {
            const doc = await res.json()
            setDocuments((prev) => [...prev, doc])
          }
        } catch {
          // individual upload failure — continue
        }
        setUploadProgress(((i + 1) / uploadedFiles.length) * 100)
      }
      setUploading(false)
      setUploadedFiles([])
      setUploadProgress(0)
      setUploadStatus('')
    }

    // 2. Run SSE analysis with retry
    setAnalyzing(true)
    setAnalyzeLog([])
    setAnalyzeProgress(0)

    abortRef.current?.abort()
    const abort = new AbortController()
    abortRef.current = abort

    const MAX_RETRIES = 3
    let done = false

    for (let attempt = 0; attempt <= MAX_RETRIES && !done && !abort.signal.aborted; attempt++) {
      if (attempt > 0) {
        const delay = attempt * 2000
        setReconnectMsg(`Connection lost — retrying in ${delay / 1000}s… (${attempt}/${MAX_RETRIES})`)
        await new Promise((r) => setTimeout(r, delay))
        if (abort.signal.aborted) break
        setReconnectMsg(`Reconnecting… (${attempt}/${MAX_RETRIES})`)
      }

      try {
        const res = await fetch(`/api/workbooks/${workbookId}/analyze`, { signal: abort.signal })
        if (!res.ok) throw new Error(await res.text() || `HTTP ${res.status}`)
        if (!res.body) throw new Error('No response stream')

        setReconnectMsg(null)

        const reader = res.body.getReader()
        const dec = new TextDecoder()
        let buf = ''

        outer: while (true) {
          const { done: streamDone, value } = await reader.read()
          if (streamDone) { done = true; break }

          buf += dec.decode(value, { stream: true })
          const parts = buf.split('\n\n')
          buf = parts.pop() ?? ''

          for (const part of parts) {
            for (const line of part.split('\n')) {
              if (!line.startsWith('data: ')) continue
              let evt: SseEvent
              try { evt = JSON.parse(line.slice(6)) } catch { continue }

              if (evt.type === 'section_start') {
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
                  if (exists) return prev.map((l) => (l.section === label || l.section === evt.section) ? { ...l, status } : l)
                  return [...prev, { section: label, status }]
                })
                setAnalyzeProgress((prev) => Math.min(prev + 9, 95))
              }

              if (evt.type === 'section_error') {
                const label = evt.displayName ?? evt.section ?? ''
                setAnalyzeLog((prev) =>
                  prev.map((l) => (l.section === label || l.section === evt.section) ? { ...l, status: 'error' } : l)
                )
              }

              if (evt.type === 'complete') {
                setAnalyzeProgress(100)
                setAnalyzing(false)
                setAnalyzeComplete(true)
                onAnalysisComplete?.()
                done = true
                break outer
              }

              if (evt.type === 'error') {
                setError(evt.error ?? 'Analysis failed')
                setAnalyzing(false)
                done = true
                break outer
              }
            }
          }
        }
      } catch (err) {
        if (abort.signal.aborted) break
        if (attempt >= MAX_RETRIES) {
          setError(err instanceof Error ? err.message : 'Lost connection during analysis')
          setAnalyzing(false)
          done = true
        }
      }
    }
  }

  const handleDelete = async () => {
    if (!workbookId) return
    setDeleting(true)
    setError(null)
    try {
      const res = await fetch(`/api/workbooks/${workbookId}`, { method: 'DELETE' })
      if (!res.ok) {
        const body = await res.json().catch(() => ({}))
        throw new Error((body as { error?: string }).error || `HTTP ${res.status}`)
      }
      onOpenChange(false)
      onDeleted?.()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete workbook')
      setDeleting(false)
      setConfirmDelete(false)
    }
  }

  const isProcessing = uploading || analyzing

  return (
    <Dialog
      open={open}
      onOpenChange={(v) => {
        abortRef.current?.abort()
        onOpenChange(v)
      }}
    >
      <DialogContent className="max-w-[90vw] max-h-[95vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-2xl font-semibold">Workbook Settings</DialogTitle>
          <DialogDescription>
            Manage documents and re-run analysis for {workbookName}
          </DialogDescription>
        </DialogHeader>

        {error && (
          <div className="flex items-center gap-2 p-3 rounded-md bg-red-50 border border-red-200 text-sm text-red-700">
            <AlertCircle className="h-4 w-4 flex-shrink-0" />
            <span className="flex-1">{error}</span>
            <button onClick={() => setError(null)} className="ml-auto flex-shrink-0">
              <X className="h-4 w-4" />
            </button>
          </div>
        )}

        <div className="space-y-8">
          {/* Integration logos + upload zone */}
          <div className="relative mt-12">
            <div className="absolute -top-8 left-0 right-0 z-10 flex items-center gap-3 px-4">
              <span className="text-xs text-gray-500 whitespace-nowrap">Get better answers from your apps</span>
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
              htmlFor="settings-file-upload"
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`relative rounded-3xl border-2 border-solid transition-all duration-300 cursor-pointer flex items-center justify-center group ${
                isDragging ? 'border-blue-400/80 bg-blue-50/60' : 'border-blue-300/90 bg-blue-50/25 hover:bg-blue-50/60 hover:border-blue-400/80'
              }`}
              style={{ minHeight: '180px' }}
            >
              <h3 className={`text-base font-semibold transition-colors duration-300 ${
                'text-blue-500/90 group-hover:text-blue-600'
              }`}>
                Upload files to this agent
              </h3>
              <input
                type="file"
                multiple
                onChange={handleFileSelect}
                className="hidden"
                id="settings-file-upload"
                accept=".pdf,.xlsx,.xls,.csv"
              />
            </label>
          </div>

          {/* Staged files */}
          {uploadedFiles.length > 0 && (
            <div>
              <h4 className="text-sm font-medium mb-2">Files to Upload ({uploadedFiles.length})</h4>
              <div className="space-y-2">
                {uploadedFiles.map((file, i) => (
                  <div key={i} className="flex items-center justify-between py-2 px-2 rounded-lg hover:bg-gray-50">
                    <div className="flex items-center gap-3 flex-1 min-w-0">
                      <FileText className="h-5 w-5 text-gray-400 flex-shrink-0" />
                      <span className="text-sm text-gray-600 truncate">{file.name}</span>
                    </div>
                    <button onClick={() => removeFile(i)} className="p-1 hover:bg-gray-100 rounded-full">
                      <X className="h-4 w-4 text-gray-400" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Upload progress */}
          {uploading && (
            <div className="space-y-2 p-4 rounded-lg border bg-muted/20">
              <div className="flex items-center gap-2">
                <Loader2 className="h-4 w-4 animate-spin text-blue-500 flex-shrink-0" />
                <span className="text-sm font-medium">{uploadStatus}</span>
              </div>
              <Progress value={uploadProgress} className="h-1" />
            </div>
          )}

          {/* Analysis progress */}
          {analyzing && (
            <div className="space-y-3 p-4 rounded-lg border bg-muted/20">
              {reconnectMsg && (
                <div className="flex items-center gap-2 text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2">
                  <Loader2 className="h-3 w-3 animate-spin flex-shrink-0" />
                  {reconnectMsg}
                </div>
              )}
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Loader2 className="h-4 w-4 animate-spin text-blue-500" />
                  <span className="text-sm font-medium">Analyzing with AI</span>
                </div>
                <span className="text-xs text-muted-foreground tabular-nums">{Math.round(analyzeProgress)}%</span>
              </div>
              <Progress value={analyzeProgress} className="h-1" />
              {analyzeLog.length > 0 && (
                <div className="rounded-md border divide-y overflow-hidden">
                  {analyzeLog.map((entry, i) => (
                    <div
                      key={i}
                      className="flex items-center justify-between px-3 py-1.5 text-xs bg-background hover:bg-muted/30 transition-colors"
                    >
                      <div className="flex items-center gap-2">
                        {entry.status === 'analyzing' ? (
                          <Loader2 className="h-3 w-3 text-blue-500 animate-spin flex-shrink-0" />
                        ) : entry.status === 'error' ? (
                          <AlertCircle className="h-3 w-3 text-gray-400 flex-shrink-0" />
                        ) : (
                          <CheckCircle2 className="h-3 w-3 text-blue-500 flex-shrink-0" />
                        )}
                        <span className="font-medium">{entry.section}</span>
                      </div>
                      <span
                        className={
                          entry.status === 'error'
                            ? 'text-gray-400'
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
            </div>
          )}

          {analyzeComplete && (
            <div className="flex items-center gap-3 p-3 rounded-lg bg-blue-50 border border-blue-200">
              <CheckCircle2 className="h-5 w-5 text-blue-600 flex-shrink-0" />
              <p className="text-sm text-blue-700 font-medium">
                Analysis complete! Your workbook data has been updated.
              </p>
            </div>
          )}

          {/* Action button */}
          {workbookId && !isProcessing && (
            <div className="flex gap-3">
              <Button
                className="flex-1"
                onClick={handleUploadAndAnalyze}
                variant={uploadedFiles.length > 0 ? 'default' : 'outline'}
              >
                {uploadedFiles.length > 0 ? (
                  <>
                    <Upload className="mr-2 h-4 w-4" />
                    Upload &amp; Re-analyze
                  </>
                ) : (
                  <>
                    <RefreshCw className="mr-2 h-4 w-4" />
                    Re-analyze Documents
                  </>
                )}
              </Button>
            </div>
          )}

          {/* Data Room */}
          <div className="space-y-4">
            <h3 className="text-lg font-semibold">Data Room</h3>
            {docsLoading ? (
              <div className="flex items-center gap-2 text-sm text-muted-foreground py-4">
                <Loader2 className="h-4 w-4 animate-spin" />
                Loading documents…
              </div>
            ) : documents.length > 0 ? (
              <div className="space-y-2">
                {documents.map((doc) => (
                  <div
                    key={doc.id}
                    className="flex items-center justify-between py-2 px-2 rounded-lg hover:bg-gray-50 transition-colors"
                  >
                    <div className="flex items-center gap-3 flex-1 min-w-0">
                      <FileText className="h-5 w-5 text-gray-400 flex-shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className="text-sm text-gray-700 truncate">{doc.file_name}</p>
                        <p className="text-xs text-muted-foreground">
                          {doc.doc_type.replace(/_/g, ' ')} · {formatFileSize(doc.file_size)}
                        </p>
                      </div>
                    </div>
                    <div className="flex-shrink-0">
                      {doc.ingestion_status === 'ready' ? (
                        <CheckCircle2 className="h-4 w-4 text-blue-500" />
                      ) : doc.ingestion_status === 'error' ? (
                        <X className="h-4 w-4 text-gray-400" />
                      ) : (
                        <Loader2 className="h-4 w-4 animate-spin text-gray-400" />
                      )}
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <p className="text-sm text-muted-foreground py-2">
                {workbookId ? 'No documents uploaded yet.' : 'Demo workbook — no real documents.'}
              </p>
            )}
          </div>
          {/* Delete workbook */}
          {workbookId && !isProcessing && (
            <div className="pt-4 border-t">
              {!confirmDelete ? (
                <Button
                  variant="ghost"
                  size="sm"
                  className="text-muted-foreground hover:text-red-600 hover:bg-red-50 -ml-2"
                  onClick={() => setConfirmDelete(true)}
                >
                  <Trash2 className="h-4 w-4 mr-2" />
                  Delete workbook
                </Button>
              ) : (
                <div className="flex items-center gap-3">
                  <span className="text-sm text-muted-foreground">Are you sure?</span>
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-red-200 text-red-600 hover:bg-red-50"
                    onClick={handleDelete}
                    disabled={deleting}
                  >
                    {deleting ? <><Loader2 className="h-3 w-3 mr-1.5 animate-spin" />Deleting…</> : 'Yes, delete'}
                  </Button>
                  <Button size="sm" variant="ghost" onClick={() => setConfirmDelete(false)} disabled={deleting}>
                    Cancel
                  </Button>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Integrations Menu Modal */}
        {showIntegrationsMenu && (
          <div
            className="fixed inset-0 bg-black/50 flex items-center justify-center z-50"
            onClick={() => setShowIntegrationsMenu(false)}
          >
            <div
              className="bg-white rounded-2xl p-8 max-w-2xl w-full mx-4 shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
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
      </DialogContent>
    </Dialog>
  )
}
