'use client'

import { useState, useEffect } from 'react'
import { Sheet, SheetContent, SheetHeader, SheetTitle } from '@/components/ui/sheet'
import { Button } from '@/components/ui/button'
import { FileText, ExternalLink, Loader2, AlertCircle, RefreshCw } from 'lucide-react'
import type { SourceRef } from './auditable-cell'

interface DocumentViewerPanelProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  sourceRef: SourceRef | null
}

export function DocumentViewerPanel({ open, onOpenChange, sourceRef }: DocumentViewerPanelProps) {
  const [signedUrl, setSignedUrl] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [retryKey, setRetryKey] = useState(0)

  useEffect(() => {
    if (!open || !sourceRef?.documentId) {
      setSignedUrl(null)
      setError(null)
      return
    }

    const controller = new AbortController()
    setLoading(true)
    setError(null)
    setSignedUrl(null)

    fetch(`/api/documents/${sourceRef.documentId}/signed-url`, {
      signal: controller.signal,
    })
      .then(async (res) => {
        if (!res.ok) throw new Error(`Could not load document (${res.status})`)
        return res.json()
      })
      .then((data) => {
        setSignedUrl(data.signedUrl)
        setLoading(false)
      })
      .catch((err) => {
        if (err.name === 'AbortError') return
        setError('Failed to load document. Check your connection and try again.')
        setLoading(false)
      })

    return () => controller.abort()
  }, [open, sourceRef?.documentId, retryKey])

  const iframeUrl =
    signedUrl && sourceRef?.page ? `${signedUrl}#page=${sourceRef.page}` : signedUrl ?? ''

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent side="right" className="w-[600px] sm:w-[700px] sm:max-w-[700px] flex flex-col p-0">
        <SheetHeader className="p-6 pb-4 border-b">
          <SheetTitle className="flex items-center gap-2">
            <FileText className="h-5 w-5" />
            {sourceRef?.documentName ?? 'Document Viewer'}
          </SheetTitle>
          {sourceRef?.page && (
            <p className="text-sm text-muted-foreground">Page {sourceRef.page}</p>
          )}
          {sourceRef?.excerpt && (
            <blockquote className="mt-2 border-l-4 border-blue-400 pl-3 text-sm italic text-muted-foreground line-clamp-3">
              &ldquo;{sourceRef.excerpt}&rdquo;
            </blockquote>
          )}
          {signedUrl && (
            <a
              href={iframeUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs text-blue-600 flex items-center gap-1 mt-1 hover:underline"
            >
              Open in new tab <ExternalLink className="h-3 w-3" />
            </a>
          )}
        </SheetHeader>

        <div className="flex-1 relative">
          {loading && (
            <div className="absolute inset-0 flex items-center justify-center">
              <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
            </div>
          )}

          {error && (
            <div className="absolute inset-0 flex flex-col items-center justify-center gap-3 p-6 text-center">
              <AlertCircle className="h-7 w-7 text-red-400" />
              <p className="text-sm font-medium text-red-500">Failed to load document</p>
              <p className="text-xs text-muted-foreground max-w-xs">{error}</p>
              <Button
                size="sm"
                variant="outline"
                onClick={() => setRetryKey((k) => k + 1)}
              >
                <RefreshCw className="mr-2 h-3 w-3" />
                Retry
              </Button>
            </div>
          )}

          {!loading && !error && signedUrl && (
            <iframe
              src={iframeUrl}
              className="w-full h-full border-0"
              title={sourceRef?.documentName ?? 'Document'}
              sandbox="allow-same-origin allow-scripts allow-forms"
            />
          )}

          {!loading && !error && !signedUrl && (
            <div className="absolute inset-0 flex items-center justify-center text-muted-foreground text-sm">
              No document selected
            </div>
          )}
        </div>
      </SheetContent>
    </Sheet>
  )
}
