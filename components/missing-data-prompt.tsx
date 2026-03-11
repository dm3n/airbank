'use client'

import { useState, useCallback } from 'react'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { AlertCircle, Upload, CheckCircle2, X, FileText, Loader2 } from 'lucide-react'

interface MissingRequest {
  id: string
  section: string
  field_key: string
  period: string | null
  reason: string | null
  suggested_doc: string | null
  status: string
}

interface MissingDataPromptProps {
  workbookId: string
  requests: MissingRequest[]
  onResolved?: () => void
}

export function MissingDataPrompt({ workbookId, requests, onResolved }: MissingDataPromptProps) {
  const [localRequests, setLocalRequests] = useState(requests)
  const [uploadingFor, setUploadingFor] = useState<string | null>(null)
  const [skipping, setSkipping] = useState<Set<string>>(new Set())

  // Group by suggested_doc
  const byDoc = localRequests.reduce<Record<string, MissingRequest[]>>((acc, r) => {
    const key = r.suggested_doc ?? 'Unknown Document'
    if (!acc[key]) acc[key] = []
    acc[key].push(r)
    return acc
  }, {})

  const handleSkip = async (requestId: string) => {
    setSkipping((prev) => new Set(prev).add(requestId))
    try {
      await fetch(`/api/workbooks/${workbookId}/missing`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requestId, status: 'skipped' }),
      })
      setLocalRequests((prev) => prev.filter((r) => r.id !== requestId))
    } finally {
      setSkipping((prev) => {
        const next = new Set(prev)
        next.delete(requestId)
        return next
      })
    }
  }

  const handleFileUpload = useCallback(
    async (docType: string, file: File) => {
      setUploadingFor(docType)
      try {
        const formData = new FormData()
        formData.append('file', file)
        const res = await fetch(`/api/workbooks/${workbookId}/documents`, {
          method: 'POST',
          body: formData,
        })
        if (!res.ok) throw new Error(await res.text())

        // Mark associated requests as resolved
        const associatedIds = localRequests
          .filter((r) => r.suggested_doc === docType)
          .map((r) => r.id)

        for (const id of associatedIds) {
          await fetch(`/api/workbooks/${workbookId}/missing`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ requestId: id, status: 'resolved' }),
          })
        }

        setLocalRequests((prev) =>
          prev.filter((r) => r.suggested_doc !== docType)
        )

        if (onResolved) onResolved()
      } catch (err) {
        void err // upload failed; uploadingFor cleared in finally
      } finally {
        setUploadingFor(null)
      }
    },
    [workbookId, localRequests, onResolved]
  )

  if (localRequests.length === 0) {
    return (
      <div className="flex items-center gap-3 p-4 rounded-lg bg-blue-50 border border-blue-200">
        <CheckCircle2 className="h-5 w-5 text-blue-600 flex-shrink-0" />
        <p className="text-sm text-blue-700 font-medium">All required data has been resolved.</p>
      </div>
    )
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3 p-4 rounded-lg bg-gray-50 border border-gray-200">
        <AlertCircle className="h-5 w-5 text-gray-600 flex-shrink-0" />
        <div>
          <p className="text-sm font-semibold text-gray-800">
            Missing data detected ({localRequests.length} fields)
          </p>
          <p className="text-xs text-gray-700 mt-0.5">
            Upload additional documents or skip individual fields to proceed.
          </p>
        </div>
      </div>

      {Object.entries(byDoc).map(([docType, reqs]) => (
        <div key={docType} className="rounded-lg border p-4 space-y-3">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <FileText className="h-4 w-4 text-muted-foreground" />
              <span className="text-sm font-semibold">{docType}</span>
              <Badge variant="outline" className="text-xs">
                {reqs.length} field{reqs.length !== 1 ? 's' : ''}
              </Badge>
            </div>

            <label className="cursor-pointer">
              <Button
                size="sm"
                variant="outline"
                disabled={uploadingFor === docType}
                asChild
              >
                <span>
                  {uploadingFor === docType ? (
                    <Loader2 className="mr-2 h-3 w-3 animate-spin" />
                  ) : (
                    <Upload className="mr-2 h-3 w-3" />
                  )}
                  {uploadingFor === docType ? 'Uploading...' : 'Upload Document'}
                </span>
              </Button>
              <input
                type="file"
                className="hidden"
                accept=".pdf,.xlsx,.xls,.csv"
                onChange={(e) => {
                  const file = e.target.files?.[0]
                  if (file) handleFileUpload(docType, file)
                }}
              />
            </label>
          </div>

          <div className="space-y-2">
            {reqs.map((req) => (
              <div
                key={req.id}
                className="flex items-start justify-between gap-3 py-2 px-3 rounded bg-muted/40 text-sm"
              >
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-medium capitalize">
                      {req.field_key.replace(/_/g, ' ')}
                    </span>
                    {req.period && (
                      <Badge variant="secondary" className="text-xs">
                        {req.period}
                      </Badge>
                    )}
                    <Badge variant="outline" className="text-xs">
                      {req.section}
                    </Badge>
                  </div>
                  {req.reason && (
                    <p className="text-xs text-muted-foreground mt-0.5">{req.reason}</p>
                  )}
                </div>
                <button
                  onClick={() => handleSkip(req.id)}
                  disabled={skipping.has(req.id)}
                  className="flex-shrink-0 text-muted-foreground hover:text-foreground transition-colors disabled:opacity-50"
                  title="Skip this field"
                >
                  {skipping.has(req.id) ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <X className="h-4 w-4" />
                  )}
                </button>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}
