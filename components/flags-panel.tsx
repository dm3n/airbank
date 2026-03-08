'use client'

import { useState, useEffect, useCallback } from 'react'
import { Button } from '@/components/ui/button'
import { Textarea } from '@/components/ui/textarea'
import { Flag, X, Check, Loader2, MessageSquare, ChevronDown, ChevronUp } from 'lucide-react'

interface FlagComment {
  id: string
  author_name: string
  body: string
  created_at: string
}

interface CellFlag {
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
  comment_count: number
}

interface FlagsPanelProps {
  workbookId: string
  open: boolean
  onClose: () => void
  onFlagResolved?: () => void
}

const FLAG_TYPE_LABELS: Record<string, string> = {
  low_confidence: 'Low Confidence',
  discrepancy: 'Discrepancy',
  missing: 'Missing Data',
  ai_note: 'AI Note',
  needs_review: 'Needs Review',
}

const SEVERITY_COLORS: Record<string, string> = {
  info: 'text-gray-500',
  warning: 'text-gray-600',
  critical: 'text-gray-700',
}

function FlagItem({
  flag,
  workbookId,
  onResolved,
}: {
  flag: CellFlag
  workbookId: string
  onResolved: () => void
}) {
  const [expanded, setExpanded] = useState(false)
  const [comments, setComments] = useState<FlagComment[]>([])
  const [commentsLoaded, setCommentsLoaded] = useState(false)
  const [newComment, setNewComment] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [resolving, setResolving] = useState(false)

  const loadComments = useCallback(async () => {
    if (commentsLoaded) return
    const res = await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}/comments`)
    if (res.ok) {
      const data = await res.json()
      setComments(Array.isArray(data) ? data : [])
      setCommentsLoaded(true)
    }
  }, [workbookId, flag.id, commentsLoaded])

  const handleExpand = async () => {
    const next = !expanded
    setExpanded(next)
    if (next) await loadComments()
  }

  const handleResolve = async () => {
    setResolving(true)
    try {
      await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ resolved: !flag.resolved_at }),
      })
      onResolved()
    } finally {
      setResolving(false)
    }
  }

  const handleAddComment = async () => {
    if (!newComment.trim()) return
    setSubmitting(true)
    try {
      const res = await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}/comments`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ body: newComment.trim() }),
      })
      if (res.ok) {
        const comment = await res.json()
        setComments(prev => [...prev, comment])
        setNewComment('')
      }
    } finally {
      setSubmitting(false)
    }
  }

  const isResolved = !!flag.resolved_at

  return (
    <div className={`border rounded-lg overflow-hidden ${isResolved ? 'opacity-60' : ''}`}>
      <div className="px-4 py-3">
        <div className="flex items-start justify-between gap-3">
          <div className="flex items-start gap-2 min-w-0">
            <Flag
              className={`h-3.5 w-3.5 mt-0.5 flex-shrink-0 ${flag.created_by_ai ? 'text-blue-500' : 'text-gray-400'}`}
            />
            <div className="min-w-0">
              <p className="text-sm font-medium leading-snug">{flag.title}</p>
              <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                <span className={`text-xs ${SEVERITY_COLORS[flag.severity] ?? 'text-gray-500'}`}>
                  {FLAG_TYPE_LABELS[flag.flag_type] ?? flag.flag_type}
                </span>
                {flag.created_by_ai && (
                  <span className="text-[10px] bg-blue-50 text-blue-600 px-1.5 py-0.5 rounded font-medium">AI</span>
                )}
                <span className="text-[10px] text-muted-foreground">
                  {flag.section} · {flag.row_key}{flag.period ? ` · ${flag.period}` : ''}
                </span>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-1 flex-shrink-0">
            <Button
              size="sm"
              variant="ghost"
              className="h-7 px-2 text-xs"
              onClick={handleResolve}
              disabled={resolving}
            >
              {resolving ? (
                <Loader2 className="h-3 w-3 animate-spin" />
              ) : isResolved ? (
                <>
                  <X className="h-3 w-3 mr-1" />
                  Reopen
                </>
              ) : (
                <>
                  <Check className="h-3 w-3 mr-1" />
                  Resolve
                </>
              )}
            </Button>
            <button
              className="p-1 text-muted-foreground hover:text-foreground transition-colors"
              onClick={handleExpand}
            >
              {expanded ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
            </button>
          </div>
        </div>

        {flag.comment_count > 0 && !expanded && (
          <button
            className="flex items-center gap-1 mt-2 text-xs text-muted-foreground hover:text-foreground transition-colors"
            onClick={handleExpand}
          >
            <MessageSquare className="h-3 w-3" />
            {flag.comment_count} comment{flag.comment_count !== 1 ? 's' : ''}
          </button>
        )}
      </div>

      {expanded && (
        <div className="border-t bg-muted/20 px-4 py-3 space-y-3">
          {/* Existing comments */}
          {comments.length > 0 && (
            <div className="space-y-2">
              {comments.map(comment => (
                <div key={comment.id} className="text-xs">
                  <div className="flex items-baseline gap-2">
                    <span className="font-medium">{comment.author_name}</span>
                    <span className="text-muted-foreground text-[10px]">
                      {new Date(comment.created_at).toLocaleDateString('en-US', {
                        month: 'short',
                        day: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </span>
                  </div>
                  <p className="text-muted-foreground mt-0.5 leading-relaxed">{comment.body}</p>
                </div>
              ))}
            </div>
          )}
          {commentsLoaded && comments.length === 0 && (
            <p className="text-xs text-muted-foreground">No comments yet.</p>
          )}

          {/* Add comment */}
          <div className="space-y-2">
            <Textarea
              value={newComment}
              onChange={e => setNewComment(e.target.value)}
              placeholder="Add a comment…"
              className="text-xs resize-none"
              rows={2}
            />
            <Button
              size="sm"
              variant="outline"
              onClick={handleAddComment}
              disabled={!newComment.trim() || submitting}
              className="w-full text-xs"
            >
              {submitting ? <Loader2 className="h-3 w-3 animate-spin mr-1" /> : null}
              Add Comment
            </Button>
          </div>
        </div>
      )}
    </div>
  )
}

export function FlagsPanel({ workbookId, open, onClose, onFlagResolved }: FlagsPanelProps) {
  const [flags, setFlags] = useState<CellFlag[]>([])
  const [loading, setLoading] = useState(false)
  const [showResolved, setShowResolved] = useState(false)

  const fetchFlags = useCallback(async () => {
    setLoading(true)
    try {
      const res = await fetch(`/api/workbooks/${workbookId}/flags`)
      if (res.ok) {
        const data = await res.json()
        setFlags(Array.isArray(data) ? data : [])
      }
    } finally {
      setLoading(false)
    }
  }, [workbookId])

  useEffect(() => {
    if (open) fetchFlags()
  }, [open, fetchFlags])

  const handleResolved = useCallback(() => {
    fetchFlags()
    onFlagResolved?.()
  }, [fetchFlags, onFlagResolved])

  if (!open) return null

  const openFlags = flags.filter(f => !f.resolved_at)
  const resolvedFlags = flags.filter(f => f.resolved_at)

  return (
    <div className="fixed inset-y-0 right-0 z-40 flex">
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-black/20"
        onClick={onClose}
      />
      {/* Panel */}
      <div className="relative ml-auto w-96 bg-background border-l shadow-xl flex flex-col h-full">
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-4 border-b flex-shrink-0">
          <div className="flex items-center gap-2">
            <Flag className="h-4 w-4 text-blue-500" />
            <h2 className="text-base font-semibold">Flags</h2>
            {openFlags.length > 0 && (
              <span className="bg-blue-500 text-white text-xs font-medium rounded-full px-2 py-0.5">
                {openFlags.length}
              </span>
            )}
          </div>
          <button
            onClick={onClose}
            className="text-muted-foreground hover:text-foreground transition-colors"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-3">
          {loading && (
            <div className="flex items-center justify-center py-10">
              <Loader2 className="h-5 w-5 animate-spin text-muted-foreground" />
            </div>
          )}

          {!loading && flags.length === 0 && (
            <div className="flex flex-col items-center justify-center py-16 text-center">
              <Flag className="h-8 w-8 text-muted-foreground/30 mb-3" />
              <p className="text-sm text-muted-foreground">No flags yet.</p>
              <p className="text-xs text-muted-foreground mt-1">
                Flags are created automatically during AI analysis, or manually from any cell.
              </p>
            </div>
          )}

          {!loading && openFlags.length > 0 && (
            <div className="space-y-2">
              <h3 className="text-xs font-semibold text-muted-foreground uppercase tracking-wide">
                Open ({openFlags.length})
              </h3>
              {openFlags.map(flag => (
                <FlagItem
                  key={flag.id}
                  flag={flag}
                  workbookId={workbookId}
                  onResolved={handleResolved}
                />
              ))}
            </div>
          )}

          {!loading && resolvedFlags.length > 0 && (
            <div className="space-y-2">
              <button
                className="text-xs font-semibold text-muted-foreground uppercase tracking-wide flex items-center gap-1 hover:text-foreground transition-colors"
                onClick={() => setShowResolved(v => !v)}
              >
                {showResolved ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                Resolved ({resolvedFlags.length})
              </button>
              {showResolved && resolvedFlags.map(flag => (
                <FlagItem
                  key={flag.id}
                  flag={flag}
                  workbookId={workbookId}
                  onResolved={handleResolved}
                />
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
