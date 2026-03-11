'use client'

import { useState, useEffect, useCallback, useRef } from 'react'
import { Button } from '@/components/ui/button'
import { Textarea } from '@/components/ui/textarea'
import {
  X, Check, Loader2, ChevronDown, ChevronUp, Send,
  Sparkles, Flag, MessageSquare,
} from 'lucide-react'

export interface CellReference {
  label: string
  period: string
  displayValue: string
  cellId?: string
}

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

interface ChatMsg {
  id: string
  role: 'user' | 'ai'
  content: string
  cellRef?: CellReference
  error?: boolean
}

interface AIChatPanelProps {
  workbookId?: string
  open: boolean
  onClose: () => void
  onFlagResolved?: () => void
  cellRef?: CellReference | null
  onClearCellRef?: () => void
}

const FLAG_TYPE_LABELS: Record<string, string> = {
  low_confidence: 'Low Confidence',
  discrepancy: 'Discrepancy',
  missing: 'Missing Data',
  ai_note: 'AI Note',
  needs_review: 'Needs Review',
}

const SEVERITY_COLOR: Record<string, string> = {
  critical: 'bg-red-500',
  warning: 'bg-amber-400',
  info: 'bg-gray-300',
}

// ─── Flag thread ──────────────────────────────────────────────────────────────

function FlagThread({
  flag, workbookId, onResolved,
}: { flag: CellFlag; workbookId?: string; onResolved: () => void }) {
  const [expanded, setExpanded] = useState(false)
  const [comments, setComments] = useState<FlagComment[]>([])
  const [commentsLoaded, setCommentsLoaded] = useState(false)
  const [newComment, setNewComment] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [resolving, setResolving] = useState(false)

  const loadComments = useCallback(async () => {
    if (commentsLoaded || !workbookId) return
    const res = await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}/comments`)
    if (res.ok) { setComments(await res.json()); setCommentsLoaded(true) }
  }, [workbookId, flag.id, commentsLoaded])

  const handleExpand = async () => { const next = !expanded; setExpanded(next); if (next) await loadComments() }

  const handleResolve = async () => {
    if (!workbookId) return
    setResolving(true)
    try {
      await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ resolved: !flag.resolved_at }),
      })
      onResolved()
    } finally { setResolving(false) }
  }

  const handleAddComment = async () => {
    if (!newComment.trim() || !workbookId) return
    setSubmitting(true)
    try {
      const res = await fetch(`/api/workbooks/${workbookId}/flags/${flag.id}/comments`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ body: newComment.trim() }),
      })
      if (res.ok) { const c = await res.json(); setComments(prev => [...prev, c]); setNewComment('') }
    } finally { setSubmitting(false) }
  }

  const isResolved = !!flag.resolved_at
  const dot = SEVERITY_COLOR[flag.severity] ?? 'bg-gray-400'

  return (
    <div className={`transition-opacity ${isResolved ? 'opacity-40' : ''}`}>
      <div className="flex gap-2.5">
        <div className="flex flex-col items-center pt-1.5 flex-shrink-0">
          <div className={`w-2 h-2 rounded-full flex-shrink-0 ${flag.created_by_ai ? 'bg-blue-500' : dot}`} />
          {expanded && <div className="w-px flex-1 bg-border mt-1.5 min-h-[12px]" />}
        </div>
        <div className="flex-1 min-w-0 pb-3.5">
          <div className="flex items-start justify-between gap-2">
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-1.5 mb-0.5 flex-wrap">
                {flag.created_by_ai && (
                  <span className="text-[10px] font-semibold text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded">AI</span>
                )}
                <span className="text-[10px] text-muted-foreground">{FLAG_TYPE_LABELS[flag.flag_type] ?? flag.flag_type}</span>
                {flag.period && <span className="text-[10px] text-muted-foreground">· {flag.period}</span>}
              </div>
              <p className="text-sm font-medium leading-snug">{flag.title}</p>
              {flag.body && !expanded && (
                <p className="text-xs text-muted-foreground mt-0.5 leading-relaxed line-clamp-2">{flag.body}</p>
              )}
            </div>
            <div className="flex items-center gap-1 flex-shrink-0 mt-0.5">
              <button
                onClick={handleResolve}
                disabled={resolving}
                className={`h-6 px-2 rounded text-xs border transition-colors flex items-center gap-1 ${
                  isResolved
                    ? 'border-border text-muted-foreground hover:text-foreground'
                    : 'border-border text-muted-foreground hover:border-blue-400 hover:text-blue-600'
                }`}
              >
                {resolving ? <Loader2 className="h-3 w-3 animate-spin" />
                  : isResolved ? <><X className="h-2.5 w-2.5" />Reopen</>
                  : <><Check className="h-2.5 w-2.5" />Resolve</>}
              </button>
              <button onClick={handleExpand} className="text-muted-foreground hover:text-foreground transition-colors p-0.5">
                {expanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
              </button>
            </div>
          </div>

          {expanded && (
            <div className="mt-3 space-y-3">
              {flag.body && <p className="text-xs text-muted-foreground leading-relaxed">{flag.body}</p>}
              {comments.map(c => (
                <div key={c.id} className={`flex ${c.author_name === 'AI' ? 'justify-start' : 'justify-end'}`}>
                  <div className={`max-w-[85%] rounded-xl px-3 py-2 text-xs leading-relaxed ${
                    c.author_name === 'AI' ? 'bg-muted text-foreground' : 'bg-blue-500 text-white'
                  }`}>
                    {c.author_name !== 'AI' && <p className="font-medium text-[10px] mb-0.5 opacity-70">{c.author_name}</p>}
                    <p>{c.body}</p>
                  </div>
                </div>
              ))}
              {commentsLoaded && comments.length === 0 && !flag.body && (
                <p className="text-xs text-muted-foreground italic">No details yet.</p>
              )}
              <div className="flex gap-2">
                <input
                  type="text"
                  value={newComment}
                  onChange={e => setNewComment(e.target.value)}
                  onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); handleAddComment() } }}
                  placeholder="Reply…"
                  className="flex-1 text-xs rounded-lg border bg-background px-3 py-1.5 outline-none focus:border-blue-400 transition-colors"
                />
                <button
                  onClick={handleAddComment}
                  disabled={!newComment.trim() || submitting}
                  className="text-blue-500 disabled:text-muted-foreground hover:text-blue-600 transition-colors"
                >
                  {submitting ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Send className="h-3.5 w-3.5" />}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// ─── Main panel ───────────────────────────────────────────────────────────────

export function AIChatPanel({
  workbookId, open, onClose, onFlagResolved, cellRef, onClearCellRef,
}: AIChatPanelProps) {
  const [tab, setTab] = useState<'chat' | 'flags'>('chat')
  const [messages, setMessages] = useState<ChatMsg[]>([])
  const [flags, setFlags] = useState<CellFlag[]>([])
  const [flagsLoading, setFlagsLoading] = useState(false)
  const [showResolved, setShowResolved] = useState(false)
  const [message, setMessage] = useState('')
  const [sending, setSending] = useState(false)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const bottomRef = useRef<HTMLDivElement>(null)

  // Chat history for multi-turn AI context
  const historyRef = useRef<Array<{ role: 'user' | 'model'; content: string }>>([])

  const fetchFlags = useCallback(async () => {
    if (!workbookId) return
    setFlagsLoading(true)
    try {
      const res = await fetch(`/api/workbooks/${workbookId}/flags`)
      if (res.ok) setFlags(Array.isArray(await res.json()) ? await res.json() : [])
    } finally { setFlagsLoading(false) }
  }, [workbookId])

  // Re-fetch flags and refetch correctly (avoid double-await)
  const refreshFlags = useCallback(async () => {
    if (!workbookId) return
    const res = await fetch(`/api/workbooks/${workbookId}/flags`)
    if (res.ok) {
      const data = await res.json()
      setFlags(Array.isArray(data) ? data : [])
    }
  }, [workbookId])

  useEffect(() => {
    if (open && workbookId) {
      refreshFlags()
      // Greet with a context-aware opener on first open
      if (messages.length === 0) {
        setMessages([{
          id: 'welcome',
          role: 'ai',
          content: 'Hi! I\'m your Workbook AI. I have full context of this workbook — all financial data, sections, and flags. Ask me anything, walk through flags, or click any metric in the report to reference it directly.',
        }])
      }
    }
  }, [open, workbookId, refreshFlags])

  useEffect(() => {
    if (cellRef && open) inputRef.current?.focus()
  }, [cellRef, open])

  const handleResolved = useCallback(() => {
    refreshFlags()
    onFlagResolved?.()
  }, [refreshFlags, onFlagResolved])

  const scrollToBottom = () =>
    setTimeout(() => bottomRef.current?.scrollIntoView({ behavior: 'smooth' }), 80)

  const handleSend = async () => {
    if (!message.trim() || !workbookId) return
    const userText = message.trim()
    const msgCellRef = cellRef ?? undefined

    // Optimistically add user message
    const userMsg: ChatMsg = {
      id: Date.now().toString(),
      role: 'user',
      content: userText,
      cellRef: msgCellRef,
    }
    setMessages(prev => [...prev, userMsg])
    setMessage('')
    onClearCellRef?.()
    setSending(true)
    scrollToBottom()

    // Add thinking placeholder
    const thinkingId = `thinking-${Date.now()}`
    setMessages(prev => [...prev, { id: thinkingId, role: 'ai', content: '…' }])

    try {
      const res = await fetch(`/api/workbooks/${workbookId}/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: userText,
          history: historyRef.current,
          cellRef: msgCellRef,
        }),
      })

      const data = await res.json()
      const aiText: string = res.ok
        ? (data.response ?? 'No response.')
        : (data.error ?? 'Something went wrong.')

      // Update history for multi-turn context
      historyRef.current = [
        ...historyRef.current,
        { role: 'user' as const, content: userText },
        { role: 'model' as const, content: aiText },
      ].slice(-20) // keep last 10 turns

      setMessages(prev =>
        prev.map(m => m.id === thinkingId
          ? { id: thinkingId, role: 'ai', content: aiText, error: !res.ok }
          : m
        )
      )

      // Refresh flags in case AI analysis surfaced new ones
      await refreshFlags()
    } catch {
      setMessages(prev =>
        prev.map(m => m.id === thinkingId
          ? { ...m, content: 'Connection error. Please try again.', error: true }
          : m
        )
      )
    } finally {
      setSending(false)
      scrollToBottom()
    }
  }

  const openFlags = flags.filter(f => !f.resolved_at)
  const resolvedFlags = flags.filter(f => f.resolved_at)

  return (
    <div className="w-[380px] flex-shrink-0 border-l bg-background flex flex-col h-full overflow-hidden">

      {/* Header */}
      <div className="flex items-center justify-between px-5 py-3.5 border-b flex-shrink-0">
        <div className="flex items-center gap-2">
          <Sparkles className="h-4 w-4 text-blue-500" />
          <span className="text-base font-semibold">Workbook AI</span>
        </div>
        <div className="flex items-center gap-2">
          {/* Tab switcher */}
          <div className="flex rounded-md border overflow-hidden text-xs">
            <button
              onClick={() => setTab('chat')}
              className={`px-3 py-1 flex items-center gap-1.5 transition-colors ${tab === 'chat' ? 'bg-foreground text-background' : 'text-muted-foreground hover:text-foreground'}`}
            >
              <MessageSquare className="h-3 w-3" />
              Chat
            </button>
            <button
              onClick={() => setTab('flags')}
              className={`px-3 py-1 flex items-center gap-1.5 transition-colors border-l ${tab === 'flags' ? 'bg-foreground text-background' : 'text-muted-foreground hover:text-foreground'}`}
            >
              <Flag className="h-3 w-3" />
              Flags
              {openFlags.length > 0 && (
                <span className={`text-[9px] font-bold rounded-full px-1.5 py-0.5 leading-none ${tab === 'flags' ? 'bg-background text-foreground' : 'bg-red-500 text-white'}`}>
                  {openFlags.length}
                </span>
              )}
            </button>
          </div>
          <button onClick={onClose} className="text-muted-foreground hover:text-foreground transition-colors">
            <X className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* ── Chat tab ── */}
      {tab === 'chat' && (
        <>
          <div className="flex-1 overflow-y-auto px-4 py-4 space-y-3 min-h-0">
            {messages.map(msg => (
              <div key={msg.id} className={`flex flex-col gap-1 ${msg.role === 'user' ? 'items-end' : 'items-start'}`}>
                {/* Cell ref chip on user messages */}
                {msg.role === 'user' && msg.cellRef && (
                  <div className="flex items-center gap-1.5 text-[10px] text-blue-600 bg-blue-50 border border-blue-200 rounded-md px-2 py-1 max-w-[85%]">
                    <Sparkles className="h-2.5 w-2.5 flex-shrink-0" />
                    <span className="truncate font-medium">{msg.cellRef.label}</span>
                    <span className="opacity-70">· {msg.cellRef.period}</span>
                    <span className="font-mono opacity-70">{msg.cellRef.displayValue}</span>
                  </div>
                )}
                <div className={`max-w-[88%] rounded-2xl px-3.5 py-2.5 text-sm leading-relaxed ${
                  msg.role === 'user'
                    ? 'bg-blue-600 text-white rounded-br-sm'
                    : msg.error
                    ? 'bg-red-50 text-red-700 border border-red-200 rounded-bl-sm'
                    : 'bg-muted text-foreground rounded-bl-sm'
                }`}>
                  {msg.content === '…' ? (
                    <span className="flex items-center gap-1.5 text-muted-foreground">
                      <Loader2 className="h-3 w-3 animate-spin" />
                      <span className="text-xs">Thinking…</span>
                    </span>
                  ) : (
                    <p className="whitespace-pre-wrap">{msg.content}</p>
                  )}
                </div>
              </div>
            ))}
            <div ref={bottomRef} />
          </div>

          {/* Input */}
          <div className="border-t px-4 pt-3 pb-4 flex-shrink-0 space-y-2.5">
            {cellRef ? (
              <div className="rounded-lg border border-blue-200 bg-blue-50 dark:bg-blue-950/30 dark:border-blue-800 overflow-hidden">
                <div className="flex items-center justify-between px-3 pt-2 pb-1">
                  <div className="flex items-center gap-1.5">
                    <Sparkles className="h-3 w-3 text-blue-500" />
                    <span className="text-[10px] font-semibold text-blue-500 uppercase tracking-wider">Referencing</span>
                  </div>
                  <button onClick={onClearCellRef} className="text-blue-400 hover:text-blue-600 transition-colors">
                    <X className="h-3 w-3" />
                  </button>
                </div>
                <div className="px-3 pb-2.5 flex items-baseline justify-between gap-3">
                  <span className="text-sm font-semibold text-blue-900 dark:text-blue-100 truncate leading-tight">{cellRef.label}</span>
                  <div className="flex items-center gap-2 flex-shrink-0">
                    <span className="text-[11px] text-blue-600 dark:text-blue-400 font-medium">{cellRef.period}</span>
                    <span className="text-sm font-mono font-semibold text-blue-800 dark:text-blue-200">{cellRef.displayValue}</span>
                  </div>
                </div>
              </div>
            ) : (
              <p className="text-[10px] text-muted-foreground">Click any metric in the report to bring it into context.</p>
            )}
            <div className="flex items-end gap-2">
              <Textarea
                ref={inputRef}
                value={message}
                onChange={e => setMessage(e.target.value)}
                onKeyDown={e => {
                  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend() }
                }}
                placeholder={cellRef ? 'Ask a question about this metric…' : 'Ask a question or add a flag…'}
                className="resize-none text-sm min-h-[60px] max-h-[120px]"
                rows={2}
                disabled={!workbookId}
              />
              <Button
                size="icon"
                onClick={handleSend}
                disabled={!message.trim() || sending || !workbookId}
                className="h-9 w-9 flex-shrink-0 mb-0.5"
              >
                {sending ? <Loader2 className="h-4 w-4 animate-spin" /> : <Send className="h-4 w-4" />}
              </Button>
            </div>
          </div>
        </>
      )}

      {/* ── Flags tab ── */}
      {tab === 'flags' && (
        <div className="flex-1 overflow-y-auto px-4 py-4 min-h-0">
          {flagsLoading && (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="h-5 w-5 animate-spin text-muted-foreground" />
            </div>
          )}

          {!flagsLoading && flags.length === 0 && (
            <div className="flex flex-col items-center justify-center py-16 text-center">
              <div className="w-10 h-10 rounded-full bg-muted flex items-center justify-center mb-3">
                <Flag className="h-5 w-5 text-muted-foreground" />
              </div>
              <p className="text-sm font-medium">No flags yet</p>
              <p className="text-xs text-muted-foreground mt-1 max-w-[220px] leading-relaxed">
                Flags are created during AI analysis or when you flag a metric manually.
              </p>
            </div>
          )}

          {!flagsLoading && openFlags.length > 0 && (
            <div className="mb-2">
              <p className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider mb-3">
                Open · {openFlags.length}
              </p>
              {openFlags.map(flag => (
                <FlagThread key={flag.id} flag={flag} workbookId={workbookId} onResolved={handleResolved} />
              ))}
            </div>
          )}

          {!flagsLoading && resolvedFlags.length > 0 && (
            <div className="mt-2">
              <button
                className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider flex items-center gap-1 hover:text-foreground transition-colors mb-3"
                onClick={() => setShowResolved(v => !v)}
              >
                {showResolved ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                Resolved · {resolvedFlags.length}
              </button>
              {showResolved && resolvedFlags.map(flag => (
                <FlagThread key={flag.id} flag={flag} workbookId={workbookId} onResolved={handleResolved} />
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
