'use client'

import { useState } from 'react'
import { useLayoutContext } from '@/lib/layout-context'
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from '@/components/ui/popover'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Eye, Edit3, Check, X, AlertCircle, Flag, Sparkles } from 'lucide-react'

export interface SourceRef {
  documentId: string
  documentName: string
  page: number | null
  excerpt: string
  confidence: number
}

export interface CellFlag {
  id: string
  flag_type: string
  severity: string
  title: string
  resolved_at: string | null
  created_by_ai?: boolean
}

export interface CellReference {
  label: string
  period: string
  displayValue: string
  cellId?: string
}

interface AuditableCellProps {
  value: string | number
  source?: string
  sourceRef?: SourceRef | null
  cellId?: string
  workbookId?: string
  flags?: CellFlag[]
  label?: string
  period?: string
  onSave?: (newValue: string | number) => void
  onViewSource?: (sourceRef: SourceRef) => void
  onFlagCreate?: (flag: { title: string; body: string; flag_type: string }) => void
  onReference?: (ctx: CellReference) => void
  className?: string
  isEditable?: boolean
}

function ConfidenceDot({ confidence }: { confidence: number }) {
  const color =
    confidence >= 0.8
      ? 'bg-blue-500'
      : confidence >= 0.6
      ? 'bg-gray-400'
      : 'bg-gray-500'
  const label =
    confidence >= 0.8 ? 'High confidence' : confidence >= 0.6 ? 'Medium confidence' : 'Low confidence'
  return (
    <span
      title={`${label} (${Math.round(confidence * 100)}%)`}
      className={`inline-block h-2 w-2 rounded-full ml-1 ${color}`}
    />
  )
}

type PopoverTab = 'source' | 'flag-form'

export function AuditableCell({
  value,
  source = 'General Ledger - Account reconciliation',
  sourceRef,
  cellId,
  workbookId,
  flags = [],
  label,
  period,
  onSave,
  onViewSource,
  onFlagCreate,
  onReference,
  className = '',
  isEditable = true,
}: AuditableCellProps) {
  const { openChat, setCellRef, cellRef: activeCellRef } = useLayoutContext()
  const isReferenced = !!(
    activeCellRef && cellId && activeCellRef.cellId === cellId
  )
  const [isOpen, setIsOpen] = useState(false)
  const [isEditing, setIsEditing] = useState(false)
  const [editValue, setEditValue] = useState(value.toString())
  const [isSaving, setIsSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [localDisplay, setLocalDisplay] = useState<string | null>(null)
  const [activeTab, setActiveTab] = useState<PopoverTab>('source')
  const [flagTitle, setFlagTitle] = useState('')
  const [flagBody, setFlagBody] = useState('')
  const [flagType, setFlagType] = useState('needs_review')
  const [isSubmittingFlag, setIsSubmittingFlag] = useState(false)

  const displayedValue = localDisplay ?? value
  const unresolvedFlags = flags.filter(f => !f.resolved_at)
  const hasFlags = unresolvedFlags.length > 0

  const handleSave = async () => {
    setIsSaving(true)
    setSaveError(null)
    try {
      if (cellId && workbookId) {
        const res = await fetch(`/api/workbooks/${workbookId}/cells/${cellId}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ raw_value: editValue }),
        })
        if (!res.ok) {
          const msg = await res.text().catch(() => 'Unknown error')
          setSaveError(`Failed to save — ${msg || 'please try again'}`)
          return
        }
      }
      const parsed = parseFloat(editValue.replace(/[^0-9.-]/g, ''))
      if (!isNaN(parsed)) {
        setLocalDisplay(
          new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 }).format(parsed)
        )
      } else {
        setLocalDisplay(editValue)
      }
      if (onSave) onSave(editValue)
      setIsEditing(false)
      setIsOpen(false)
    } finally {
      setIsSaving(false)
    }
  }

  const handleCancel = () => {
    setEditValue(value.toString())
    setSaveError(null)
    setIsEditing(false)
  }

  const handleViewSource = () => {
    if (sourceRef && onViewSource) {
      onViewSource(sourceRef)
      setIsOpen(false)
    }
  }

  const handleFlagSubmit = async () => {
    if (!flagTitle.trim()) return
    setIsSubmittingFlag(true)
    try {
      if (onFlagCreate) {
        await onFlagCreate({ title: flagTitle.trim(), body: flagBody.trim(), flag_type: flagType })
      } else if (workbookId) {
        const row_key = cellId ?? 'unknown'
        await fetch(`/api/workbooks/${workbookId}/flags`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            section: 'manual',
            row_key,
            flag_type: flagType,
            severity: 'warning',
            title: flagTitle.trim(),
            body: flagBody.trim(),
            cell_id: cellId ?? null,
          }),
        })
      }
      setFlagTitle('')
      setFlagBody('')
      setFlagType('needs_review')
      // Go back to source tab — the parent will have updated flags prop by now
      setActiveTab('source')
    } finally {
      setIsSubmittingFlag(false)
    }
  }

  const handleClose = (open: boolean) => {
    setIsOpen(open)
    if (!open) {
      setIsEditing(false)
      setSaveError(null)
      setActiveTab('source')
    }
  }

  return (
    <Popover open={isOpen} onOpenChange={handleClose}>
      <PopoverTrigger asChild>
        <span
          className={`cursor-pointer transition-colors hover:text-blue-600 ${
            isReferenced
              ? 'text-blue-600 font-semibold underline decoration-blue-300 underline-offset-2'
              : isOpen
              ? 'text-blue-600 font-semibold'
              : hasFlags
              ? 'text-red-500 font-medium'
              : ''
          } ${className}`}
          onClick={() => setIsOpen(true)}
        >
          {displayedValue}
          {sourceRef && <ConfidenceDot confidence={sourceRef.confidence} />}
          {hasFlags && <Flag className="inline-block ml-1 h-3 w-3 text-red-500" />}
        </span>
      </PopoverTrigger>
      <PopoverContent className="w-80" align="start">
        {!isEditing ? (
          <div className="space-y-3">
            {/* Tabs: Source | + Flag */}
            <div className="flex gap-1 border-b pb-2">
              <button
                className={`text-xs px-2 py-1 rounded flex items-center gap-1 ${activeTab === 'source' ? 'bg-muted font-medium' : 'text-muted-foreground hover:text-foreground'}`}
                onClick={() => setActiveTab('source')}
              >
                Source
                {unresolvedFlags.length > 0 && (
                  <span className="bg-red-500 text-white text-[10px] rounded-full px-1 leading-none py-0.5">
                    {unresolvedFlags.length}
                  </span>
                )}
              </button>
              <button
                className={`text-xs px-2 py-1 rounded ${activeTab === 'flag-form' ? 'bg-muted font-medium' : 'text-muted-foreground hover:text-foreground'}`}
                onClick={() => setActiveTab('flag-form')}
              >
                + Flag
              </button>
            </div>

            {activeTab === 'source' && (
              <>
                {/* Inline flags — shown when any exist */}
                {unresolvedFlags.length > 0 && (
                  <div className="rounded-md border border-red-200 bg-red-50/60 p-2.5 space-y-1.5">
                    <div className="flex items-center gap-1.5 text-xs font-semibold text-red-600 mb-1">
                      <Flag className="h-3 w-3" />
                      {unresolvedFlags.length} Open Flag{unresolvedFlags.length !== 1 ? 's' : ''}
                    </div>
                    {unresolvedFlags.map(flag => (
                      <div key={flag.id} className="text-xs text-foreground/80 flex items-start gap-1.5 leading-snug">
                        <span className="text-red-400 mt-px shrink-0">•</span>
                        <span>
                          <span className="font-medium">{flag.title}</span>
                          {' '}
                          <span className="text-muted-foreground capitalize">({flag.flag_type.replace(/_/g, ' ')})</span>
                          {flag.created_by_ai && <span className="ml-1 text-[10px] bg-blue-50 text-blue-600 px-1 rounded">AI</span>}
                        </span>
                      </div>
                    ))}
                  </div>
                )}

                <div>
                  <div className="text-sm font-semibold mb-1">Value</div>
                  <div className="text-lg font-mono">{displayedValue}</div>
                </div>
                <div>
                  <div className="text-sm font-semibold mb-1">Source</div>
                  <div className="text-sm text-muted-foreground">
                    {sourceRef ? (
                      <>
                        <div className="font-medium text-foreground">{sourceRef.documentName}</div>
                        {sourceRef.page && (
                          <div className="text-xs text-muted-foreground">Page {sourceRef.page}</div>
                        )}
                        {sourceRef.excerpt && (
                          <div className="text-xs mt-1 italic text-muted-foreground line-clamp-3">
                            &ldquo;{sourceRef.excerpt}&rdquo;
                          </div>
                        )}
                        <div className="flex items-center gap-1 mt-1">
                          <ConfidenceDot confidence={sourceRef.confidence} />
                          <span className="text-xs text-muted-foreground">
                            {Math.round(sourceRef.confidence * 100)}% confidence
                          </span>
                        </div>
                      </>
                    ) : (
                      source
                    )}
                  </div>
                </div>
                <div className="flex gap-2">
                  <Button
                    variant="outline"
                    size="sm"
                    className="flex-1"
                    onClick={handleViewSource}
                    disabled={!sourceRef || !onViewSource}
                  >
                    <Eye className="mr-2 h-3 w-3" />
                    Source
                  </Button>
                  {isEditable && (
                    <Button
                      variant="outline"
                      size="sm"
                      className="flex-1"
                      onClick={() => { setEditValue(value.toString()); setIsEditing(true) }}
                    >
                      <Edit3 className="mr-2 h-3 w-3" />
                      Edit
                    </Button>
                  )}
                </div>
                <button
                  onClick={() => {
                    setCellRef({
                      label: label ?? source,
                      period: period ?? '',
                      displayValue: displayedValue.toString(),
                      cellId: cellId,
                    })
                    openChat()
                    setIsOpen(false)
                    if (onReference) onReference({ label: label ?? source, period: period ?? '', displayValue: displayedValue.toString(), cellId })
                  }}
                  className="w-full flex items-center justify-center gap-2 rounded-md border border-blue-300/90 bg-blue-50/25 hover:bg-blue-50/60 hover:border-blue-400/80 text-blue-500/90 hover:text-blue-600 text-xs font-medium py-2 transition-colors"
                >
                  <Sparkles className="h-3 w-3" />
                  Ask Workbook AI
                </button>
              </>
            )}

            {activeTab === 'flag-form' && (
              <div className="space-y-3">
                <div className="text-sm font-semibold">Flag for Review</div>
                <div>
                  <label className="text-xs text-muted-foreground mb-1 block">Type</label>
                  <select
                    value={flagType}
                    onChange={e => setFlagType(e.target.value)}
                    className="w-full text-xs border rounded px-2 py-1.5 bg-background"
                  >
                    <option value="needs_review">Needs Review</option>
                    <option value="discrepancy">Discrepancy</option>
                    <option value="low_confidence">Low Confidence</option>
                    <option value="ai_note">AI Note</option>
                  </select>
                </div>
                <div>
                  <label className="text-xs text-muted-foreground mb-1 block">Title</label>
                  <Input
                    value={flagTitle}
                    onChange={e => setFlagTitle(e.target.value)}
                    placeholder="Brief description..."
                    className="text-sm"
                    autoFocus
                    onKeyDown={e => { if (e.key === 'Enter' && flagTitle.trim()) handleFlagSubmit() }}
                  />
                </div>
                <div>
                  <label className="text-xs text-muted-foreground mb-1 block">Comment (optional)</label>
                  <Textarea
                    value={flagBody}
                    onChange={e => setFlagBody(e.target.value)}
                    placeholder="Additional context..."
                    className="text-sm resize-none"
                    rows={2}
                  />
                </div>
                <div className="flex gap-2">
                  <Button variant="outline" size="sm" className="flex-1" onClick={() => setActiveTab('source')}>
                    Cancel
                  </Button>
                  <Button
                    size="sm"
                    className="flex-1"
                    onClick={handleFlagSubmit}
                    disabled={!flagTitle.trim() || isSubmittingFlag}
                  >
                    <Flag className="mr-2 h-3 w-3" />
                    {isSubmittingFlag ? 'Saving…' : 'Flag'}
                  </Button>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="space-y-3">
            <div>
              <div className="text-sm font-semibold mb-2">Edit Value</div>
              <Input
                value={editValue}
                onChange={(e) => { setEditValue(e.target.value); setSaveError(null) }}
                className="font-mono"
                autoFocus
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleSave()
                  if (e.key === 'Escape') handleCancel()
                }}
              />
            </div>
            {saveError && (
              <div className="flex items-center gap-1.5 text-xs text-gray-600">
                <AlertCircle className="h-3 w-3 flex-shrink-0" />
                {saveError}
              </div>
            )}
            <div className="flex gap-2">
              <Button variant="outline" size="sm" className="flex-1" onClick={handleCancel} disabled={isSaving}>
                <X className="mr-2 h-3 w-3" />
                Cancel
              </Button>
              <Button size="sm" className="flex-1" onClick={handleSave} disabled={isSaving}>
                <Check className="mr-2 h-3 w-3" />
                {isSaving ? 'Saving…' : 'Save'}
              </Button>
            </div>
          </div>
        )}
      </PopoverContent>
    </Popover>
  )
}
