'use client'

import { useState, useEffect } from 'react'
import Link from 'next/link'
import { FileText, ChevronRight, Loader2, CheckCircle2, AlertCircle, Plus } from 'lucide-react'
import { Button } from '@/components/ui/button'

interface Workbook {
  id: string
  company_name: string
  status: string
  periods: string[]
  created_at: string
  updated_at: string
}

function StatusBadge({ status }: { status: string }) {
  if (status === 'ready') {
    return (
      <span className="inline-flex items-center gap-1.5 text-xs font-medium text-blue-600">
        <CheckCircle2 className="h-3.5 w-3.5" />
        Complete
      </span>
    )
  }
  return (
    <span className="inline-flex items-center gap-1.5 text-xs font-medium text-muted-foreground">
      <Loader2 className="h-3.5 w-3.5 animate-spin" />
      In review
    </span>
  )
}

function relativeTime(iso: string): string {
  const diff = Date.now() - new Date(iso).getTime()
  const mins = Math.floor(diff / 60000)
  if (mins < 1) return 'just now'
  if (mins < 60) return `${mins}m ago`
  const hrs = Math.floor(mins / 60)
  if (hrs < 24) return `${hrs}h ago`
  const days = Math.floor(hrs / 24)
  if (days < 7) return `${days}d ago`
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
}

const SANDBOX: Workbook = {
  id: 'sandbox',
  company_name: 'Sandbox — Alpine Outdoor Co.',
  status: 'ready',
  periods: ['FY20', 'FY21', 'FY22', 'TTM'],
  created_at: '2024-01-01T00:00:00Z',
  updated_at: new Date().toISOString(),
}

export default function DashboardOverview() {
  const [workbooks, setWorkbooks] = useState<Workbook[] | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => { document.title = 'Airbank - Overview' }, [])

  useEffect(() => {
    fetch('/api/workbooks')
      .then(async (res) => {
        if (!res.ok) throw new Error(`${res.status}`)
        return res.json()
      })
      .then((data) => setWorkbooks([SANDBOX, ...(Array.isArray(data) ? data : [])]))
      .catch((e) => { setWorkbooks([SANDBOX]); setError(e.message) })
  }, [])

  return (
    <div className="h-full flex flex-col px-10 py-8">
      {/* Header */}
      <div className="mb-8">
        <h1 className="text-xl font-semibold tracking-tight">Overview</h1>
        <p className="text-sm text-muted-foreground mt-0.5">All Quality of Earnings workbooks</p>
      </div>

      {/* Loading skeletons */}
      {workbooks === null && !error && (
        <div className="space-y-px">
          {[1, 2, 3].map((i) => (
            <div key={i} className="h-14 bg-muted/30 animate-pulse rounded-md" />
          ))}
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="flex flex-col items-center justify-center py-20 text-center">
          <AlertCircle className="h-7 w-7 text-muted-foreground/40 mb-3" />
          <p className="text-sm text-muted-foreground">Could not load workbooks</p>
        </div>
      )}

      {/* Empty state */}
      {workbooks !== null && workbooks.length === 0 && (
        <div className="flex flex-col items-center justify-center flex-1 text-center">
          <FileText className="h-9 w-9 text-muted-foreground/25 mb-4" />
          <h3 className="text-sm font-medium mb-1">No workbooks yet</h3>
          <p className="text-xs text-muted-foreground mb-5 max-w-xs">
            Create your first Quality of Earnings workbook to get started.
          </p>
          <Button size="sm" variant="outline">
            <Plus className="mr-2 h-3.5 w-3.5" />
            New Workbook
          </Button>
        </div>
      )}

      {/* Table */}
      {workbooks && workbooks.length > 0 && (
        <div className="flex-1 flex flex-col min-h-0">
          {/* Column headers */}
          <div className="grid grid-cols-[1fr_140px_180px_110px_28px] gap-6 px-4 pb-2 border-b">
            <span className="text-[11px] font-medium text-muted-foreground/60 uppercase tracking-widest">Company</span>
            <span className="text-[11px] font-medium text-muted-foreground/60 uppercase tracking-widest">Status</span>
            <span className="text-[11px] font-medium text-muted-foreground/60 uppercase tracking-widest">Periods</span>
            <span className="text-[11px] font-medium text-muted-foreground/60 uppercase tracking-widest">Updated</span>
            <span />
          </div>

          {/* Rows */}
          <div className="divide-y divide-border/50">
            {workbooks.map((wb) => (
              <Link key={wb.id} href={`/dashboard/workbook/${wb.id}`}>
                <div className="grid grid-cols-[1fr_140px_180px_110px_28px] gap-6 items-center px-4 py-3.5 hover:bg-muted/20 transition-colors group cursor-pointer">
                  {/* Company */}
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="h-7 w-7 rounded bg-muted flex items-center justify-center flex-shrink-0">
                      <FileText className="h-3.5 w-3.5 text-muted-foreground/50" />
                    </div>
                    <span className="text-sm font-medium truncate">{wb.company_name}</span>
                  </div>

                  {/* Status */}
                  <StatusBadge status={wb.status} />

                  {/* Periods */}
                  <div className="flex gap-1 flex-wrap">
                    {(wb.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']).map((p) => (
                      <span key={p} className="text-[11px] bg-muted/60 text-muted-foreground px-1.5 py-0.5 rounded font-mono">
                        {p}
                      </span>
                    ))}
                  </div>

                  {/* Updated */}
                  <span className="text-xs text-muted-foreground/60">
                    {relativeTime(wb.updated_at || wb.created_at)}
                  </span>

                  {/* Arrow */}
                  <ChevronRight className="h-3.5 w-3.5 text-muted-foreground/25 group-hover:text-muted-foreground/50 transition-colors" />
                </div>
              </Link>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}
