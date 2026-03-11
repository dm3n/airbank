'use client'

import { useState, useEffect, useCallback } from 'react'
import Link from 'next/link'
import Image from 'next/image'
import { usePathname } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Plus, LayoutDashboard, FileText, ChevronRight, UserCircle2, ChevronsLeft, ChevronsRight } from 'lucide-react'
import { NewWorkbookDialog } from '@/components/new-workbook-dialog'
import { AccountDialog } from '@/components/account-dialog'
import { AIChatPanel } from '@/components/ai-chat-panel'
import { LayoutProvider, useLayoutContext } from '@/lib/layout-context'
import { createBrowserClient } from '@supabase/ssr'

interface Workbook {
  id: string
  company_name: string
  status: string
}

function displayStatus(status: string): 'complete' | 'pending' | 'attention' {
  if (status === 'ready') return 'complete'
  if (status === 'needs_input' || status === 'error') return 'attention'
  return 'pending'
}

const FALLBACK_WORKBOOKS: Workbook[] = [
  { id: 'sandbox', company_name: 'Sandbox — Alpine Outdoor Co.', status: 'ready' },
]

function DashboardLayoutInner({ children }: { children: React.ReactNode }) {
  const pathname = usePathname()
  const { chatOpen, closeChat, cellRef, setCellRef, sidebarCollapsed, setSidebarCollapsed, flagsRefreshRef } = useLayoutContext()
  const [newWorkbookOpen, setNewWorkbookOpen] = useState(false)
  const [accountOpen, setAccountOpen] = useState(false)
  const [workbooks, setWorkbooks] = useState<Workbook[] | null>(null)
  const [userEmail, setUserEmail] = useState<string>('')

  // Extract workbook ID from URL so AI panel can fetch flags for the active workbook
  const workbookId = pathname.match(/\/dashboard\/workbook\/([^/]+)/)?.[1]
  const isDemoWorkbook = workbookId ? /^\d$/.test(workbookId) : false
  const activeWorkbookId = isDemoWorkbook ? undefined : workbookId

  const fetchWorkbooks = useCallback(async () => {
    try {
      const res = await fetch('/api/workbooks')
      if (!res.ok) { setWorkbooks(FALLBACK_WORKBOOKS); return }
      const data: Workbook[] = await res.json()
      const real = Array.isArray(data) ? data : []
      // Always show sandbox first, then any real workbooks
      setWorkbooks([FALLBACK_WORKBOOKS[0], ...real])
    } catch {
      setWorkbooks(FALLBACK_WORKBOOKS)
    }
  }, [])

  useEffect(() => {
    fetchWorkbooks()
    const supabase = createBrowserClient(
      process.env.NEXT_PUBLIC_SUPABASE_URL!,
      process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
    )
    supabase.auth.getUser().then(({ data }) => {
      if (data.user?.email) setUserEmail(data.user.email)
    })
  }, [fetchWorkbooks])

  // Only re-fetch when navigating away from a workbook (e.g. after creation)
  useEffect(() => {
    if (pathname === '/dashboard') fetchWorkbooks()
  }, [pathname, fetchWorkbooks])

  useEffect(() => {
    const hasActive = workbooks?.some(w => w.status === 'analyzing' || w.status === 'uploading')
    if (!hasActive) return
    const id = setInterval(fetchWorkbooks, 30_000)
    return () => clearInterval(id)
  }, [workbooks, fetchWorkbooks])

  const getStatusColor = (status: string) => {
    switch (displayStatus(status)) {
      case 'complete': return 'bg-blue-500'
      case 'attention': return 'bg-gray-400'
      default: return 'bg-gray-300'
    }
  }

  const collapsed = sidebarCollapsed

  return (
    <div className="flex h-screen overflow-hidden">
      {/* ── Sidebar ──────────────────────────────────────────────── */}
      {/*
        Use inline style for width so the browser can composite the
        transition on the GPU via will-change, avoiding layout thrash.
        Both content layers stay mounted; opacity/pointer-events toggle
        between them so there's no DOM churn mid-animation.
      */}
      <div
        className="border-r bg-background flex flex-col flex-shrink-0 relative overflow-hidden"
        style={{
          width: collapsed ? 48 : 256,
          transition: 'width 180ms cubic-bezier(0.4, 0, 0.2, 1)',
          willChange: 'width',
        }}
      >
        {/* ── Collapsed layer ── */}
        <div
          className="absolute inset-0 flex flex-col items-center py-3 gap-1"
          style={{
            opacity: collapsed ? 1 : 0,
            pointerEvents: collapsed ? 'auto' : 'none',
            transition: 'opacity 120ms ease',
          }}
        >
          <button
            title="Expand sidebar"
            onClick={() => setSidebarCollapsed(false)}
            className="p-2 rounded-md hover:bg-accent transition-colors text-muted-foreground hover:text-foreground"
          >
            <ChevronsRight className="h-4 w-4" />
          </button>
          <div className="w-full border-b my-1" />
          <button
            title="New Workbook"
            onClick={() => setNewWorkbookOpen(true)}
            className="p-2 rounded-md hover:bg-accent transition-colors text-muted-foreground hover:text-foreground"
          >
            <Plus className="h-4 w-4" />
          </button>
          <Link href="/dashboard">
            <div
              title="Overview"
              className={`p-2 rounded-md hover:bg-accent transition-colors cursor-pointer ${pathname === '/dashboard' ? 'bg-accent text-foreground' : 'text-muted-foreground hover:text-foreground'}`}
            >
              <LayoutDashboard className="h-4 w-4" />
            </div>
          </Link>
          <div className="w-full border-b my-1" />
          <div className="flex flex-col items-center gap-2 mt-1">
            {workbooks?.map(wb => (
              <Link key={wb.id} href={`/dashboard/workbook/${wb.id}`}>
                <div
                  title={wb.company_name}
                  className={`h-2.5 w-2.5 rounded-full flex-shrink-0 hover:opacity-70 transition-opacity ${
                    pathname === `/dashboard/workbook/${wb.id}` ? 'ring-2 ring-offset-1 ring-gray-400' : ''
                  } ${getStatusColor(wb.status)}`}
                />
              </Link>
            ))}
          </div>
          <div className="flex-1" />
          <button
            title={userEmail || 'Account'}
            onClick={() => setAccountOpen(true)}
            className="p-2 rounded-md hover:bg-accent transition-colors text-muted-foreground hover:text-foreground"
          >
            <UserCircle2 className="h-5 w-5" />
          </button>
        </div>

        {/* ── Full layer ── */}
        <div
          className="flex flex-col flex-1 min-h-0"
          style={{
            opacity: collapsed ? 0 : 1,
            pointerEvents: collapsed ? 'none' : 'auto',
            transition: 'opacity 120ms ease',
          }}
        >
          <div className="p-4 border-b flex items-center justify-between flex-shrink-0">
            <Image src="/logo.png" alt="QoE Platform" width={50} height={50} className="object-contain" />
            <button
              title="Collapse sidebar"
              onClick={() => setSidebarCollapsed(true)}
              className="p-1 rounded-md hover:bg-accent transition-colors text-muted-foreground hover:text-foreground"
            >
              <ChevronsLeft className="h-4 w-4" />
            </button>
          </div>
          <div className="p-4 flex-shrink-0">
            <Button className="w-full" size="sm" onClick={() => setNewWorkbookOpen(true)}>
              <Plus className="mr-2 h-4 w-4" />
              New Workbook
            </Button>
          </div>
          <div className="flex-1 overflow-y-auto min-h-0">
            <div className="px-3 py-2">
              <Link href="/dashboard">
                <div className={`flex items-center px-3 py-2 rounded-md cursor-pointer hover:bg-accent ${pathname === '/dashboard' ? 'bg-accent' : ''}`}>
                  <LayoutDashboard className="mr-2 h-4 w-4" />
                  <span className="text-sm">Overview</span>
                </div>
              </Link>
            </div>
            <div className="px-3 py-2">
              <div className="text-xs font-semibold text-muted-foreground mb-2 px-3">WORKBOOKS</div>
              {workbooks === null ? (
                <div className="px-3 py-2 text-xs text-muted-foreground">Loading...</div>
              ) : workbooks.length === 0 ? (
                <div className="px-3 py-2 text-xs text-muted-foreground">
                  No workbooks yet. Click &ldquo;New Workbook&rdquo; to get started.
                </div>
              ) : (
                workbooks.map(workbook => (
                  <Link key={workbook.id} href={`/dashboard/workbook/${workbook.id}`} prefetch={true}>
                    <div className={`flex items-center justify-between px-3 py-2 rounded-md cursor-pointer hover:bg-accent ${pathname === `/dashboard/workbook/${workbook.id}` ? 'bg-accent' : ''}`}>
                      <div className="flex items-center min-w-0">
                        <FileText className="mr-2 h-4 w-4 flex-shrink-0" />
                        <span className="text-sm truncate">{workbook.company_name}</span>
                      </div>
                      <div className={`h-2 w-2 rounded-full flex-shrink-0 ml-2 ${getStatusColor(workbook.status)}`} />
                    </div>
                  </Link>
                ))
              )}
            </div>
          </div>
          <div className="border-t p-3 flex-shrink-0">
            <button
              onClick={() => setAccountOpen(true)}
              className="w-full flex items-center gap-3 px-2 py-2 rounded-lg hover:bg-accent transition-colors text-left"
            >
              <UserCircle2 className="h-7 w-7 text-muted-foreground flex-shrink-0" />
              <div className="flex-1 min-w-0">
                <p className="text-xs text-muted-foreground truncate">{userEmail || 'Loading...'}</p>
              </div>
              <ChevronRight className="h-4 w-4 text-muted-foreground flex-shrink-0" />
            </button>
          </div>
        </div>
      </div>

      {/* ── Main + AI Panel ──────────────────────────────────────── */}
      <div className="flex flex-1 min-w-0 overflow-hidden">
        <div className="flex-1 overflow-y-auto min-w-0">
          {children}
        </div>

        {/* AI panel — always mounted, slides in/out via transform */}
        <div
          className="flex-shrink-0 overflow-hidden"
          style={{
            width: chatOpen ? 380 : 0,
            transition: 'width 200ms cubic-bezier(0.4, 0, 0.2, 1)',
            willChange: 'width',
          }}
        >
          <AIChatPanel
            workbookId={activeWorkbookId}
            open={chatOpen}
            onClose={closeChat}
            onFlagResolved={() => flagsRefreshRef.current?.()}
            cellRef={cellRef}
            onClearCellRef={() => setCellRef(null)}
          />
        </div>
      </div>

      <NewWorkbookDialog
        open={newWorkbookOpen}
        onOpenChange={setNewWorkbookOpen}
        onWorkbookCreated={fetchWorkbooks}
      />
      <AccountDialog
        open={accountOpen}
        onOpenChange={setAccountOpen}
        workbookCount={workbooks?.length ?? 0}
        userEmail={userEmail}
      />
    </div>
  )
}

export default function DashboardLayout({ children }: { children: React.ReactNode }) {
  return (
    <LayoutProvider>
      <DashboardLayoutInner>{children}</DashboardLayoutInner>
    </LayoutProvider>
  )
}
