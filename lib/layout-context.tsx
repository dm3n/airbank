'use client'

import { createContext, useCallback, useContext, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { CellReference } from '@/components/auditable-cell'

interface LayoutContextValue {
  chatOpen: boolean
  openChat: () => void
  closeChat: () => void
  cellRef: CellReference | null
  setCellRef: (ref: CellReference | null) => void
  sidebarCollapsed: boolean
  setSidebarCollapsed: (v: boolean) => void
  /** Workbook page registers a callback; layout calls it after a flag resolves */
  flagsRefreshRef: React.MutableRefObject<(() => void) | undefined>
}

const LayoutContext = createContext<LayoutContextValue>({
  chatOpen: false,
  openChat: () => {},
  closeChat: () => {},
  cellRef: null,
  setCellRef: () => {},
  sidebarCollapsed: false,
  setSidebarCollapsed: () => {},
  flagsRefreshRef: { current: undefined },
})

export function LayoutProvider({ children }: { children: ReactNode }) {
  const [chatOpen, setChatOpen] = useState(false)
  const [cellRef, setCellRef] = useState<CellReference | null>(null)
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
  const flagsRefreshRef = useRef<(() => void) | undefined>(undefined)

  const openChat = useCallback(() => setChatOpen(true), [])
  const closeChat = useCallback(() => setChatOpen(false), [])

  return (
    <LayoutContext.Provider value={{
      chatOpen, openChat, closeChat,
      cellRef, setCellRef,
      sidebarCollapsed, setSidebarCollapsed,
      flagsRefreshRef,
    }}>
      {children}
    </LayoutContext.Provider>
  )
}

export function useLayoutContext() {
  return useContext(LayoutContext)
}
