'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { createBrowserClient } from '@supabase/ssr'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { Progress } from '@/components/ui/progress'
import { Separator } from '@/components/ui/separator'
import { CreditCard, LogOut, Check, ShieldCheck, UserCircle2 } from 'lucide-react'

interface AccountDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  workbookCount?: number
  userEmail?: string
}

const PLAN_FEATURES = [
  '10 workbooks',
  'Unlimited exports',
  'Priority support',
  'Advanced analytics',
]

export function AccountDialog({ open, onOpenChange, workbookCount = 0, userEmail = '' }: AccountDialogProps) {
  const router = useRouter()
  const [signingOut, setSigningOut] = useState(false)

  const workbookLimit = 10
  const plan = 'Pro'
  const usagePct = (workbookCount / workbookLimit) * 100

  const handleSignOut = async () => {
    setSigningOut(true)
    const supabase = createBrowserClient(
      process.env.NEXT_PUBLIC_SUPABASE_URL!,
      process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
    )
    await supabase.auth.signOut()
    onOpenChange(false)
    router.push('/')
    router.refresh()
  }

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md p-0 gap-0 overflow-hidden">
        {/* Header */}
        <div className="px-6 pt-6 pb-4">
          <DialogHeader>
            <div className="flex items-center gap-4">
              <div className="h-14 w-14 rounded-full bg-muted flex items-center justify-center flex-shrink-0">
                <UserCircle2 className="h-10 w-10 text-muted-foreground" />
              </div>
              <div className="min-w-0">
                <DialogTitle className="text-lg leading-tight">{userEmail || 'Account'}</DialogTitle>
                <Badge variant="secondary" className="mt-1 text-xs text-blue-500">{plan}</Badge>
              </div>
            </div>
          </DialogHeader>
        </div>

        <Separator />

        <div className="px-6 py-4">
          <Tabs defaultValue="billing">
            <TabsList className="w-full mb-5">
              <TabsTrigger value="billing" className="flex-1">Billing</TabsTrigger>
              <TabsTrigger value="security" className="flex-1">Security</TabsTrigger>
            </TabsList>

            {/* ── Billing Tab ── */}
            <TabsContent value="billing" className="space-y-4 mt-0">
              <div className="rounded-xl border bg-card p-4 space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="font-medium text-sm">Current plan</p>
                    <p className="text-xs text-muted-foreground mt-0.5">Billed monthly · $49/mo</p>
                  </div>
                  <Badge variant="secondary" className="text-blue-500 px-3">
                    {plan}
                  </Badge>
                </div>
                <Separator />
                <div className="space-y-2">
                  <div className="flex justify-between text-sm">
                    <span className="text-muted-foreground">Workbooks used</span>
                    <span className="font-medium">{workbookCount} / {workbookLimit}</span>
                  </div>
                  <Progress value={usagePct} className="h-1.5" />
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">Next billing date</span>
                  <span className="font-medium">Apr 1, 2026</span>
                </div>
              </div>

              <div className="space-y-2">
                <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Plan includes</p>
                <ul className="space-y-1.5">
                  {PLAN_FEATURES.map((f) => (
                    <li key={f} className="flex items-center gap-2 text-sm">
                      <Check className="h-3.5 w-3.5 text-green-500 flex-shrink-0" />
                      {f}
                    </li>
                  ))}
                </ul>
              </div>

              <Button variant="outline" className="w-full">
                <CreditCard className="mr-2 h-4 w-4" />
                Manage billing
              </Button>
            </TabsContent>

            {/* ── Security Tab ── */}
            <TabsContent value="security" className="space-y-4 mt-0">
              <div className="space-y-3">
                <div className="space-y-1.5">
                  <Label>Current password</Label>
                  <Input type="password" placeholder="••••••••" />
                </div>
                <div className="space-y-1.5">
                  <Label>New password</Label>
                  <Input type="password" placeholder="••••••••" />
                </div>
                <div className="space-y-1.5">
                  <Label>Confirm new password</Label>
                  <Input type="password" placeholder="••••••••" />
                </div>
                <Button className="w-full">Update password</Button>
              </div>

              <Separator />

              <div className="flex items-center justify-between py-1">
                <div className="flex items-center gap-3">
                  <div className="h-8 w-8 rounded-lg bg-muted flex items-center justify-center">
                    <ShieldCheck className="h-4 w-4 text-muted-foreground" />
                  </div>
                  <div>
                    <p className="text-sm font-medium">Two-factor authentication</p>
                    <p className="text-xs text-muted-foreground">Add an extra layer of security</p>
                  </div>
                </div>
                <Button variant="outline" size="sm">Enable</Button>
              </div>
            </TabsContent>
          </Tabs>
        </div>

        <Separator />

        {/* Sign out */}
        <div className="px-6 py-3">
          <Button
            variant="ghost"
            className="w-full justify-start text-muted-foreground hover:text-destructive gap-2"
            onClick={handleSignOut}
            disabled={signingOut}
          >
            <LogOut className="h-4 w-4" />
            {signingOut ? 'Signing out...' : 'Sign out'}
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  )
}
