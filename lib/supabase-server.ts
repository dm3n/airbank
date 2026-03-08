import { createServerClient } from '@supabase/ssr'
import { createClient } from '@supabase/supabase-js'
import { cookies, headers } from 'next/headers'

/**
 * Server-side Supabase client.
 * - Reads session from cookies (browser flow via createBrowserClient).
 * - Falls back to Authorization: Bearer <token> header for API clients.
 */
export async function createSupabaseServerClient() {
  const cookieStore = await cookies()
  const headersList = await headers()
  const authHeader = headersList.get('authorization')

  // If an Authorization: Bearer token is present, use it directly.
  // This supports API clients and server-to-server calls.
  if (authHeader?.startsWith('Bearer ')) {
    const token = authHeader.slice(7)
    return createClient(
      process.env.NEXT_PUBLIC_SUPABASE_URL!,
      process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
      {
        auth: { autoRefreshToken: false, persistSession: false },
        global: { headers: { Authorization: `Bearer ${token}` } },
      }
    )
  }

  // Default: cookie-based session (browser flow)
  return createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        getAll() {
          return cookieStore.getAll()
        },
        setAll(cookiesToSet) {
          try {
            cookiesToSet.forEach(({ name, value, options }) =>
              cookieStore.set(name, value, options)
            )
          } catch {
            // Server component — cookie writes are best-effort
          }
        },
      },
    }
  )
}

/** Service-role client: bypasses RLS — only use server-side */
export function createSupabaseServiceClient() {
  return createClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.SUPABASE_SERVICE_ROLE_KEY!,
    { auth: { autoRefreshToken: false, persistSession: false } }
  )
}
