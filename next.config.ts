import type { NextConfig } from 'next'

const nextConfig: NextConfig = {
  // Keep these as native Node.js modules — do NOT webpack-bundle them.
  // google-auth-library uses dynamic require() and native addons that break
  // when Next.js tries to bundle them for the server runtime.
  serverExternalPackages: [
    'google-auth-library',
    '@google-cloud/storage',
    'google-gax',
  ],

  experimental: {
    // Allow larger uploads — financial documents (PDFs, Excel) can be 20-50 MB
    serverActions: {
      bodySizeLimit: '50mb',
    },
  },
}

export default nextConfig
