import type { NextConfig } from 'next'

const nextConfig: NextConfig = {
  // Keep these as native Node.js modules — do NOT webpack-bundle them.
  serverExternalPackages: [
    'google-auth-library',
    '@google-cloud/storage',
    'google-gax',
  ],

  // Compress responses
  compress: true,

  // Optimize images
  images: {
    formats: ['image/avif', 'image/webp'],
    minimumCacheTTL: 3600,
  },

  experimental: {
    serverActions: {
      bodySizeLimit: '50mb',
    },
    // Optimize CSS at build time
    optimizeCss: true,
  },

  // Aggressive HTTP caching for static assets (production only — dev needs no-store so HMR works)
  async headers() {
    if (process.env.NODE_ENV !== 'production') return []
    return [
      {
        source: '/_next/static/(.*)',
        headers: [
          { key: 'Cache-Control', value: 'public, max-age=31536000, immutable' },
        ],
      },
    ]
  },
}

export default nextConfig
