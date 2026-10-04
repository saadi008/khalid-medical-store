import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { VitePWA } from 'vite-plugin-pwa'

export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    VitePWA({
      registerType: 'autoUpdate',
      manifest: {
        id: '/',
        name: 'Khalid Medical Store',
        short_name: 'Khalid Medical Store',
        description: 'Medical Store POS & Cash Management',
        start_url: '/',
        display: 'standalone',
        theme_color: '#2563eb',
        background_color: '#ffffff',
        icons: [
          {
            src: '/km-store-logo.png',
            sizes: '512x512',
            type: 'image/png',
            purpose: 'any',
          },
          {
            src: '/km-store-logo.png',
            sizes: '192x192',
            type: 'image/png',
            purpose: 'any',
          },
        ],
      },
    }),
  ],
})