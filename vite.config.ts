import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {defineConfig} from 'vite';
import {VitePWA} from 'vite-plugin-pwa';

export default defineConfig(({ mode }) => {
  // In development, base is '/'
  // When building for production (GitHub Pages at kenya-one.github.io/chuka-efootball/), base is '/chuka-efootball/'
  const isProd = mode === 'production' || process.env.NODE_ENV === 'production';
  let base = process.env.BASE_URL || (isProd ? '/chuka-efootball/' : '/');
  if (!base.startsWith('/')) base = '/' + base;
  if (!base.endsWith('/')) base = base + '/';

  return {
    base,
    plugins: [
      react(),
      tailwindcss(),
      VitePWA({
        registerType: 'autoUpdate',
        includeAssets: ['favicon.ico', 'apple-touch-icon.png', 'icon.svg', 'logo.jpg'],
        manifest: {
          id: base,
          name: 'Chuka eFootball League',
          short_name: 'Chuka eFoot',
          description: 'Official Chuka eFootball League web app for Knockout and League tournaments.',
          theme_color: '#080c09',
          background_color: '#080c09',
          display: 'standalone',
          orientation: 'portrait',
          start_url: base,
          scope: base,
          icons: [
            {
              src: `${base}pwa-192x192.png`,
              sizes: '192x192',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: `${base}pwa-512x512.png`,
              sizes: '512x512',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: `${base}pwa-maskable-512x512.png`,
              sizes: '512x512',
              type: 'image/png',
              purpose: 'maskable',
            },
          ],
        },
        workbox: {
          globPatterns: ['**/*.{js,css,html,ico,png,svg,jpg,jpeg,woff,woff2}'],
          runtimeCaching: [
            {
              urlPattern: /^https:\/\/fonts\.googleapis\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'google-fonts-cache',
                expiration: {
                  maxEntries: 10,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
            {
              urlPattern: /^https:\/\/fonts\.gstatic\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'gstatic-fonts-cache',
                expiration: {
                  maxEntries: 10,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
          ],
        },
        devOptions: {
          enabled: false,
        },
      }),
      // Automatically copy index.html to 404.html for seamless GitHub Pages routing
      {
        name: 'generate-github-pages-404',
        closeBundle() {
          try {
            const distIndex = path.resolve(__dirname, 'dist', 'index.html');
            const dist404 = path.resolve(__dirname, 'dist', '404.html');
            if (fs.existsSync(distIndex)) {
              fs.copyFileSync(distIndex, dist404);
            }
          } catch (e) {
            console.warn('Could not generate 404.html:', e);
          }
        },
      },
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
