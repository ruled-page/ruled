import { defineConfig } from 'vitest/config';
import { glidePatches } from './vite.glide.ts';

// Glide's build-time dependencies conflict with @vitejs/plugin-react's optional
// Babel 8 peer, so JSX is compiled by Vite's built-in Oxc transform instead.
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  worker: { format: 'es' },
  plugins: [glidePatches()],
  // The dev server bundles Glide ahead of time, outside the plugins above.
  optimizeDeps: { rolldownOptions: { plugins: [glidePatches()] } },
  // Two pages: the editor demo, and the editor embedded in a host app. Each is
  // also built on its own as a site to publish (SITE=demo or SITE=portal).
  build:
    process.env.SITE === 'demo' || process.env.SITE === 'portal'
      ? { outDir: `dist-sites/${process.env.SITE}`, emptyOutDir: true, rolldownOptions: { input: process.env.SITE === 'demo' ? 'index.html' : 'embedded.html' } }
      : { rolldownOptions: { input: { main: 'index.html', embedded: 'embedded.html' } } },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    testTimeout: 60_000,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.{ts,tsx}'],
      // The React screens and the apps' entry points are tested in a browser by the Playwright suite (e2e/).
      exclude: [
        'src/ui/**/*.tsx',
        'src/ui/useDialog.ts',
        'src/editor/{WorkbookEditor.tsx,ErrorBoundary.tsx,mount.tsx,element.ts,standalone.ts,index.ts}',
        'src/demo/{main.tsx,Welcome.tsx,i18n.ts}',
        'src/embedded/{*.tsx,i18n.ts,format.ts}',
        'src/{demo,embedded}/locales/**',
        'src/i18n/messages/**',
        'src/sites.ts',
      ],
      reporter: ['text-summary', 'html'],
      // Just under what the suite reaches, so coverage can't quietly slip.
      thresholds: { lines: 98, statements: 97, functions: 97, branches: 90 },
    },
  },
});
