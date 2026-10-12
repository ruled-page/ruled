import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';
import { glidePatches } from './vite.glide.ts';

// Glide's build-time dependencies conflict with @vitejs/plugin-react's optional
// Babel 8 peer, so JSX is compiled by Vite's built-in Oxc transform instead.
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  // The worker's own chunks (each calculation engine is one) are named for it, so the bundle check tells them from the page's.
  worker: { format: 'es', rolldownOptions: { output: { chunkFileNames: 'assets/worker-[name]-[hash].js' } } },
  plugins: [glidePatches()],
  // The grid's markdown cells are never shown, so their parser isn't needed.
  resolve: { alias: [{ find: /^marked$/, replacement: fileURLToPath(new URL('src/editor/shims/marked.ts', import.meta.url)) }] },
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
    // `npm run engine:bench` runs the engine comparison in bench/ instead.
    include: process.env.ENGINE_BENCH ? ['bench/**/*.bench.ts', 'packages/*/bench/**/*.bench.ts'] : ['tests/**/*.test.ts', 'packages/*/tests/**/*.test.ts'],
    setupFiles: ['tests/setup.ts'],
    testTimeout: 60_000,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.{ts,tsx}', 'packages/*/src/**/*.ts'],
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
