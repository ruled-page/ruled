import { fileURLToPath } from 'node:url';
import { defaultClientConditions, defineConfig } from 'vite';
import { glidePatches } from '../../packages/spreadsheet/vite.glide.ts';

/**
 * The editor demo with the sample workbooks, using the editor from its source. `npm run build` builds it into
 * dist; with SITE=1 (as scripts/build-sites.mjs runs it), into the site to
 * publish, dist-sites/demo at the repository's root.
 */
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  // The worker's own chunks (each calculation engine is one) are named for it, so the bundle check tells them from the page's.
  worker: { format: 'es', rolldownOptions: { output: { chunkFileNames: 'assets/worker-[name]-[hash].js' } } },
  plugins: [glidePatches()],
  resolve: {
    conditions: ['source', ...defaultClientConditions],
    // The grid's markdown cells are never shown, so their parser isn't needed.
    alias: [{ find: /^marked$/, replacement: fileURLToPath(new URL('../../packages/spreadsheet/src/editor/shims/marked.ts', import.meta.url)) }],
  },
  // The dev server bundles Glide ahead of time, outside the plugins above.
  optimizeDeps: { rolldownOptions: { plugins: [glidePatches()] } },
  server: { port: 5180 },
  // The editor is one large chunk by design; scripts/check-bundle.mjs holds it to its budget instead.
  build: { chunkSizeWarningLimit: 4096, ...(process.env.SITE ? { outDir: '../../dist-sites/demo', emptyOutDir: true } : {}) },
});
