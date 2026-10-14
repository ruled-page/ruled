import { defaultClientConditions, defineConfig } from 'vite';

/**
 * The editor embedded in a React supplier portal, using the editor from its source. `npm run build` builds it into
 * dist; with SITE=1 (as scripts/build-sites.mjs runs it), into the site to
 * publish, dist-sites/portal at the repository's root.
 */
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  // The worker in one file, the engine inlined, as the library builds it. Split, its engine chunk imported
  // the worker's entry for their shared code, which WebKit runs a second time: that copy took the
  // messages after the workbook opened in the first, and every edit failed with "No workbook is open".
  worker: { format: 'es', rolldownOptions: { output: { codeSplitting: false } } },
  resolve: { conditions: ['source', ...defaultClientConditions] },
  // The supplier workbooks are among the demo's samples.
  publicDir: '../demo/public',
  server: { port: 5181 },
  // The editor is one large chunk by design; scripts/check-bundle.mjs holds it to its budget instead.
  build: { chunkSizeWarningLimit: 4096, ...(process.env.SITE ? { outDir: '../../dist-sites/portal', emptyOutDir: true } : {}) },
});
