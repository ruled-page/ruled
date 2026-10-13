import { builtinModules } from 'node:module';
import { defaultServerConditions, defineConfig } from 'vite';

/**
 * The Node build, dist/node/index.js: `recalculate`, with the workbook
 * reader, writer and engine bundled in, so the package still has no
 * dependencies of its own. Only Node's own modules are left to it.
 */
export default defineConfig({
  publicDir: false,
  // The engine from its source, as the rest of the workspace reads it, rather than its own build.
  resolve: { conditions: ['source', ...defaultServerConditions] },
  ssr: { resolve: { conditions: ['source', ...defaultServerConditions] } },
  build: {
    outDir: 'dist/node',
    emptyOutDir: true,
    target: 'node20',
    sourcemap: true,
    lib: {
      entry: { index: 'src/node/index.ts' },
      formats: ['es'],
      fileName: (_format, name) => `${name}.js`,
    },
    rollupOptions: {
      external: [...builtinModules, ...builtinModules.map((m) => `node:${m}`)],
    },
  },
});
