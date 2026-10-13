import { builtinModules } from 'node:module';
import { defaultServerConditions, defineConfig } from 'vite';

/**
 * The round-trip checker, a development tool run from this repository
 * (dist-verify/verify.js): it isn't part of the package.
 */
export default defineConfig({
  publicDir: false,
  // The engine from its source, as the rest of the workspace reads it, rather than its own build.
  resolve: { conditions: ['source', ...defaultServerConditions] },
  ssr: { resolve: { conditions: ['source', ...defaultServerConditions] } },
  build: {
    outDir: 'dist-verify',
    emptyOutDir: true,
    target: 'node20',
    sourcemap: true,
    minify: false,
    lib: {
      entry: { verify: 'src/verify/bin.ts' },
      formats: ['es'],
      fileName: (_format, name) => `${name}.js`,
    },
    rollupOptions: {
      external: [...builtinModules, ...builtinModules.map((m) => `node:${m}`), 'fflate', '@borgar/fx', 'numfmt'],
      output: { banner: '#!/usr/bin/env node' },
    },
  },
});
