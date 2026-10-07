import { builtinModules } from 'node:module';
import { defineConfig } from 'vite';

/** The Node build: the core as a library (dist-node/index.js) and the workbook-check command (dist-node/cli.js). */
export default defineConfig({
  publicDir: false,
  build: {
    outDir: 'dist-node',
    emptyOutDir: true,
    target: 'node20',
    sourcemap: true,
    minify: false,
    lib: {
      entry: { index: 'src/node/index.ts', cli: 'src/node/bin.ts' },
      formats: ['es'],
      fileName: (_format, name) => `${name}.js`,
    },
    rollupOptions: {
      external: [...builtinModules, ...builtinModules.map((m) => `node:${m}`), 'hyperformula', 'fflate', '@borgar/fx', 'numfmt'],
      output: { banner: (chunk) => (chunk.name === 'cli' ? '#!/usr/bin/env node' : '') },
    },
  },
});
