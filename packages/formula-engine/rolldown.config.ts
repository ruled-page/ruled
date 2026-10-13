import { defineConfig } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';

/**
 * The published package: dist/index.js, the engine and `createWorkbook` as
 * one module, and dist/index.d.ts, its types. Its two dependencies, the
 * formula parser and the number formatter, stay packages of their own.
 * Inside the workspace the package is read from its source instead.
 */
export default defineConfig({
  input: { index: 'src/index.ts' },
  external: [/^@borgar\/fx(\/|$)/, /^numfmt(\/|$)/],
  plugins: [dts({ tsconfig: 'tsconfig.lib.json' })],
  output: { dir: 'dist', format: 'esm', sourcemap: true, cleanDir: true },
});
