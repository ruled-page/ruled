import { defineConfig } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';

/**
 * The published package: a module per connector under dist/, with its
 * types. MSAL stays a package of its own, imported only when someone first
 * reaches for OneDrive, and @ruled/sheet is only ever a source of types.
 * Inside the workspace the package is read from its source instead.
 */
export default defineConfig({
  input: { 'onedrive/index': 'src/onedrive/index.ts', 'onedrive/redirect': 'src/onedrive/redirect.ts' },
  external: [/^@azure\/msal-browser(\/|$)/, /^@ruled\/sheet(\/|$)/],
  plugins: [dts({ tsconfig: 'tsconfig.lib.json' })],
  output: { dir: 'dist', format: 'esm', sourcemap: true, cleanDir: true },
});
