import { fileURLToPath } from 'node:url';
import { defaultClientConditions, defaultServerConditions } from 'vite';
import { defineConfig } from 'vitest/config';
import { glidePatches } from './packages/spreadsheet/vite.glide.ts';

/**
 * The unit tests of every package and example, run together so coverage
 * spans them all; `npm run engine:bench` and the other benchmarks run the
 * packages' bench/ instead. Workspace packages are read from their source
 * (their exports' "source" condition) rather than their builds.
 */
const source = { conditions: ['source', ...defaultClientConditions] };
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  plugins: [glidePatches()],
  resolve: {
    ...source,
    // The grid's markdown cells are never shown, so their parser isn't needed.
    alias: [{ find: /^marked$/, replacement: fileURLToPath(new URL('packages/spreadsheet/src/editor/shims/marked.ts', import.meta.url)) }],
  },
  ssr: { resolve: { conditions: ['source', ...defaultServerConditions] } },
  test: {
    environment: 'node',
    include: process.env.ENGINE_BENCH ? ['packages/*/bench/**/*.bench.ts'] : ['packages/*/tests/**/*.test.ts', 'examples/*/tests/**/*.test.ts'],
    setupFiles: ['packages/spreadsheet/tests/setup.ts'],
    testTimeout: 60_000,
    coverage: {
      provider: 'v8',
      include: ['packages/*/src/**/*.{ts,tsx}', 'examples/{demo,react-portal}/src/**/*.{ts,tsx}'],
      // The React screens and the apps' entry points are tested in a browser by the Playwright suite (packages/e2e).
      exclude: [
        'packages/spreadsheet/src/ui/**/*.tsx',
        'packages/spreadsheet/src/ui/useDialog.ts',
        'packages/spreadsheet/src/editor/hooks.ts',
        'packages/spreadsheet/src/editor/{WorkbookEditor.tsx,ErrorBoundary.tsx,mount.tsx,element.ts,standalone.ts,index.ts}',
        'examples/demo/src/{main.tsx,Welcome.tsx,i18n.ts,sites.ts}',
        'examples/react-portal/src/{*.tsx,i18n.ts,format.ts,sites.ts}',
        'examples/*/src/locales/**',
        'packages/spreadsheet/src/i18n/messages/**',
      ],
      reporter: ['text-summary', 'html'],
      // Just under what the suite reaches, so coverage can't quietly slip.
      thresholds: { lines: 98, statements: 97, functions: 97, branches: 90 },
    },
  },
});
