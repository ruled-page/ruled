import { defineConfig } from 'vitest/config';

// Glide's build-time dependencies conflict with @vitejs/plugin-react's optional
// Babel 8 peer, so JSX is compiled by Vite's built-in Oxc transform instead.
export default defineConfig({
  oxc: { jsx: { runtime: 'automatic' } },
  worker: { format: 'es' },
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
  },
});
