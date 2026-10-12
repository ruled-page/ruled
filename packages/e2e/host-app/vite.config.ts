import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

// Served under /app/, so nothing in the library may assume it sits at the site root.
export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  base: '/app/',
  oxc: { jsx: { runtime: 'automatic' } },
  build: { outDir: 'dist', emptyOutDir: true },
  logLevel: 'warn',
});
