import { defineConfig } from 'vite';

/**
 * Library builds of the editor:
 *   --mode react       React component for React hosts (react and react-dom are peers)
 *   --mode standalone  everything bundled, plus the <workbook-editor> element
 */
export default defineConfig(({ mode }) => {
  const standalone = mode === 'standalone';
  return {
    oxc: { jsx: { runtime: 'automatic' } },
    publicDir: false,
    worker: { format: 'es' },
    define: { 'process.env.NODE_ENV': JSON.stringify('production') },
    build: {
      outDir: `dist-lib/${standalone ? 'standalone' : 'react'}`,
      emptyOutDir: true,
      sourcemap: true,
      lib: {
        entry: standalone ? 'src/editor/standalone.ts' : 'src/editor/index.ts',
        formats: ['es'],
        fileName: () => 'workbook-editor.js',
        cssFileName: 'workbook-editor',
      },
      rollupOptions: {
        external: standalone ? [] : ['react', 'react-dom', 'react-dom/client', 'react/jsx-runtime'],
      },
    },
  };
});
