import { fileURLToPath } from 'node:url';
import { defaultClientConditions, defineConfig, type Plugin } from 'vite';
import { glidePatches } from './vite.glide.ts';

/**
 * Vite writes the React build's worker URL as an expression that other
 * bundlers are told to ignore, so a host app's bundler would never copy the
 * worker. This puts back the plain form that Vite, webpack and Parcel all
 * recognise and bundle as a worker:
 * new Worker(new URL('./assets/workbook.worker.js', import.meta.url)).
 */
function plainWorkerUrl(): Plugin {
  return {
    name: 'plain-worker-url',
    apply: 'build',
    enforce: 'post',
    generateBundle(_options, bundle) {
      const vite = /new URL\(\s*\/\* @vite-ignore \*\/\s*new URL\("assets\/workbook\.worker\.js", import\.meta\.url\)\.href,\s*"" \+ import\.meta\.url\s*\)/;
      for (const chunk of Object.values(bundle)) {
        if (chunk.type !== 'chunk' || !chunk.code.includes('workbook.worker.js')) continue;
        const m = vite.exec(chunk.code);
        if (!m) this.error("The worker's URL isn't in the form this build expects; check the output of new Worker(...).");
        // As many line breaks as before, so the source map still lines up.
        chunk.code = chunk.code.replace(m[0], `new URL("./assets/workbook.worker.js", import.meta.url)${'\n'.repeat(m[0].split('\n').length - 1)}`);
      }
    },
  };
}

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
    // Paths relative to the library, wherever the host serves it from.
    base: './',
    // A fixed name, which the package exports as @ruled/sheet/worker.
    // One file, the engine inlined, so hosts' bundlers see a single worker.
    worker: { format: 'es', rolldownOptions: { output: { entryFileNames: 'assets/workbook.worker.js', codeSplitting: false } } },
    plugins: standalone ? [glidePatches()] : [glidePatches(), plainWorkerUrl()],
    resolve: {
      // The engine from its source, as the rest of the workspace reads it, rather than its own build.
      conditions: ['source', ...defaultClientConditions],
      // CommonJS modules that require('react'), which an ES module library can't do once React is external.
      alias: [
        { find: /^use-sync-external-store\/shim$/, replacement: fileURLToPath(new URL('src/editor/shims/useSyncExternalStore.ts', import.meta.url)) },
        { find: /^react-responsive-carousel$/, replacement: fileURLToPath(new URL('src/editor/shims/carousel.ts', import.meta.url)) },
        // Not CommonJS, but the markdown it renders is never shown.
        { find: /^marked$/, replacement: fileURLToPath(new URL('src/editor/shims/marked.ts', import.meta.url)) },
      ],
    },
    define: { 'process.env.NODE_ENV': JSON.stringify('production') },
    build: {
      outDir: `dist/${standalone ? 'standalone' : 'react'}`,
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
