import type { Plugin } from 'vite';

/**
 * Glide loads its cell editor with React.lazy, so the first time a key opens
 * it, the editor isn't there until its chunk has loaded and React shows it
 * (React 19 holds back showing suspended content for up to 300 ms). Keys
 * typed meanwhile reopen the editor and only the last one is kept: "1.25"
 * becomes "5". Importing the editor up front makes it open in the same
 * render as the key.
 */
export function glideEagerEditor(): Plugin {
  const lazy = /const DataGridOverlayEditor = React\.lazy\(async \(\) => await import\(("[^"]+data-grid-overlay-editor\.js")\)\);/;
  return {
    name: 'glide-eager-editor',
    enforce: 'pre',
    transform(code, id) {
      if (!/glide-data-grid[\\/]dist[\\/]esm[\\/]data-editor[\\/]data-editor\.js$/.test(id)) return;
      const m = lazy.exec(code);
      if (!m) this.error("Glide's editor is no longer loaded the way this plugin expects; check data-editor.js.");
      return { code: `import DataGridOverlayEditor from ${m[1]};\n${code.replace(m[0], '')}`, map: null };
    },
  };
}
