import type { Plugin } from 'vite';

/** A change to one of Glide's modules: the file, and what to replace in it. */
interface Patch {
  file: RegExp;
  find: RegExp;
  replace: (m: RegExpExecArray) => string;
  /** Code to put before the module's own. */
  prepend?: (m: RegExpExecArray) => string;
}

const PATCHES: Patch[] = [
  /*
   * Glide loads its cell editor with React.lazy, so the first time a key
   * opens it, the editor isn't there until its chunk has loaded and React
   * shows it (React 19 holds back showing suspended content for up to
   * 300 ms). Keys typed meanwhile reopen the editor and only the last one is
   * kept: "1.25" becomes "5". Importing the editor up front makes it open in
   * the same render as the key.
   */
  {
    file: /glide-data-grid[\\/]dist[\\/]esm[\\/]data-editor[\\/]data-editor\.js$/,
    find: /const DataGridOverlayEditor = React\.lazy\(async \(\) => await import\(("[^"]+data-grid-overlay-editor\.js")\)\);/,
    replace: () => '',
    prepend: (m) => `import DataGridOverlayEditor from ${m[1]};\n`,
  },
  /*
   * Glide opens cell editors in the page's element with id "portal", which a
   * host page may well have for its own use. The editor makes its own
   * (src/ui/glideDom.ts, PORTAL_ID).
   */
  {
    file: /glide-data-grid[\\/]dist[\\/]esm[\\/]internal[\\/]data-grid-overlay-editor[\\/]data-grid-overlay-editor\.js$/,
    find: /document\.getElementById\("portal"\)/,
    replace: () => 'document.getElementById("wbe-portal")',
  },
];

/** Changes to Glide that the editor needs, applied as it is bundled. */
export function glidePatches(): Plugin {
  return {
    name: 'glide-patches',
    enforce: 'pre',
    transform(code, id) {
      const patch = PATCHES.find((p) => p.file.test(id));
      if (!patch) return;
      const m = patch.find.exec(code);
      if (!m) this.error(`Glide's ${id.split(/[\\/]/).pop()} has changed, so a patch in vite.glide.ts no longer applies.`);
      return { code: (patch.prepend?.(m) ?? '') + code.replace(m[0], patch.replace(m)), map: null };
    },
  };
}
