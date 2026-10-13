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
  /*
   * Glide's overlay schedules a finish for every key pressed in a cell
   * editor, and each reads where to move once it runs, not when its key was
   * pressed. A Backspace's, still waiting when Enter comes a moment later,
   * finishes without saving, and Enter's then finds editing over: the edit is
   * lost. Only the keys that finish editing (Enter, Tab, Escape) schedule one.
   */
  {
    file: /glide-data-grid[\\/]dist[\\/]esm[\\/]internal[\\/]data-grid-overlay-editor[\\/]data-grid-overlay-editor\.js$/,
    find: /(\n\s*)window\.setTimeout\(\(\) => \{\s*if \(!finished\.current && customMotion\.current !== undefined\)/,
    replace: (m) => `${m[1]}if (!save && event.key !== "Escape") return;${m[0]}`,
  },
  /*
   * Glide's table for screen readers numbers rows and columns by their place
   * in the grid, with the column letters as row 1, so B5 is read as row 6,
   * column 3, and rows and columns after hidden ones, or below frozen ones,
   * are numbered wrongly. Given `experimental.sheetAccess`, the table is
   * named, numbered as the sheet is, and leaves out the row numbers and
   * column letters; each cell is named with its reference instead.
   */
  ...accessTreePatches(),
];

/** The patches for `experimental.sheetAccess`; see SheetAccess in src/ui/SheetGrid.tsx. */
function accessTreePatches(): Patch[] {
  const file = /glide-data-grid[\\/]dist[\\/]esm[\\/]internal[\\/]data-grid[\\/]data-grid\.js$/;
  const swap = (find: RegExp, to: string): Patch => ({ file, find, replace: () => to });
  return [
    swap(
      /const colOffset = firstColAccessible \? 0 : -1;\s*if \(!firstColAccessible && /,
      'const sheetAccess = experimental?.sheetAccess; const colOffset = firstColAccessible && !sheetAccess ? 0 : -1; if (colOffset === -1 && ',
    ),
    swap(
      /React\.createElement\("table", \{ key: "access-tree", role: "grid", "aria-rowcount": rows \+ 1, "aria-multiselectable": "true", "aria-colcount": mappedColumns\.length \+ colOffset \},\s*React\.createElement\("thead"/,
      'React.createElement("table", { key: "access-tree", role: "grid", "aria-label": sheetAccess?.label, "aria-rowcount": sheetAccess?.rowCount ?? rows + 1, "aria-multiselectable": "true", "aria-colcount": sheetAccess?.colCount ?? mappedColumns.length + colOffset }, sheetAccess ? null : React.createElement("thead"',
    ),
    swap(/"aria-rowindex": row \+ 2 \}/, '"aria-rowindex": sheetAccess?.rowIndex(row) ?? row + 2 }'),
    swap(
      /role: "gridcell", "aria-colindex": col \+ 1 \+ colOffset,/,
      'role: "gridcell", "aria-colindex": sheetAccess?.colIndex(col) ?? col + 1 + colOffset, "aria-label": sheetAccess?.cellLabel(cellContent),',
    ),
    // The table is built again when what it numbers by changes.
    swap(/getBoundsForItem,\s*onCellFocused,\s*\], 200\);/, 'getBoundsForItem, onCellFocused, experimental?.sheetAccess, ], 200);'),
  ];
}

/** Changes to Glide that the editor needs, applied as it is bundled. */
export function glidePatches(): Plugin {
  return {
    name: 'glide-patches',
    enforce: 'pre',
    transform(code, id) {
      const patches = PATCHES.filter((p) => p.file.test(id));
      if (!patches.length) return;
      let out = code;
      for (const patch of patches) {
        const m = patch.find.exec(out);
        if (!m) this.error(`Glide's ${id.split(/[\\/]/).pop()} has changed, so a patch in vite.glide.ts no longer applies.`);
        out = (patch.prepend?.(m) ?? '') + out.replace(m[0], patch.replace(m));
      }
      return { code: out, map: null };
    },
  };
}
