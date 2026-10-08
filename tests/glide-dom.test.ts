import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { GLIDE_CLICK_OUTSIDE_IGNORE, GLIDE_CLIP_REGION, GLIDE_SCROLLER, PORTAL_ID } from '../src/ui/glideDom';
import { glidePatches } from '../vite.glide.ts';

const esm = join(__dirname, '..', 'node_modules', '@glideapps', 'glide-data-grid', 'dist', 'esm');

function modules(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? modules(join(dir, e.name)) : e.name.endsWith('.js') ? [join(dir, e.name)] : []));
}

describe('the installed Glide', () => {
  it('still has the parts of its DOM the editor relies on', () => {
    const code = modules(esm).map((f) => readFileSync(f, 'utf8')).join('\n');
    for (const name of [GLIDE_SCROLLER, GLIDE_CLIP_REGION]) expect(code, name).toMatch(new RegExp(`className: "${name}[ "]`));
    expect(code).toContain(`classList.contains("${GLIDE_CLICK_OUTSIDE_IGNORE}")`);
  });

  it('still takes each of the editor’s patches', () => {
    const plugin = glidePatches();
    const transform = plugin.transform as (this: unknown, code: string, id: string) => { code: string } | undefined;
    const context = { error: (m: string) => { throw new Error(m); } };
    const patch = (file: string) => transform.call(context, readFileSync(join(esm, file), 'utf8'), join(esm, file))!.code;

    const editor = patch('data-editor/data-editor.js');
    expect(editor).toMatch(/^import DataGridOverlayEditor from "[^"]+data-grid-overlay-editor\.js";/);
    expect(editor).not.toContain('React.lazy(async () => await import("../internal/data-grid-overlay-editor');
    expect(patch('internal/data-grid-overlay-editor/data-grid-overlay-editor.js')).toContain(`document.getElementById("${PORTAL_ID}")`);
    const grid = patch('internal/data-grid/data-grid.js');
    for (const part of ['"aria-label": sheetAccess?.label', 'sheetAccess ? null : React.createElement("thead"', 'sheetAccess?.rowIndex(row) ?? row + 2', 'sheetAccess?.colIndex(col) ?? col + 1 + colOffset, "aria-label": sheetAccess?.cellLabel(cellContent)', 'experimental?.sheetAccess, ], 200)']) {
      expect(grid, part).toContain(part);
    }
    expect(transform.call(context, 'export {}', join(esm, 'index.js'))).toBeUndefined();
    expect(() => transform.call(context, 'export {}', join(esm, 'data-editor', 'data-editor.js'))).toThrow(/no longer applies/);
  });
});
