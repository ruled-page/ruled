import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { GLIDE_SCROLLER } from '../src/ui/glideDom';

const src = join(__dirname, '..', 'src');

describe("the editor's class names", () => {
  it('all start wbe-, so host styles for common names never reach the editor', () => {
    const css = readFileSync(join(src, 'editor', 'editor.css'), 'utf8');
    const preludes = css.replace(/\/\*[\s\S]*?\*\//g, '').match(/[^{};]+(?=\{)/g) ?? [];
    const classes = new Set(preludes.flatMap((p) => [...p.matchAll(/\.([a-zA-Z_][\w-]*)/g)].map((m) => m[1])));
    expect(classes.size).toBeGreaterThan(100);
    expect([...classes].filter((c) => c !== 'wbe' && !c.startsWith('wbe-') && c !== GLIDE_SCROLLER)).toEqual([]);
  });

  it('are the ones the components use', () => {
    const dir = join(src, 'ui');
    const tsx = [...readdirSync(dir).map((f) => join(dir, f)), join(src, 'editor', 'WorkbookEditor.tsx'), join(src, 'editor', 'ErrorBoundary.tsx')]
      .filter((f) => f.endsWith('.tsx'))
      .map((f) => readFileSync(f, 'utf8'))
      .join('\n');
    // Plain className="…" values, the common case.
    const used = [...tsx.matchAll(/className="([^"]*)"/g)].flatMap((m) => m[1].split(/\s+/)).filter(Boolean);
    expect(used.length).toBeGreaterThan(100);
    expect(used.filter((c) => !c.startsWith('wbe'))).toEqual([]);
  });
});
