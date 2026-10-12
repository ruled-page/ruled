import { parseA1Ref } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS } from '@sitepen/formula-engine/refs';
import type { ListSource, Rect } from './types';

// The parts of data validation that the page uses as well as the worker, kept
// apart from reading and writing the file so the page doesn't load that.

export interface ListContext {
  /** Lower-cased sheet name → sheet index. */
  sheetIndex: Map<string, number>;
  ownSheet: number;
  names: { name: string; formula: string; localSheet?: number }[];
}

/** Values of an inline list, e.g. "North,South" or "Say ""hi"",Bye". */
function inlineValues(text: string): string[] {
  const body = text.slice(1, text.endsWith('"') ? -1 : undefined).replace(/""/g, '"');
  return body
    .split(',')
    .map((v) => v.trim())
    .filter((v) => v !== '');
}

export function resolveListSource(formula: string, ctx: ListContext): ListSource {
  let text = formula.trim().replace(/^=/, '');
  if (text.startsWith('"')) return { kind: 'values', values: inlineValues(text) };
  // A defined name: prefer one scoped to this sheet over a workbook-wide one.
  const lower = text.toLowerCase();
  const name =
    ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === ctx.ownSheet) ??
    ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === undefined);
  if (name) text = name.formula.trim().replace(/^=/, '');
  const ref = parseA1Ref(text, { allowTernary: false });
  if (!ref || !('range' in ref) || (ref.context?.length ?? 0) > 1) return { kind: 'formula', formula: formula.trim().replace(/^=/, '') };
  const sheetName = ref.context?.[0];
  const sheet = sheetName === undefined ? ctx.ownSheet : ctx.sheetIndex.get(sheetName.toLowerCase());
  if (sheet === undefined) return { kind: 'formula', formula: formula.trim().replace(/^=/, '') };
  const r = ref.range;
  return {
    kind: 'range',
    sheet,
    rect: {
      r1: r.top ?? 0,
      c1: r.left ?? 0,
      r2: r.bottom ?? (r.top === null ? MAX_ROWS - 1 : r.top),
      c2: r.right ?? (r.left === null ? MAX_COLS - 1 : r.left),
    },
  };
}

/** The parts of a range left once `cut` is taken out of it: up to four ranges around the cut. */
export function subtractRect(r: Rect, cut: Rect): Rect[] {
  if (cut.r1 > r.r2 || cut.r2 < r.r1 || cut.c1 > r.c2 || cut.c2 < r.c1) return [r];
  const out: Rect[] = [];
  if (cut.r1 > r.r1) out.push({ ...r, r2: cut.r1 - 1 });
  if (cut.r2 < r.r2) out.push({ ...r, r1: cut.r2 + 1 });
  const r1 = Math.max(r.r1, cut.r1);
  const r2 = Math.min(r.r2, cut.r2);
  if (cut.c1 > r.c1) out.push({ r1, r2, c1: r.c1, c2: cut.c1 - 1 });
  if (cut.c2 < r.c2) out.push({ r1, r2, c1: cut.c2 + 1, c2: r.c2 });
  return out;
}

/** Ranges with others taken out of them. */
export function withoutRects(rects: readonly Rect[], cuts: readonly Rect[]): Rect[] {
  return cuts.reduce<Rect[]>((left, cut) => left.flatMap((r) => subtractRect(r, cut)), [...rects]);
}

