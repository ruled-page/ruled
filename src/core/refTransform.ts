import { parseA1Ref, stringifyA1Ref, stringifyTokens, tokenize } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS, parseRange, rangeToA1 } from './refs';
import type { Rect } from './types';

/** A change to the shape of a sheet, or a move of cells, that references must follow. */
export type StructuralOp =
  | { type: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; sheet: number; at: number; count: number }
  /** Cut and paste: the cells move, and everything that pointed at them follows. */
  | { type: 'move'; sheet: number; src: Rect; toSheet: number; row: number; col: number };

/** Inserting or deleting whole rows or columns. */
export type LineOp = Exclude<StructuralOp, { type: 'move' }>;

export interface Located {
  sheet: number;
  rect: Rect;
}

/** Where cells end up after a structural change. */
export interface RefMap {
  /** Sheets whose cells move; references to any other sheet are unchanged. */
  sheets: Set<number>;
  /** Where a range went; null when it no longer exists (a reference to it becomes #REF!). */
  range(sheet: number, rect: Rect): Located | null;
}

/** Positions along one axis after inserting or deleting `count` lines at `at`. */
function lineShift(insert: boolean, at: number, count: number, max: number) {
  return (a: number, b: number): [number, number] | null => {
    if (insert) {
      if (a >= at) return a + count > max ? null : [a + count, Math.min(max, b + count)];
      // A range that straddles the insertion point grows to take in the new lines.
      if (b >= at) return [a, Math.min(max, b + count)];
      return [a, b];
    }
    const end = at + count;
    const na = a < at ? a : a >= end ? a - count : at;
    const nb = b < at ? b : b >= end ? b - count : at - 1;
    return nb < na ? null : [na, nb];
  };
}

export function refMapFor(op: StructuralOp): RefMap {
  if (op.type === 'move') {
    const { src } = op;
    const dr = op.row - src.r1;
    const dc = op.col - src.c1;
    const dst: Rect = { r1: op.row, c1: op.col, r2: src.r2 + dr, c2: src.c2 + dc };
    const within = (r: Rect, box: Rect) => r.r1 >= box.r1 && r.r2 <= box.r2 && r.c1 >= box.c1 && r.c2 <= box.c2;
    return {
      sheets: new Set([op.sheet, op.toSheet]),
      range(sheet, rect) {
        if (sheet === op.sheet && within(rect, src)) {
          return { sheet: op.toSheet, rect: { r1: rect.r1 + dr, r2: rect.r2 + dr, c1: rect.c1 + dc, c2: rect.c2 + dc } };
        }
        // Cells pasted over are gone, so references that only covered them break, as in Excel.
        if (sheet === op.toSheet && within(rect, dst)) return null;
        return { sheet, rect };
      },
    };
  }
  const rows = op.type === 'insertRows' || op.type === 'deleteRows';
  const shift = lineShift(op.type.startsWith('insert'), op.at, op.count, (rows ? MAX_ROWS : MAX_COLS) - 1);
  return {
    sheets: new Set([op.sheet]),
    range(sheet, rect) {
      if (sheet !== op.sheet) return { sheet, rect };
      const moved = rows ? shift(rect.r1, rect.r2) : shift(rect.c1, rect.c2);
      if (!moved) return null;
      return { sheet, rect: rows ? { ...rect, r1: moved[0], r2: moved[1] } : { ...rect, c1: moved[0], c2: moved[1] } };
    },
  };
}

export interface FormulaContext {
  /** Sheet names by index, for writing qualified references. */
  names: string[];
  /** Sheet index for a name, case-insensitively; undefined if not in this workbook. */
  indexOf(name: string): number | undefined;
}

/**
 * Rewrites the references in a formula (file form, no "=") for a structural
 * change. Unqualified references are read as on `fromSheet` and written for
 * `toSheet`, which differ when a formula is moved to another sheet. Whole
 * rows and columns keep their open ends; references to deleted cells become
 * #REF!. Returns the formula unchanged when nothing in it moves.
 */
export function transformFormula(formula: string, fromSheet: number, toSheet: number, map: RefMap, ctx: FormulaContext): string {
  let tokens;
  try {
    tokens = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return formula;
  }
  let changed = false;
  for (const t of tokens) {
    if (t.type !== 'range' && t.type !== 'range_beam') continue;
    const ref = parseA1Ref(t.value, { allowTernary: false });
    if (!ref || !('range' in ref)) continue;
    const context = ref.context ?? [];
    // External workbooks and 3-D references are left as they are.
    if (context.length > 1 || context[0]?.includes(':')) continue;
    const sheet = context.length ? ctx.indexOf(context[0]) : fromSheet;
    if (sheet === undefined || sheet < 0) continue;
    if (!map.sheets.has(sheet) && (context.length || fromSheet === toSheet)) continue;
    const r = ref.range;
    const rect: Rect = {
      r1: r.top ?? 0,
      r2: r.bottom ?? (r.top === null ? MAX_ROWS - 1 : r.top),
      c1: r.left ?? 0,
      c2: r.right ?? (r.left === null ? MAX_COLS - 1 : r.left),
    };
    const to = map.range(sheet, rect);
    let text: string;
    if (!to) {
      text = '#REF!';
    } else {
      if (r.top !== null) {
        r.top = to.rect.r1;
        r.bottom = r.bottom === null ? null : to.rect.r2;
      }
      if (r.left !== null) {
        r.left = to.rect.c1;
        r.right = r.right === null ? null : to.rect.c2;
      }
      // Qualify when the reference now points at a sheet other than the formula's own.
      ref.context = to.sheet !== toSheet ? [ctx.names[to.sheet]] : context.length && to.sheet === sheet ? context : [];
      text = stringifyA1Ref(ref);
    }
    if (text !== t.value) {
      t.value = text;
      changed = true;
    }
  }
  return changed ? stringifyTokens(tokens) : formula;
}

/**
 * Rewrites a space-separated list of ranges on `sheet` (a sqref). Ranges that
 * are deleted, or moved to another sheet, are dropped; returns '' if none remain.
 */
export function transformSqref(sqref: string, sheet: number, map: RefMap): string {
  if (!map.sheets.has(sheet)) return sqref;
  const out: string[] = [];
  for (const part of sqref.trim().split(/\s+/)) {
    const rect = parseRange(part);
    if (!rect) {
      out.push(part);
      continue;
    }
    const to = map.range(sheet, rect);
    if (to && to.sheet === sheet) out.push(rangeToA1(to.rect));
  }
  return out.join(' ');
}

/** Whether `rect` is split by the change: some of it moves or goes and some doesn't, or it is pushed apart. */
export function splits(op: StructuralOp, rect: Rect): boolean {
  if (op.type === 'move') {
    const s = op.src;
    const overlaps = rect.r1 <= s.r2 && rect.r2 >= s.r1 && rect.c1 <= s.c2 && rect.c2 >= s.c1;
    const inside = rect.r1 >= s.r1 && rect.r2 <= s.r2 && rect.c1 >= s.c1 && rect.c2 <= s.c2;
    return overlaps && !inside;
  }
  const rows = op.type === 'insertRows' || op.type === 'deleteRows';
  const [a, b] = rows ? [rect.r1, rect.r2] : [rect.c1, rect.c2];
  if (op.type.startsWith('insert')) return a < op.at && op.at <= b;
  const end = op.at + op.count - 1;
  const overlaps = a <= end && b >= op.at;
  return overlaps && !(a >= op.at && b <= end);
}
