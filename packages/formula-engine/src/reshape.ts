import { parseA1Ref, stringifyA1Ref, stringifyTokens, tokenize } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS, type Rect } from './refs';

/**
 * Formulas rewritten as a workbook changes shape, as Excel rewrites them:
 * references follow cells when rows or columns are inserted or deleted,
 * and follow a sheet when it is renamed or deleted.
 */

/** Inserting or deleting `count` whole rows or columns at `at`. */
export interface LineOp {
  type: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';
  sheet: number;
  at: number;
  count: number;
}

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

/** Where cells end up when whole rows or columns are inserted or deleted. */
export function lineMap(op: LineOp): RefMap {
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

/** A sheet renamed or deleted, for the references that name it. */
export interface SheetRefChange {
  /** The sheet renamed or deleted, in lower case. */
  sheet: string;
  /** Its new name; null when it is deleted. */
  to: string | null;
  /** Every sheet's name in order, for 3-D references that lose an end. */
  order: string[];
  /** Tables that went with a deleted sheet, in lower case. */
  tables: Set<string>;
}

/** Whether a formula may refer to the sheet or its tables: a quick test before tokenizing it. */
export function mentionsSheet(text: string, c: SheetRefChange): boolean {
  const lower = text.toLowerCase();
  if (lower.includes(c.sheet) || lower.includes(c.sheet.replace(/'/g, "''"))) return true;
  for (const t of c.tables) if (lower.includes(t)) return true;
  return false;
}

/**
 * A sheet reference's context ("Data", or "Jan:Mar" for a 3-D reference)
 * after the change: undefined when it doesn't involve the sheet, null when it
 * is lost. A 3-D reference that loses an end sheet ends at the next one in.
 */
function changeContext(context: string, c: SheetRefChange): string | null | undefined {
  const ends = context.split(':');
  const hit = ends.map((e) => e.toLowerCase() === c.sheet);
  if (ends.length > 2 || !hit.some(Boolean)) return undefined;
  if (c.to !== null) return ends.map((e, i) => (hit[i] ? c.to : e)).join(':');
  if (ends.length === 1) return null;
  const [a, b] = ends.map((e) => c.order.findIndex((n) => n.toLowerCase() === e.toLowerCase()));
  const step = a < b ? 1 : -1;
  const from = hit[0] ? a + step : a;
  const to = hit[1] ? b - step : b;
  if (a < 0 || b < 0 || a === b || (to - from) * step < 0) return null;
  return from === to ? c.order[from] : `${c.order[from]}:${c.order[to]}`;
}

/** A formula (file form) after a sheet is renamed or deleted; the same string when nothing in it changes. */
export function changeSheetRefs(formula: string, c: SheetRefChange): string {
  if (!mentionsSheet(formula, c)) return formula;
  let tokens;
  try {
    tokens = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return formula;
  }
  let changed = false;
  for (const t of tokens) {
    let next: string | undefined;
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_named') {
      const ref = parseA1Ref(t.value, { allowTernary: false });
      // References into other workbooks have the workbook in their context too.
      if (!ref || ref.context?.length !== 1) continue;
      const context = changeContext(ref.context[0], c);
      if (context === undefined) continue;
      if (context === null) {
        next = '#REF!';
      } else {
        ref.context = [context];
        next = stringifyA1Ref(ref);
      }
    } else if (t.type === 'structured') {
      const table = t.value.slice(0, t.value.indexOf('['));
      if (c.tables.has(table.slice(table.lastIndexOf('!') + 1).toLowerCase())) next = '#REF!';
    }
    if (next !== undefined && next !== t.value) {
      t.value = next;
      changed = true;
    }
  }
  return changed ? stringifyTokens(tokens) : formula;
}
