import { parseRange, rangeToA1 } from '@ruled/formula/refs';
import { lineMap, type LineOp, type Located, type RefMap } from '@ruled/formula/reshape';
import { msg, TextError } from '../i18n/message';
import type { Rect } from './types';

/** A change to the shape of a sheet, or a move of cells, that references must follow. */
export type StructuralOp =
  | LineOp
  /** Cut and paste: the cells move, and everything that pointed at them follows. */
  | { type: 'move'; sheet: number; src: Rect; toSheet: number; row: number; col: number }
  /** An Excel table taking in the rows just below it, as typing there makes it do. Nothing moves. */
  | { type: 'growTable'; sheet: number; table: string; rows: number }
  /** Excel's Insert Cells: a block of empty cells, the cells from it down (or right) moving along to make room. */
  | { type: 'insertCells'; sheet: number; rect: Rect; shift: 'down' | 'right' }
  /** Excel's Delete Cells: the block goes, and the cells below it (or to its right) move up (or left) to close the gap. */
  | { type: 'deleteCells'; sheet: number; rect: Rect; shift: 'up' | 'left' }
  /**
   * Excel's Shift+drag: the cells move between others, which make room where
   * they land (moving down or right) and close up where they were.
   */
  | { type: 'moveInsert'; sheet: number; src: Rect; row: number; col: number; shift: 'down' | 'right' };

/** Inserting or deleting a block of cells, the cells beside it moving along. */
export type CellShiftOp = Extract<StructuralOp, { type: 'insertCells' | 'deleteCells' }>;

export const isCellShift = (op: { type: string }): op is CellShiftOp => op.type === 'insertCells' || op.type === 'deleteCells';

export const isLineOp = (op: { type: string }): op is LineOp => ['insertRows', 'deleteRows', 'insertCols', 'deleteCols'].includes(op.type);

/** Whether a cell shift moves cells up or down (rather than sideways). */
const vertical = (op: CellShiftOp) => op.shift === 'down' || op.shift === 'up';

/**
 * The whole-line change a cell shift is, within the block's own columns (or
 * rows): inserting cells in B3:C4, shifting down, is inserting rows 3 and 4
 * in columns B and C.
 */
function asLines(op: CellShiftOp): LineOp {
  const b = op.rect;
  const down = vertical(op);
  const insert = op.type === 'insertCells';
  const type = down ? (insert ? 'insertRows' : 'deleteRows') : insert ? 'insertCols' : 'deleteCols';
  return down ? { type, sheet: op.sheet, at: b.r1, count: b.r2 - b.r1 + 1 } : { type, sheet: op.sheet, at: b.c1, count: b.c2 - b.c1 + 1 };
}

/** Whether a range lies within a cell shift's columns (or rows), the only ones whose cells move. */
const inBand = (op: CellShiftOp, r: Rect) => (vertical(op) ? r.c1 >= op.rect.c1 && r.c2 <= op.rect.c2 : r.r1 >= op.rect.r1 && r.r2 <= op.rect.r2);

/**
 * Shift+drag as three simpler changes: room made where the cells land, the
 * cells moved in, and the gap they left closed.
 */
export function moveInsertSteps(op: Extract<StructuralOp, { type: 'moveInsert' }>): StructuralOp[] {
  const room: Rect = { r1: op.row, c1: op.col, r2: op.row + op.src.r2 - op.src.r1, c2: op.col + op.src.c2 - op.src.c1 };
  const insert: StructuralOp = { type: 'insertCells', sheet: op.sheet, rect: room, shift: op.shift };
  // Where the cells are once the room is made; making it mustn't tear them apart.
  const src = refMapFor(insert).range(op.sheet, op.src)?.rect;
  if (!src || splits(insert, op.src)) throw new TextError(msg('core.structure.moveInsertSplits'));
  return [insert, { type: 'move', sheet: op.sheet, src, toSheet: op.sheet, row: op.row, col: op.col }, { type: 'deleteCells', sheet: op.sheet, rect: src, shift: op.shift === 'down' ? 'up' : 'left' }];
}

/** Where Shift+drag's cells end up, once the gap they left is closed. */
export function moveInsertLanding(op: Extract<StructuralOp, { type: 'moveInsert' }>): Rect {
  const steps = moveInsertSteps(op);
  const room = (steps[0] as CellShiftOp).rect;
  return refMapFor(steps[2]).range(op.sheet, room)?.rect ?? room;
}

/**
 * Whether Shift+drag's cells land in their own columns (or rows, for a move
 * across), and a range's rows (or columns) take in both where they were and
 * where they land. Its cells are then only reordered within it, so it stays
 * as it was: a total over a list keeps every row of it when one is dragged
 * to the top or the bottom.
 */
export function reorderedWithin(op: Extract<StructuralOp, { type: 'moveInsert' }>, rect: Rect): boolean {
  if (op.shift === 'down' ? op.col !== op.src.c1 : op.row !== op.src.r1) return false;
  const [a, b] = moveInsertSpan(op);
  return op.shift === 'down' ? rect.r1 <= a && rect.r2 >= b : rect.c1 <= a && rect.c2 >= b;
}

/** The rows (or columns) Shift+drag reorders: from where its cells were to where they land, whichever comes first. */
export function moveInsertSpan(op: Extract<StructuralOp, { type: 'moveInsert' }>): [number, number] {
  const to = moveInsertLanding(op);
  return op.shift === 'down' ? [Math.min(op.src.r1, to.r1), Math.max(op.src.r2, to.r2)] : [Math.min(op.src.c1, to.c1), Math.max(op.src.c2, to.c2)];
}

export type { LineOp, Located, RefMap };

export function refMapFor(op: StructuralOp): RefMap {
  if (op.type === 'growTable') return { sheets: new Set(), range: (sheet, rect) => ({ sheet, rect }) };
  if (op.type === 'moveInsert') {
    // Room made where the cells land, the cells moved in, and the gap they left closed, as one.
    const maps = moveInsertSteps(op).map(refMapFor);
    return {
      sheets: new Set([op.sheet]),
      range: (sheet, rect) =>
        sheet === op.sheet && reorderedWithin(op, rect) ? { sheet, rect } : maps.reduce<Located | null>((at, map) => (at ? map.range(at.sheet, at.rect) : null), { sheet, rect }),
    };
  }
  if (isCellShift(op)) {
    // Only what lies within the block's columns (or rows) moves, as whole lines would; anything wider stays put.
    const lines = refMapFor(asLines(op));
    return { sheets: new Set([op.sheet]), range: (sheet, rect) => (sheet === op.sheet && inBand(op, rect) ? lines.range(sheet, rect) : { sheet, rect }) };
  }
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
  return lineMap(op);
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
  if (op.type === 'growTable') return false;
  if (op.type === 'moveInsert') {
    // Each step in turn, on the range as the steps before left it.
    let at: Rect | undefined = rect;
    for (const step of moveInsertSteps(op)) {
      if (splits(step, at)) return true;
      at = refMapFor(step).range(op.sheet, at)?.rect;
      if (!at) return false;
    }
    return false;
  }
  if (isCellShift(op)) {
    const b = op.rect;
    // The cells that move: the block's columns from its first row down (or its rows from its first column right).
    const moving = vertical(op) ? { r1: b.r1, c1: b.c1, r2: Infinity, c2: b.c2 } : { r1: b.r1, c1: b.c1, r2: b.r2, c2: Infinity };
    if (rect.r1 > moving.r2 || rect.r2 < moving.r1 || rect.c1 > moving.c2 || rect.c2 < moving.c1) return false;
    return !inBand(op, rect) || splits(asLines(op), rect);
  }
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
