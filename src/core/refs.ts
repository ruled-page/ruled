import type { Rect } from './types';

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

/** Stable numeric key for a cell, unique within a sheet. */
export function cellKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

export function keyRow(key: number): number {
  return Math.floor(key / MAX_COLS);
}

export function keyCol(key: number): number {
  return key % MAX_COLS;
}

/** 0 → "A", 25 → "Z", 26 → "AA". */
export function colToLetters(col: number): string {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" → 0, "AA" → 26. Case-insensitive. */
export function lettersToCol(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + ((letters.charCodeAt(i) & ~32) - 64);
  }
  return n - 1;
}

export function toA1(row: number, col: number): string {
  return colToLetters(col) + (row + 1);
}

/** Parses "B12" or "$B$12" into 0-based coordinates. */
export function parseA1(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  if (!m) return null;
  return { row: parseInt(m[2], 10) - 1, col: lettersToCol(m[1]) };
}

/** Parses "A1:C3" (or a single cell) into a rectangle. */
export function parseRange(ref: string): Rect | null {
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const [a, b] = parts;
  const p1 = parseA1(a);
  const p2 = b ? parseA1(b) : p1;
  if (!p1 || !p2) return null;
  return {
    r1: Math.min(p1.row, p2.row),
    c1: Math.min(p1.col, p2.col),
    r2: Math.max(p1.row, p2.row),
    c2: Math.max(p1.col, p2.col),
  };
}

export function rangeToA1(r: Rect): string {
  const a = toA1(r.r1, r.c1);
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${toA1(r.r2, r.c2)}`;
}

export function inRect(r: Rect, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

/** Quotes a sheet name for use in a formula when Excel would. */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseA1(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}
