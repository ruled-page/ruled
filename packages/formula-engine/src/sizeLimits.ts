import { parseA1Ref, type Token } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS } from './refs';

/**
 * Limits on what a calculation engine is asked to do, so that a workbook
 * can't hang the editor or run it out of memory with a few short formulas.
 * Formulas over the limits keep Excel's saved value, like any other formula
 * the engine can't calculate.
 */

/** Most cells one reference may cover: ten whole columns. */
export const MAX_RANGE_CELLS = 10 * MAX_ROWS;
/** Most values a function such as SEQUENCE may produce. */
export const MAX_ARRAY_CELLS = 1_000_000;

/** Functions whose literal arguments give the size of the array they return. */
const ARRAY_SIZE: Record<string, (args: number[]) => number> = {
  SEQUENCE: ([rows = 1, cols = 1]) => rows * cols,
  MUNIT: ([n = 1]) => n * n,
  RANDARRAY: ([rows = 1, cols = 1]) => rows * cols,
};

/** Whether a tokenized formula reads a range or builds an array too large to calculate here. */
export function tooLargeToCalculate(tokens: Token[]): boolean {
  const solid = tokens.filter((t) => t.type !== 'whitespace' && t.type !== 'newline');
  for (let i = 0; i < solid.length; i++) {
    const t = solid[i];
    if (t.type === 'range' || t.type === 'range_beam') {
      const r = parseA1Ref(t.value, { allowTernary: false });
      const range = r && 'range' in r ? r.range : undefined;
      if (!range) continue;
      const rows = (range.bottom ?? MAX_ROWS - 1) - (range.top ?? 0) + 1;
      const cols = (range.right ?? MAX_COLS - 1) - (range.left ?? 0) + 1;
      if (rows * cols > MAX_RANGE_CELLS) return true;
    }
    if (t.type === 'func') {
      const size = ARRAY_SIZE[t.value.toUpperCase().replace(/^_XLFN\./, '')];
      if (!size || solid[i + 1]?.value !== '(') continue;
      const args: number[] = [];
      for (let k = i + 2; k < solid.length && solid[k].type === 'number'; k += 2) {
        args.push(Math.trunc(Number(solid[k].value)));
        if (solid[k + 1]?.value !== ',') break;
      }
      if (size(args) > MAX_ARRAY_CELLS) return true;
    }
  }
  return false;
}

