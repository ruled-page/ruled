import { describe, expect, it } from 'vitest';
import { CellFlag, VT } from '../src/core/types';
import { autoComplete, columnEntries } from '../src/state/autoComplete';
import type { CellRecord } from '../src/state/format';

describe("AutoComplete's entries", () => {
  /** Column A, from row 0: text, a number, a formula's text, an empty cell, then more text. */
  const cells: (CellRecord | null)[] = [
    { t: VT.String, s: 'North', style: 0, flags: 0 },
    { t: VT.Number, n: 4, style: 0, flags: 0 },
    { t: VT.String, s: 'Sums', style: 0, flags: CellFlag.Formula },
    { t: VT.String, s: 'Spilled', style: 0, flags: CellFlag.ArrayMember },
    { t: VT.String, s: 'South', style: 0, flags: 0 },
    null,
    { t: VT.String, s: 'Far away', style: 0, flags: 0 },
  ];
  const get = (row: number, col: number) => (col === 0 ? (cells[row] ?? null) : null);

  it("are the text above and below the cell in its column, up to an empty cell, not numbers or formulas' results", () => {
    expect(columnEntries(get, 2, 0)).toEqual(['North', 'South']);
    expect(columnEntries(get, 4, 0)).toEqual(['North']);
    expect(columnEntries(get, 6, 0)).toEqual([]);
  });

  it('stop at a cell that has only a format', () => {
    const blank = (row: number) => (row === 1 ? { t: VT.None, style: 3, flags: 0 } : { t: VT.String, s: `Row ${row}`, style: 0, flags: 0 });
    expect(columnEntries((r) => (r < 4 ? blank(r) : null), 2, 0)).toEqual(['Row 3']);
  });
});

describe('AutoComplete', () => {
  const entries = ['Apple', 'Apricot', 'Banana', 'banana', 'Yes', 'Yesterday'];

  it('finishes text that starts just one entry, ignoring case', () => {
    expect(autoComplete('b', entries)).toBe('Banana');
    expect(autoComplete('APP', entries)).toBe('Apple');
    expect(autoComplete('apple', entries)).toBe('Apple');
  });

  it('waits while the text starts more than one entry, or none', () => {
    expect(autoComplete('ap', entries)).toBeUndefined();
    expect(autoComplete('yes', entries)).toBeUndefined();
    expect(autoComplete('yest', entries)).toBe('Yesterday');
    expect(autoComplete('cherry', entries)).toBeUndefined();
  });

  it('leaves formulas, text on several lines and nothing at all alone', () => {
    expect(autoComplete('=b', ['=bob'])).toBeUndefined();
    expect(autoComplete('b\nb', ['b\nbanana'])).toBeUndefined();
    expect(autoComplete('', entries)).toBeUndefined();
  });

  it("doesn't finish text whose lower case is longer than the entry's", () => {
    // "İ" lowers to two characters, so "i̇" typed is the start of it, but the two don't line up.
    expect(autoComplete('i̇', ['İ'])).toBeUndefined();
  });
});
