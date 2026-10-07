import { describe, expect, it } from 'vitest';
import { MAX_ROWS } from '../src/core/refs';
import { VT, type ListSource, type ListValidation } from '../src/core/types';
import { english } from '../src/i18n';
import { DEFAULT_LIST_ERROR, listAllows, listOptions, listValidationAt } from '../src/state/validation';
import { sheet, sheetInfo, style, type FileCell } from './state-fixtures';

const rule = (source: ListSource, rects = [{ r1: 0, c1: 0, r2: 9, c2: 0 }]): ListValidation => ({
  rects,
  source,
  inCellDropdown: true,
  allowBlank: true,
  showError: true,
  errorStyle: 'stop',
});

const styles = [style(), style({ numFmt: '0.00' }), style({ numFmt: 'd mmm yyyy' })];
const range = (sheet: number, r1: number, c1: number, r2: number, c2: number): ListSource => ({ kind: 'range', sheet, rect: { r1, c1, r2, c2 } });

describe('finding the list rule for a cell', () => {
  it('picks the rule whose ranges cover the cell', () => {
    const first = rule({ kind: 'values', values: ['a'] }, [{ r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 5, c1: 2, r2: 6, c2: 2 }]);
    const second = rule({ kind: 'values', values: ['b'] }, [{ r1: 0, c1: 1, r2: 9, c2: 1 }]);
    const info = sheetInfo({ listValidations: [first, second] });
    expect(listValidationAt(info, 6, 2)).toBe(first);
    expect(listValidationAt(info, 3, 1)).toBe(second);
    expect(listValidationAt(info, 3, 2)).toBeUndefined();
  });
});

describe('the options a list offers', () => {
  const lists = sheet(
    [
      [0, 0, 'North'],
      [1, 0, '  South  '],
      [2, 0, 'North'],
      [4, 0, 'East'],
      [5, 0, 1.5, 1],
      [6, 0, 46023, 2],
      [7, 0, true],
      [8, 0, '   '],
      [9, 0, 'Odd style', 42],
      [0, 1, 'Second column'],
    ],
    { name: 'Lists', index: 1 },
  );
  const views = [sheet([[0, 0, 'x']]), lists];

  it('gives typed-in values as they are', () => {
    expect(listOptions(rule({ kind: 'values', values: ['Yes', 'No'] }), views, styles, false)).toEqual(['Yes', 'No']);
  });

  it("shows a range's cells as they are displayed, without blanks or repeats", () => {
    expect(listOptions(rule(range(1, 0, 0, 9, 0)), views, styles, false)).toEqual(['North', 'South', 'East', '1.50', '1 Jan 2026', 'TRUE', 'Odd style']);
  });

  it('reads across the columns of a range, row by row', () => {
    expect(listOptions(rule(range(1, 0, 0, 1, 1)), views, styles, false)).toEqual(['North', 'Second column', 'South']);
  });

  it('shows dates in the 1904 date system', () => {
    expect(listOptions(rule(range(1, 6, 0, 6, 0)), views, styles, true)).toEqual(['2 Jan 2030']);
  });

  it('has nothing to offer from a formula, or from a sheet that is not there', () => {
    expect(listOptions(rule({ kind: 'formula', formula: 'OFFSET(A1,0,0,5)' }), views, styles, false)).toBeUndefined();
    expect(listOptions(rule(range(7, 0, 0, 9, 0)), views, styles, false)).toBeUndefined();
  });

  it('follows edits to the source cells', () => {
    const live = sheet([[0, 0, 'Old']]);
    live.apply({ sheet: 0, row: 0, col: 0, t: VT.String, s: 'New', style: 0, flags: 0 });
    live.apply({ sheet: 0, row: 150, col: 0, t: VT.String, s: 'Typed further down', style: 0, flags: 0 });
    expect(listOptions(rule(range(0, 0, 0, MAX_ROWS - 1, 0)), [live], styles, false)).toEqual(['New', 'Typed further down']);
  });

  it("reads a whole-column source only a little past the sheet's used range", () => {
    const view = sheet([[0, 0, 'Top']]);
    view.apply({ sheet: 0, row: 5000, col: 0, t: VT.String, s: 'Far below', style: 0, flags: 0 });
    expect(listOptions(rule(range(0, 0, 0, MAX_ROWS - 1, 0)), [view], styles, false)).toEqual(['Top']);
  });

  it('stops at a thousand options to keep the dropdown usable', () => {
    const cells: FileCell[] = Array.from({ length: 1500 }, (_, i) => [i, 0, `Item ${i}`]);
    const options = listOptions(rule(range(0, 0, 0, 1499, 0)), [sheet(cells)], styles, false);
    expect(options).toHaveLength(1000);
    expect(options?.at(-1)).toBe('Item 999');
  });
});

describe('checking typed text against a list', () => {
  it('ignores case and surrounding spaces', () => {
    expect(listAllows(['North', 'South'], '  SOUTH ')).toBe(true);
    expect(listAllows([' North '], 'north')).toBe(true);
    expect(listAllows(['North'], 'Northern')).toBe(false);
    expect(listAllows([], 'North')).toBe(false);
  });

  it('compares numbers by value, with or without grouping', () => {
    expect(listAllows(['1,000', '2'], '1000')).toBe(true);
    expect(listAllows(['1000'], '1,000.00')).toBe(true);
    expect(listAllows(['0.5'], '.5')).toBe(true);
    expect(listAllows(['1500'], '1.5e3')).toBe(true);
    expect(listAllows(['-3'], '-3.0')).toBe(true);
    expect(listAllows(['10'], '1')).toBe(false);
  });

  it('does not read blanks as zero', () => {
    expect(listAllows(['0'], '')).toBe(false);
    expect(listAllows(['0'], '   ')).toBe(false);
    expect(listAllows(['', 'a'], '0')).toBe(false);
  });

  it('does not read hex, binary or octal as the number they stand for', () => {
    // The editor stores these as text, which is not one of the options.
    expect(listAllows(['16'], '0x10')).toBe(false);
    expect(listAllows(['2'], '0b10')).toBe(false);
    expect(listAllows(['8'], '0o10')).toBe(false);
  });

  it("falls back to Excel's own wording when a rule has no message", () => {
    expect(english.text(DEFAULT_LIST_ERROR)).toMatch(/^This value doesn't match the data validation restrictions/);
  });
});
