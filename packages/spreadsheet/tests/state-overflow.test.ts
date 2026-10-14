import { describe, expect, it } from 'vitest';
import type { DisplayStyle, SheetInfo } from '../src/core/types';
import { CELL_PADDING, excelTextWidth, fitNumber, indentPx, overflowTextX, rowOverflows, textTop, type Overflow } from '../src/state/overflow';
import { sheet, style, type FileCell } from './state-fixtures';

/** 7px a character; columns are 64px, so a cell holds 7 characters after padding. */
const measure = (text: string) => text.length * 7;
const TWENTY = 'twenty characters!!!';

function spills(cells: FileCell[], styles: DisplayStyle[] = [style()], extra: Partial<SheetInfo> = {}, row = 0) {
  const view = sheet(cells, extra);
  return rowOverflows(view, row, (_col, rec) => styles[rec.style] ?? styles[0], false, measure).map((o) => [o.col, o.from, o.to, o.align]);
}

const left = [style()];
const right = [style({ hAlign: 'right' })];
const centre = [style({ hAlign: 'center' })];

describe('text running into empty neighbours', () => {
  it('runs left-aligned text right until it fits', () => {
    // 140px of text plus padding needs 85px more than its own 64px: two more columns.
    expect(spills([[0, 1, TWENTY]], left)).toEqual([[1, 1, 3, 'left']]);
    expect(spills([[0, 1, 'twelve chars']], left)).toEqual([[1, 1, 2, 'left']]);
  });

  it('runs right-aligned text left, and stops at the first column', () => {
    expect(spills([[0, 4, TWENTY]], right)).toEqual([[4, 2, 4, 'right']]);
    expect(spills([[0, 1, TWENTY]], right)).toEqual([[1, 0, 1, 'right']]);
    expect(spills([[0, 0, TWENTY]], right)).toEqual([]);
  });

  it('runs centred text both ways, half the extra each side', () => {
    expect(spills([[0, 4, TWENTY]], centre)).toEqual([[4, 3, 5, 'center']]);
    // Blocked on one side, it still runs into the other.
    expect(spills([[0, 4, TWENTY], [0, 3, 'x']], centre)).toEqual([[4, 4, 5, 'center']]);
  });

  it('stops at a cell with a value or a merged cell', () => {
    expect(spills([[0, 1, TWENTY], [0, 2, 'x']], left)).toEqual([]);
    expect(spills([[0, 1, TWENTY], [0, 3, 0]], left)).toEqual([[1, 1, 2, 'left']]);
    expect(spills([[0, 1, TWENTY]], left, { merges: [{ r1: 0, c1: 3, r2: 1, c2: 4 }] })).toEqual([[1, 1, 2, 'left']]);
  });

  it('steps over hidden columns, which take no room', () => {
    expect(spills([[0, 1, 'twelve chars']], left, { hiddenCols: [2] })).toEqual([[1, 1, 3, 'left']]);
  });

  it('stops at the last column of the sheet', () => {
    const view = sheet([[0, 24, 'x'.repeat(200)]]);
    const last = view.visibleCols[view.colCount - 1];
    expect(rowOverflows(view, 0, () => style(), false, measure).map((o) => [o.col, o.from, o.to])).toEqual([[24, 24, last]]);
  });

  it('never spills numbers, wrapped, filled or justified text, merged cells or text in hidden columns', () => {
    expect(spills([[0, 1, 1234567890123]], left)).toEqual([]);
    for (const hAlign of ['fill', 'justify']) expect(spills([[0, 1, TWENTY]], [style({ hAlign })]), hAlign).toEqual([]);
    expect(spills([[0, 1, TWENTY]], [style({ wrap: true })])).toEqual([]);
    expect(spills([[0, 1, TWENTY]], left, { merges: [{ r1: 0, c1: 1, r2: 0, c2: 2 }] })).toEqual([]);
    expect(spills([[0, 1, TWENTY]], left, { hiddenCols: [1] })).toEqual([]);
  });

  it('leaves text alone that fits, or that shows nothing', () => {
    expect(spills([[0, 1, 'short'], [0, 2, '']], left)).toEqual([]);
    // Eight characters and the padding fill the column exactly; a ninth runs over.
    const exact = { defaultColWidthPx: 8 * 7 + CELL_PADDING * 2 };
    expect(spills([[0, 1, 'x'.repeat(8)]], left, exact)).toEqual([]);
    expect(spills([[0, 1, 'x'.repeat(9)]], left, exact)).toEqual([[1, 1, 2, 'left']]);
  });

  it('measures the text as its number format shows it', () => {
    const view = sheet([[0, 1, 'id', 1]]);
    const styles = [style(), style({ numFmt: '"Customer reference: "@' })];
    const [o] = rowOverflows(view, 0, (_c, rec) => styles[rec.style], false, measure);
    expect(o).toMatchObject({ text: 'Customer reference: id', width: 22 * 7, style: styles[1], align: 'left' });
  });

  it('only looks at the row asked for, and uses custom column widths', () => {
    const cells: FileCell[] = [
      [0, 1, TWENTY],
      [1, 1, TWENTY],
    ];
    expect(spills(cells, left, { colWidths: { 1: 200 } }, 1)).toEqual([]);
    expect(spills(cells, left, { colWidths: { 2: 200 } }, 1)).toEqual([[1, 1, 2, 'left']]);
  });
});

describe('where spilled text starts', () => {
  const o = (align: Overflow['align']): Overflow => ({ col: 0, from: 0, to: 1, text: 'x', style: style(), align, width: 100 });

  it('pads left-aligned text from its own left edge', () => {
    expect(overflowTextX(o('left'), 64)).toBe(CELL_PADDING);
  });

  it('ends right-aligned text at its own right edge', () => {
    expect(overflowTextX(o('right'), 64)).toBe(64 - CELL_PADDING - 100);
  });

  it('centres centred text on its own cell', () => {
    expect(overflowTextX(o('center'), 64)).toBe(32 - 50);
  });
});

describe('fitting numbers into narrow columns', () => {
  it('leaves numbers that fit as they are', () => {
    expect(fitNumber('1,234.50', 1234.5, false, 100, measure)).toBe('1,234.50');
  });

  it('drops decimals one at a time before going scientific', () => {
    expect(fitNumber('-2.71828', -2.71828, true, 7 * 5, measure)).toBe('-2.72');
    expect(fitNumber('99.96', 99.96, true, 7 * 3, measure)).toBe('100');
  });

  it('goes straight to scientific notation for very large and very small numbers', () => {
    expect(fitNumber('123456789012', 123456789012, true, 7 * 8, measure)).toBe('1.23E+11');
    expect(fitNumber('0.00001234', 0.00001234, true, 7 * 8, measure)).toBe('1.23E-05');
    expect(fitNumber('1E+100', 1e100, true, 7 * 6, measure)).toBe('1E+100');
  });

  it('uses scientific notation for whole numbers with no decimals to drop', () => {
    expect(fitNumber('123456789', 123456789, true, 7 * 5, measure)).toBe('1E+08');
  });

  it('fills the cell with # when nothing fits, always at least one', () => {
    expect(fitNumber('123456789', 123456789, true, 7 * 4, measure)).toBe('####');
    expect(fitNumber('12', 12, false, 3, measure)).toBe('#');
    expect(fitNumber('Infinity', Infinity, true, 7 * 3, measure)).toBe('###');
    // A measure that gives # no width still yields a count, not a division by zero.
    expect(fitNumber('12', 12, false, 3.5, (s) => (s === '#' ? 0 : 99))).toBe('###');
  });
});

describe("Excel's widths for numbers", () => {
  const plain = style();

  it('adds up Calibri character widths at 11pt', () => {
    expect(excelTextWidth('1,234.5', plain)).toBeCloseTo(7 * 5 + 3.6 * 2);
    expect(excelTextWidth('50%', plain)).toBeCloseTo(7 * 2 + 10.4);
    expect(excelTextWidth('1.2E+05', plain)).toBeCloseTo(7 * 4 + 3.6 + 6.4 + 7);
    expect(excelTextWidth('ab', plain)).toBeCloseTo(6.5 * 2);
    expect(excelTextWidth('', plain)).toBe(0);
  });

  it('scales with the font size, and a little for bold', () => {
    expect(excelTextWidth('12', style({ fontSizePx: 29.4 }))).toBeCloseTo(28);
    expect(excelTextWidth('12', style({ bold: true }))).toBeCloseTo(14 * 1.04);
  });
});

describe('indents and vertical alignment', () => {
  it('indent left- and right-aligned text by three spaces a level, never centred text', () => {
    expect([indentPx(style({ indent: 2 }), 'left'), indentPx(style({ indent: 2 }), 'right'), indentPx(style({ indent: 2 }), 'center')]).toEqual([18, 18, 0]);
    const o = (align: Overflow['align']): Overflow => ({ col: 0, from: 0, to: 1, text: 'x', style: style({ indent: 1 }), align, width: 100 });
    expect([overflowTextX(o('left'), 64), overflowTextX(o('right'), 64)]).toEqual([CELL_PADDING + 9, 64 - CELL_PADDING - 100 - 9]);
    // An indent pushes text further into the next cells: "twelve chars" fits in two columns, but not with six levels.
    expect(spills([[0, 1, 'twelve chars']], [style({ indent: 6 })])).toEqual([[1, 1, 3, 'left']]);
  });

  it("put text at Excel's top, middle or bottom, bottom unless set", () => {
    // A 40px row, 12px of text and 2px of padding.
    expect([textTop('top', 40, 12, 2), textTop('center', 40, 12, 2), textTop(undefined, 40, 12, 2)]).toEqual([2, 14, 26]);
    expect(textTop('distributed', 40, 12, 2)).toBe(14);
    // Text taller than its cell: at the top its first lines show, at the bottom its last.
    expect([textTop('top', 20, 30, 2), textTop('center', 20, 30, 2), textTop('bottom', 20, 30, 2)]).toEqual([2, -5, -12]);
  });
});
