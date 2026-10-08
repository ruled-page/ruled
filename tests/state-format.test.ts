import { describe, expect, it } from 'vitest';
import { VT } from '../src/core/types';
import { alignment, cellJson, displayText, editText, exactText, isDateStyle } from '../src/state/format';
import { bool, err, num, str, style } from './state-fixtures';

const general = style();
const isoDate = style({ numFmt: 'yyyy-mm-dd' });
/** 1 January 2026 in each date system. */
const NEW_YEAR_1900 = 46023;
const NEW_YEAR_1904 = NEW_YEAR_1900 - 1462;

describe('telling date formats apart', () => {
  it('recognises date and time patterns, and nothing else', () => {
    expect(isDateStyle(isoDate)).toBe(true);
    expect(isDateStyle(style({ numFmt: '[h]:mm' }))).toBe(true);
    expect(isDateStyle(style({ numFmt: 'mmm' }))).toBe(true);
    expect(isDateStyle(general)).toBe(false);
    expect(isDateStyle(style({ numFmt: '0.00%' }))).toBe(false);
    expect(isDateStyle(style({ numFmt: '[' }))).toBe(false);
  });
});

describe('the text a cell shows', () => {
  it('formats numbers with their pattern, and General as Excel does', () => {
    expect(displayText(num(1234.5), style({ numFmt: '#,##0.00' }), false)).toEqual({ text: '1,234.50', color: undefined });
    expect(displayText(num(0.1 + 0.2), general, false).text).toBe('0.3');
    expect(displayText(num(123456789012345), general, false).text).toBe('1.23457E+14');
    expect(displayText(num(0.125), style({ numFmt: '0.0%' }), false).text).toBe('12.5%');
    // An empty pattern is General.
    expect(displayText(num(2), style({ numFmt: '' }), false).text).toBe('2');
  });

  it('treats a number cell with no value as zero', () => {
    expect(displayText({ t: VT.Number, style: 0, flags: 0 }, style({ numFmt: '0.00' }), false).text).toBe('0.00');
  });

  it('passes on the colour a number format asks for', () => {
    const redNegatives = style({ numFmt: '0.00;[Red]-0.00' });
    expect(displayText(num(-1.5), redNegatives, false)).toEqual({ text: '-1.50', color: 'red' });
    expect(displayText(num(1.5), redNegatives, false)).toEqual({ text: '1.50', color: undefined });
  });

  it('reads serials in the 1900 and 1904 date systems as the same day', () => {
    expect(displayText(num(NEW_YEAR_1900), isoDate, false).text).toBe('2026-01-01');
    expect(displayText(num(NEW_YEAR_1904), isoDate, true).text).toBe('2026-01-01');
    expect(displayText(num(NEW_YEAR_1900 + 0.75), style({ numFmt: 'd/m/yyyy h:mm' }), false).text).toBe('1/1/2026 18:00');
    // Only dates shift: a plain number means the same in either system.
    expect(displayText(num(42), general, true).text).toBe('42');
  });

  it('shows text through a text format, and as it is otherwise', () => {
    expect(displayText(str('A-17'), style({ numFmt: '"Code: "@' }), false).text).toBe('Code: A-17');
    expect(displayText(str('plain'), style({ numFmt: '0.00' }), false).text).toBe('plain');
    expect(displayText(str('plain'), general, false).text).toBe('plain');
    expect(displayText({ t: VT.String, style: 0, flags: 0 }, general, false).text).toBe('');
  });

  it('shows logicals in capitals, errors as their code and empty cells as nothing', () => {
    expect(displayText(bool(true), general, false).text).toBe('TRUE');
    expect(displayText(bool(false), general, false).text).toBe('FALSE');
    expect(displayText(err('#DIV/0!'), general, false).text).toBe('#DIV/0!');
    expect(displayText(err(), general, false).text).toBe('#VALUE!');
    expect(displayText({ t: VT.None, style: 0, flags: 0 }, general, false).text).toBe('');
  });

  it('shows every stored digit when asked for exact numbers', () => {
    const twoPlaces = style({ numFmt: '0.00' });
    expect(displayText(num(3.14159265), twoPlaces, false).text).toBe('3.14');
    expect(displayText(num(3.14159265), twoPlaces, false, true).text).toBe('3.14159265');
    // Grouping follows the browser's locale, so compare the digits only.
    expect(displayText(num(1234567.125), twoPlaces, false, true).text.replace(/\D/g, '')).toBe('1234567125');
    // Dates and text keep their usual display.
    expect(displayText(num(NEW_YEAR_1900), isoDate, false, true).text).toBe('2026-01-01');
    expect(displayText(str('x'), twoPlaces, false, true).text).toBe('x');
  });
});

describe('exact numbers', () => {
  it("trims binary noise to Excel's 15 significant digits", () => {
    expect(exactText(num(0.1 + 0.2), general, false)).toBe('0.3');
    expect(exactText(num(1 / 3), style({ numFmt: '0.0' }), false)).toBe('0.333333333333333');
    expect(exactText({ t: VT.Number, style: 0, flags: 0 }, general, false)).toBe('0');
  });

  it('keeps percentages as percentages, but not a quoted percent sign', () => {
    expect(exactText(num(0.123456), style({ numFmt: '0%' }), false)).toBe('12.3456%');
    expect(exactText(num(5), style({ numFmt: '"%"0' }), false)).toBe('5');
  });

  it('groups thousands only when asked to', () => {
    expect(exactText(num(1234567.5), general, false)).toBe('1234567.5');
    expect(exactText(num(1234567.5), general, true)).toBe((1234567.5).toLocaleString(undefined, { maximumFractionDigits: 15 }));
  });

  it('has nothing to add for dates, text and other values', () => {
    expect(exactText(num(NEW_YEAR_1900), isoDate, false)).toBeUndefined();
    expect(exactText(str('12'), general, false)).toBeUndefined();
    expect(exactText(bool(true), general, false)).toBeUndefined();
  });
});

describe("Excel's default alignment", () => {
  it('puts numbers right, logicals and errors in the middle and text left', () => {
    expect(alignment(num(1), general)).toBe('right');
    expect(alignment(bool(true), general)).toBe('center');
    expect(alignment(err('#N/A'), general)).toBe('center');
    expect(alignment(str('a'), general)).toBe('left');
    expect(alignment({ t: VT.None, style: 0, flags: 0 }, general)).toBe('left');
  });

  it("follows the cell's own alignment over the default", () => {
    const aligned = (hAlign: string) => alignment(num(1), style({ hAlign }));
    expect(['center', 'centerContinuous', 'distributed'].map(aligned)).toEqual(['center', 'center', 'center']);
    expect(['left', 'fill', 'justify'].map(aligned)).toEqual(['left', 'left', 'left']);
    expect(alignment(str('a'), style({ hAlign: 'right' }))).toBe('right');
    // "general" (and anything unknown) falls back to the value's default.
    expect(aligned('general')).toBe('right');
  });
});

describe('what the cell editor and formula bar show', () => {
  it('is empty for an empty cell', () => {
    expect(editText(null, general, false)).toBe('');
  });

  it("shows formulas without the file's function prefixes", () => {
    expect(editText(num(6, { f: 'B2*C2' }), general, false)).toBe('=B2*C2');
    expect(editText(num(6, { f: '_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlfn._xlws.SORT(A1:A2))' }), general, false)).toBe('=LET(x,2,x*SORT(A1:A2))');
    expect(editText(num(1, { f: '_xlfn.XLOOKUP(A1,B:B,C:C)' }), general, false)).toBe('=XLOOKUP(A1,B:B,C:C)');
    expect(editText(num(1, { f: '_xlfn._xlws.SORT(A1:A3)+_XLWS.FILTER(A1:A3,B1:B3)' }), general, false)).toBe('=SORT(A1:A3)+FILTER(A1:A3,B1:B3)');
    expect(editText(str('', { f: '""' }), general, false)).toBe('=""');
  });

  it('shows numbers in full, percentages as percentages and dates as dates', () => {
    expect(editText(num(1234.5678), style({ numFmt: '#,##0.00' }), false)).toBe('1234.5678');
    expect(editText(num(0.075), style({ numFmt: '0%' }), false)).toBe('7.5%');
    expect(editText(num(NEW_YEAR_1904), isoDate, true)).toBe('2026-01-01');
    expect(editText({ t: VT.Number, style: 0, flags: 0 }, general, false)).toBe('');
  });

  it('marks text that would otherwise be read as a number, formula or logical', () => {
    for (const s of ['=1+1', '+44 20 7946 0000', '-5', '@home', '007', ' 1,234.50 ', '$5', '50%', 'true', 'FALSE']) {
      expect(editText(str(s), general, false), s).toBe("'" + s);
    }
    for (const s of ['abc', '12 apples', 'trueish', '']) expect(editText(str(s), general, false), s).toBe(s);
  });

  // Text cells holding '1/2/2026', '1E5', '(5)', '12:30', '#N/A' or 'TRUE ' are
  // shown without the apostrophe, so confirming the editor unchanged (or a
  // replace that leaves them as they were) stores a date, number, error or
  // logical instead. Excel itself only shows the apostrophe for cells with
  // the quote-prefix flag, so whether these should be marked is open.
  it.skip('marks text that would be read as a date, a number in other notation, an error or a padded logical', () => {
    for (const s of ['1/2/2026', '1E5', '(5)', '12:30', '#N/A', 'TRUE ']) expect(editText(str(s), general, false), s).toBe("'" + s);
  });

  it('shows logicals and errors as they are displayed', () => {
    expect(editText(bool(false), general, false)).toBe('FALSE');
    expect(editText(err('#REF!'), general, false)).toBe('#REF!');
  });
});

describe('cell values as JSON', () => {
  it('gives null for empty cells', () => {
    expect(cellJson(null, general, false)).toBeNull();
    expect(cellJson({ t: VT.None, style: 0, flags: 0 }, general, false)).toBeNull();
  });

  it('keeps numbers, logicals, text and errors as their own kinds', () => {
    expect(cellJson(num(12.5), style({ numFmt: '0.00' }), false)).toBe(12.5);
    expect(cellJson({ t: VT.Number, style: 0, flags: 0 }, general, false)).toBe(0);
    expect(cellJson(bool(true), general, false)).toBe(true);
    expect(cellJson(bool(false), general, false)).toBe(false);
    expect(cellJson(str('x'), general, false)).toBe('x');
    expect(cellJson({ t: VT.String, style: 0, flags: 0 }, general, false)).toBe('');
    expect(cellJson(err('#N/A'), general, false)).toBe('#N/A');
  });

  it('gives dates as ISO text, with the time only when there is one', () => {
    expect(cellJson(num(NEW_YEAR_1900), isoDate, false)).toBe('2026-01-01');
    expect(cellJson(num(NEW_YEAR_1904), isoDate, true)).toBe('2026-01-01');
    expect(cellJson(num(NEW_YEAR_1900 + 0.75), isoDate, false)).toBe('2026-01-01T18:00:00Z');
    // Rounded to the millisecond, which stays when there is one.
    expect(cellJson(num(NEW_YEAR_1900 + 0.5 + 9 / 86_400_000), isoDate, false)).toBe('2026-01-01T12:00:00.009Z');
  });
});
