import { describe, expect, it } from 'vitest';
import { planFill } from '../src/core/fill';
import { displayFormula, makeShifter, toFileFormula, validateFormula } from '../src/core/formula';
import { parseInput } from '../src/core/input';
import { decodeOoxmlText, encodeOoxmlText } from '../src/core/xml';
import { parseClipboardText } from '../src/state/clipboard';

describe('clipboard text', () => {
  it('splits Excel tab-separated rows and keeps quoted multi-line cells together', () => {
    expect(parseClipboardText('a\tb\r\n"x\ny"\t"say ""hi"""\r\n')).toEqual([
      ['a', 'b'],
      ['x\ny', 'say "hi"'],
    ]);
    expect(parseClipboardText('5" screen')).toEqual([['5" screen']]);
    expect(parseClipboardText('one')).toEqual([['one']]);
  });
});

describe('fill', () => {
  const n = (value: number) => ({ content: { kind: 'number' as const, value }, isDate: false });
  const s = (value: string) => ({ content: { kind: 'string' as const, value }, isDate: false });
  const values = (plans: ReturnType<typeof planFill>) => plans.map((p) => (p.kind === 'value' ? (p.content as { value: unknown }).value : `copy ${p.source}`));

  it('extends linear series and copies single numbers', () => {
    expect(values(planFill([n(1), n(3)], [2, 3, 4]))).toEqual([5, 7, 9]);
    expect(values(planFill([n(0.1), n(0.2)], [2]))).toEqual([0.3]);
    expect(values(planFill([n(5)], [1, 2]))).toEqual(['copy 0', 'copy 0']);
    expect(values(planFill([{ ...n(46000), isDate: true }], [1, 2]))).toEqual([46001, 46002]);
    expect(values(planFill([n(10), n(20)], [-1, -2]))).toEqual([0, -10]);
  });

  it('continues text with numbers, months and weekdays', () => {
    expect(values(planFill([s('Item 9')], [1, 2]))).toEqual(['Item 10', 'Item 11']);
    expect(values(planFill([s('Q01')], [1]))).toEqual(['Q02']);
    expect(values(planFill([s('Nov')], [1, 2]))).toEqual(['Dec', 'Jan']);
    expect(values(planFill([s('MONDAY')], [1]))).toEqual(['TUESDAY']);
    expect(values(planFill([s('a'), s('b')], [2, 3]))).toEqual(['copy 0', 'copy 1']);
  });
});

describe('formulas', () => {
  it('shifts relative references and keeps absolute ones', () => {
    const shift = makeShifter("SUM(A1:B2,$C$3,Sheet2!D4,'My Sheet'!E5)", 0, 0);
    expect(shift(2, 1)).toBe("SUM(B3:C4,$C$3,Sheet2!E6,'My Sheet'!F7)");
    expect(makeShifter('A1', 1, 1)(0, 0)).toBe('#REF!');
  });

  it('adds and hides the _xlfn prefix for newer functions', () => {
    expect(toFileFormula('XLOOKUP(1,A:A,B:B)+SUM(1)')).toBe('_xlfn.XLOOKUP(1,A:A,B:B)+SUM(1)');
    expect(toFileFormula('filter(A1:A9,B1:B9>0)')).toBe('_xlfn._xlws.filter(A1:A9,B1:B9>0)');
    expect(displayFormula('_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))')).toBe('=SORT(UNIQUE(A1:A3))');
  });

  it('rejects malformed formulas', () => {
    expect(validateFormula('SUM(A1')).toBeTruthy();
    expect(validateFormula('SUM(A1:A3)')).toBeUndefined();
  });
});

describe('typed input', () => {
  it('parses like Excel', () => {
    expect(parseInput('12%', 'General')).toEqual({ content: { kind: 'number', value: 0.12 }, numFmt: '0%' });
    expect(parseInput('007', '@')).toEqual({ content: { kind: 'string', value: '007' } });
    expect(parseInput("'=not a formula", 'General')).toEqual({ content: { kind: 'string', value: '=not a formula' } });
    expect(parseInput('#n/a', 'General')).toEqual({ content: { kind: 'error', value: '#N/A' } });
    expect(parseInput('true', 'General')).toEqual({ content: { kind: 'bool', value: true } });
  });
});

describe('OOXML text escapes', () => {
  it('round-trips control characters and literal escape sequences', () => {
    const text = 'tab\there\u0001 and _x0041_';
    expect(decodeOoxmlText(encodeOoxmlText(text))).toBe(text);
  });
});
