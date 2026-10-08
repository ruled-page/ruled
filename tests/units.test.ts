import { describe, expect, it } from 'vitest';
import { planFill } from '../src/core/fill';
import { displayFormula, engineFormula, expandLets, makeShifter, toFileFormula, validateFormula } from '../src/core/formula';
import { parseInput } from '../src/core/input';
import { decodeOoxmlText, encodeOoxmlText } from '../src/core/xml';
import { clipboardNumbers, parseClipboardText } from '../src/state/clipboard';

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

  it('writes LET out for the engine, each name as its value in brackets, the innermost first', () => {
    expect(expandLets('_xlfn.LET(_xlpm.revenue,F9,_xlpm.cost,_xlpm.revenue*0.6,_xlpm.revenue-_xlpm.cost)')).toBe('((F9)-((F9)*0.6))');
    // Written without Excel's prefix, as some files have it.
    expect(expandLets('LET(x, A1*2, x + x)')).toBe('((A1*2) + (A1*2))');
    // A nested LET's own x is its own.
    expect(expandLets('LET(x,1,LET(x,2,x)+x)')).toBe('(((2))+(1))');
    expect(expandLets('SUM(A1:A3)')).toBe('SUM(A1:A3)');
    expect(engineFormula('_xlfn.LET(_xlpm.n,A1,n*2)')).toBe('=((A1)*2)');
  });

  it('leaves a LET for Excel when it is volatile, malformed or too long once written out', () => {
    expect(expandLets('LET(x,RAND(),x+x)')).toBeUndefined();
    expect(expandLets('LET(x,1)')).toBeUndefined();
    expect(expandLets('LET(1,2,3)')).toBeUndefined();
    const doubling = Array.from({ length: 14 }, (_, i) => `x${i + 1},x${i}+x${i}`).join(',');
    expect(expandLets(`LET(x0,A1,${doubling},x14)`)).toBeUndefined();
  });
});

describe('typed input', () => {
  it('parses like Excel', () => {
    expect(parseInput('12%', 'General')).toEqual({ content: { kind: 'number', value: 0.12 }, numFmt: '0%' });
    // A comma after the decimal point is no thousands separator: Excel keeps it as text.
    expect(parseInput('1.234,5', 'General', 'en-US')).toEqual({ content: { kind: 'string', value: '1.234,5' } });
    expect(parseInput('1.234,5', 'General', 'de-DE').content).toEqual({ kind: 'number', value: 1234.5 });
    expect(parseInput('1,234.5', 'General', 'en-US').content).toEqual({ kind: 'number', value: 1234.5 });
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

describe('fill by dates and quarters', () => {
  const date = (y: number, m: number, d: number, base = Date.UTC(1899, 11, 30)) => ({
    content: { kind: 'number' as const, value: (Date.UTC(y, m - 1, d) - base) / 86_400_000 },
    isDate: true,
  });
  const iso = (serial: unknown, base = Date.UTC(1899, 11, 30)) => new Date(base + (serial as number) * 86_400_000).toISOString().slice(0, 10);
  const days = (plans: ReturnType<typeof planFill>, base?: number) => plans.map((p) => (p.kind === 'value' ? iso((p.content as { value: number }).value, base) : 'copy'));
  const text = (plans: ReturnType<typeof planFill>) => plans.map((p) => (p.kind === 'value' ? (p.content as { value: string }).value : 'copy'));
  const s = (value: string) => ({ content: { kind: 'string' as const, value }, isDate: false });

  it('steps by months when the dates fall on the same day of each month', () => {
    expect(days(planFill([date(2026, 1, 1), date(2026, 2, 1)], [2, 3, 4]))).toEqual(['2026-03-01', '2026-04-01', '2026-05-01']);
    // The last day of each month stays the last day.
    expect(days(planFill([date(2026, 1, 31), date(2026, 2, 28)], [2, 3]))).toEqual(['2026-03-31', '2026-04-30']);
    // A day that some months lack is cut to their last day.
    expect(days(planFill([date(2025, 12, 30), date(2026, 1, 30)], [2, 3]))).toEqual(['2026-02-28', '2026-03-30']);
    expect(days(planFill([date(2025, 1, 15), date(2026, 1, 15)], [2]))).toEqual(['2027-01-15']);
    expect(days(planFill([date(2026, 3, 1), date(2026, 4, 1)], [-1, -2]))).toEqual(['2026-02-01', '2026-01-01']);
    // Weekly dates still step by days.
    expect(days(planFill([date(2026, 1, 1), date(2026, 1, 8)], [2]))).toEqual(['2026-01-15']);
  });

  it('steps by months in the 1904 date system too', () => {
    const base = Date.UTC(1904, 0, 1);
    expect(days(planFill([date(2026, 1, 1, base), date(2026, 2, 1, base)], [2], true), base)).toEqual(['2026-03-01']);
  });

  it('goes round the quarters', () => {
    expect(text(planFill([s('Qtr3')], [1, 2, 3]))).toEqual(['Qtr4', 'Qtr1', 'Qtr2']);
    expect(text(planFill([s('Q1'), s('Q2')], [2, 3, 4]))).toEqual(['Q3', 'Q4', 'Q1']);
    expect(text(planFill([s('Quarter 4')], [1]))).toEqual(['Quarter 1']);
    // Other numbered text keeps counting.
    expect(text(planFill([s('Item 4')], [1]))).toEqual(['Item 5']);
  });
});

describe('typed input in other conventions', () => {
  it('reads dates into the 1904 system when the workbook uses it', () => {
    expect(parseInput('2026-03-01', 'General', 'en-US', true).content).toEqual({ kind: 'number', value: 46082 - 1462 });
    expect(parseInput('3/1/2026', 'General', 'en-US', true).content).toEqual({ kind: 'number', value: 46082 - 1462 });
    // Times and plain numbers are the same in both systems.
    expect(parseInput('12:30', 'General', 'en-US', true).content).toEqual({ kind: 'number', value: 0.5208333333333334 });
    expect(parseInput('45000', 'yyyy-mm-dd', 'en-US', true).content).toEqual({ kind: 'number', value: 45000 });
  });

  it('reads amounts with a currency symbol and a decimal comma', () => {
    const value = (text: string) => parseInput(text, 'General', 'de-DE').content;
    expect(value('1.234,50 €')).toEqual({ kind: 'number', value: 1234.5 });
    expect(value('€ 1.234,50')).toEqual({ kind: 'number', value: 1234.5 });
    expect(value('-1.234,50 €')).toEqual({ kind: 'number', value: -1234.5 });
    expect(value('1.234 €')).toEqual({ kind: 'number', value: 1234 });
    expect(value('12,5 %')).toEqual({ kind: 'number', value: 0.125 });
    expect(parseInput('1.234,50 €', 'General', 'de-DE').numFmt).toMatch(/€/);
    // Not an amount.
    expect(value('€ 1,5 €')).toEqual({ kind: 'string', value: '€ 1,5 €' });
  });
});

describe('clipboard HTML from Excel', () => {
  it('gives the exact value of each number, laid out like the text', () => {
    const html =
      '<html><body><table><thead><tr><th>Item</th><th>Amount</th></tr></thead><tbody>' +
      '<tr><td>Bolts</td><td class=xl65 align=right x:num="1234.5678">1,234.57</td></tr>' +
      '<tr><td colspan=2>Merged</td><td x:num>7</td></tr>' +
      '<tr><td rowspan=2 x:num="0.333333333333333">0.33</td><td x:num=\'2\'>2</td></tr>' +
      '<tr><td x:num="3">3</td></tr>' +
      '</tbody></table></body></html>';
    expect(clipboardNumbers(html)).toEqual([
      [null, null],
      [null, 1234.5678],
      [null, null, null],
      [0.333333333333333, 2],
      [null, 3],
    ]);
    expect(clipboardNumbers('<table><tr><td>1</td></tr></table>')).toBeUndefined();
  });
});
