import { describe, expect, it } from 'vitest';
import { valuesMatch, type EngineSheet } from '@ruled/formula/engine';
import { describeValue, toRaw, unsupportedExplanation, unsupportedLabel } from '../src/core/engine';
import { nativeEngine } from '@ruled/formula/native';

const unsupportedReason = (f: string) => nativeEngine.unsupportedReason(f);
import { isUnsupportedError } from '@ruled/formula/engine';
import { english } from '../src/i18n';

const sheet = (name: string, rows: EngineSheet['rows'] = []): EngineSheet => ({ name, rows });

/** A one-sheet engine; returns what a formula works out to there. */
function calc(formula: string, date1904 = false) {
  const e = nativeEngine.create([sheet('S', [[1, 2, 3]])], [], date1904);
  try {
    return e.calculate('=' + formula, 0);
  } finally {
    e.destroy();
  }
}

describe('formulas the engine cannot calculate', () => {
  it('are recognised before they are tried, with a reason for each', () => {
    const why = (f: string) => {
      const u = unsupportedReason(f);
      return u && [english.text(unsupportedLabel(u)), english.text(unsupportedExplanation(u))];
    };
    expect(why('SUM(A1:A3)')).toBeUndefined();
    expect(why('[1]Prices!A1*2')?.[1]).toBe("Uses links to other workbooks, which this editor can't calculate.");
    expect(why('Sales[Amount]')?.[1]).toBe("Uses table references (Table[Column]), which this editor can't calculate.");
    expect(why('_xludf.MYFN(1)')?.[1]).toBe("Uses the MYFN function, which this editor can't calculate.");
    expect(why('SUM(1,')?.[1]).toBe("Uses formula syntax this editor can't read.");
  });
});

describe('formulas written as Excel writes them', () => {
  it('calculate with upper-case exponents and quotes doubled inside text', () => {
    const cells = ['1E+3', '9.99E+307>1', 'LOOKUP(9.99E+307,A2:A3)', 'LEN("say ""hi""")', '"""x"""', 'COUNTIF(A2:A3,"a""b")', 'TRUE+1E-1'];
    for (const f of cells) expect(unsupportedReason(f), f).toBeUndefined();
    const e = nativeEngine.create([sheet('S', [cells.map((f) => toRaw({ kind: 'formula', formula: f })), [3], ['a"b']])], [], false);
    expect(cells.map((_, c) => e.value(0, 0, c))).toEqual([1000, true, 3, 8, '"x"', 1, 1.1]);
    e.destroy();
  });

  it('read a doubled quote in an array constant', () => {
    expect(unsupportedReason('INDEX({"a""b"},1)')).toBeUndefined();
    expect(calc('INDEX({"a""b"},1)')).toBe('a"b');
  });
});

describe('the engine', () => {
  it('skips names it cannot hold: print areas, broken references and scopes that are gone', () => {
    const e = nativeEngine.create(
      [sheet('S', [[5]]), null],
      [
        { name: '_xlnm.Print_Area', formula: 'S!$A$1:$B$2' },
        { name: 'Broken', formula: '#REF!' },
        { name: 'Orphan', formula: 'S!$A$1', localSheet: 1 },
        { name: 'Five', formula: 'S!$A$1' },
        { name: 'Bad name!', formula: 'S!$A$1' },
      ],
      false,
    );
    expect(e.calculate('=Five*2', 0)).toBe(10);
    expect(e.calculate('=Orphan', 0)).toMatchObject({ error: '#NAME?' });
    e.destroy();
  });

  it('answers nothing for sheets it was not given, until they are added', () => {
    const e = nativeEngine.create([sheet('S', [['=Other!A1*2']]), null], [], false);
    expect(e.has(1)).toBe(false);
    expect(e.value(1, 0, 0)).toBeNull();
    expect(e.value(7, 0, 0)).toBeNull();
    expect(e.calculate('=1', 1)).toEqual({ error: '#REF!' });
    expect(e.write([{ sheet: 1, row: 0, col: 0, raw: 3 }])).toEqual([]);
    expect(e.write([])).toEqual([]);
    // Adding the sheet recalculates the formulas that read it.
    expect(e.addSheet(1, sheet('Other', [[21]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 42 }]);
    expect(e.addSheet(1, sheet('Other', [[0]]))).toEqual([]);
    expect(e.value(1, 0, 0)).toBe(21);
    e.destroy();
  });

  it('gives #VALUE! for a formula it cannot read, which the session marks unsupported before it gets there', () => {
    const e = nativeEngine.create([sheet('S', [[1]])], [], false);
    const [change] = e.write([{ sheet: 0, row: 0, col: 1, raw: '=SUM(1,' }]);
    expect(change).toMatchObject({ sheet: 0, row: 0, col: 1, value: { error: '#VALUE!' } });
    expect(unsupportedReason('SUM(1,')).toEqual({ kind: 'syntax' });
    // The session's mark for what is left to Excel.
    expect(isUnsupportedError({ error: '#ERROR!' })).toBe(true);
    expect(isUnsupportedError({ error: '#NAME?' })).toBe(false);
    e.destroy();
  });

  it('reports a formula it cannot evaluate as an error rather than throwing', () => {
    const e = nativeEngine.create([sheet('S')], [], false);
    expect(e.calculate('=SUM(', 0)).toMatchObject({ error: expect.stringMatching(/^#/) });
    e.destroy();
  });

  it('finds the formulas that depend on a cell, through ranges too', () => {
    const e = nativeEngine.create([sheet('S', [[1, '=A1*2', '=SUM(A1:B1)', '=C1+1']])], [], false);
    expect(e.dependentsOf([{ sheet: 0, row: 0, col: 0 }, { sheet: 5, row: 0, col: 0 }]).map((c) => c.col).sort()).toEqual([1, 2, 3]);
    expect(e.dependentsOf([{ sheet: 0, row: 0, col: 0 }], 1).length).toBeLessThanOrEqual(2);
    e.destroy();
  });
});

describe('values', () => {
  it('match when Excel saved nothing and the engine works out empty or zero', () => {
    expect(valuesMatch('', 0)).toBe(true);
    expect(valuesMatch(null, '')).toBe(true);
    expect(valuesMatch(null, 1)).toBe(false);
    expect(valuesMatch({ error: '#N/A' }, { error: '#N/A' })).toBe(true);
    expect(valuesMatch({ error: '#N/A' }, null)).toBe(false);
    expect(valuesMatch(1e12, 1e12 + 1)).toBe(true);
  });

  it('are described as messages show them', () => {
    expect([null, 'a "b"', { error: '#DIV/0!' }, true, false, 1.5].map((v) => english.text(describeValue(v)))).toEqual(['(empty)', '"a \\"b\\""', '#DIV/0!', 'TRUE', 'FALSE', '1.5']);
  });

  it('go to the engine with text protected from being read as numbers', () => {
    expect(toRaw({ kind: 'string', value: '007' })).toBe("'007");
    expect(toRaw({ kind: 'error', value: '#N/A' })).toBe('#N/A');
    expect(toRaw({ kind: 'bool', value: true })).toBe(true);
    expect(toRaw({ kind: 'empty' })).toBeNull();
  });
});

describe('functions that follow Excel rather than the engine', () => {
  it('DATEDIF counts in every unit Excel has', () => {
    const d = (unit: string, a = 'DATE(2025,6,15)', b = 'DATE(2026,3,5)') => calc(`DATEDIF(${a},${b},"${unit}")`);
    expect(d('Y')).toBe(0);
    expect(d('M')).toBe(8);
    expect(d('D')).toBe(263);
    expect(d('YM')).toBe(8);
    expect(d('MD')).toBe(18);
    expect(d('MD', 'DATE(2026,1,2)')).toBe(3);
    expect(d('YD')).toBe(263);
    expect(d('YD', 'DATE(2025,1,10)')).toBe(54);
    expect(d('W')).toEqual(expect.objectContaining({ error: '#NUM!' }));
    expect(d('D', 'DATE(2026,3,6)')).toEqual(expect.objectContaining({ error: '#NUM!' }));
  });

  it('TEXT reads dates in the 1904 system and rejects formats it cannot use', () => {
    expect(calc('TEXT(0,"yyyy-mm-dd")', true)).toBe('1904-01-01');
    // Excel's day zero in the 1900 system.
    expect(calc('TEXT(0,"yyyy-mm-dd")')).toBe('1900-01-00');
    expect(calc('TEXT(1,"0;0;0;0;0")')).toEqual(expect.objectContaining({ error: '#VALUE!' }));
  });

  it('ROUND leaves zero alone', () => {
    expect(calc('ROUND(0,2)')).toBe(0);
    expect(calc('ROUNDUP(-0.0001,2)')).toBe(-0.01);
  });
});

describe('precision', () => {
  it('keeps the doubles Excel keeps, comparing them to 15 significant digits as Excel does', () => {
    const e = nativeEngine.create([sheet('S', [[2234567890.12, 0.1, 0.2]])], [], false);
    expect(e.calculate('=A1*1', 0)).toBe(2234567890.12);
    expect(e.calculate('=ROUND(88765432.1234567,4)', 0)).toBe(88765432.1235);
    expect(e.calculate('=123456789012345*1', 0)).toBe(123456789012345);
    expect(e.calculate('=B1+C1', 0)).toBe(0.1 + 0.2);
    expect(e.calculate('=B1+C1=0.3', 0)).toBe(true);
    expect(e.calculate('=1/3', 0)).toBeCloseTo(1 / 3, 14);
    e.destroy();
  });
});
