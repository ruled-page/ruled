import { describe, expect, it } from 'vitest';
import type { EngineSheet } from '../src/api';
import { engineFormula } from '../src/formula';
import { nativeEngine } from '../src/native';

const sheet = (name: string, rows: EngineSheet['rows'] = []): EngineSheet => ({ name, rows });

describe('three-dimensional references and unions in the native engine', () => {
  it('are calculated, and worked out again when a cell they read changes', () => {
    expect(nativeEngine.unsupportedReason('SUM(Jan:Mar!B2)')).toBeUndefined();
    const e = nativeEngine.create(
      [sheet('S', [[engineFormula('SUM(Jan:Mar!A1)'), engineFormula('SUM((Feb!A1:A2,Feb!B1))')]]), sheet('Jan', [[1]]), sheet('Feb', [[2, 5], [3]]), sheet('Mar', [[4]])],
      [],
      false,
    );
    expect([e.value(0, 0, 0), e.value(0, 0, 1)]).toEqual([7, 10]);
    const changes = e.write([{ sheet: 2, row: 0, col: 0, raw: 20 }]);
    expect(changes).toContainEqual({ sheet: 0, row: 0, col: 0, value: 25 });
    expect(changes).toContainEqual({ sheet: 0, row: 0, col: 1, value: 28 });
    e.destroy();
  });

  it('read a sheet that arrives within their run of sheets', () => {
    const e = nativeEngine.create([sheet('S', [[engineFormula('SUM(Jan:Mar!A1)')]]), sheet('Jan', [[1]]), null, sheet('Mar', [[4]])], [], false);
    expect(e.value(0, 0, 0)).toBe(5);
    expect(e.addSheet(2, sheet('Feb', [[2]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 7 }]);
    e.destroy();
  });
});

describe('a sheet that arrives after the formulas naming it', () => {
  it('is read by them', () => {
    const e = nativeEngine.create([sheet('S', [[engineFormula('Other!A1*2')]]), null], [], false);
    expect(e.value(0, 0, 0)).toEqual({ error: '#REF!' });
    expect(e.addSheet(1, sheet('Other', [[21]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 42 }]);
    e.destroy();
  });
});

describe('a name qualified with a sheet', () => {
  it("reads the name local to that sheet, else the workbook's, from any sheet", () => {
    expect(nativeEngine.unsupportedReason('Prices!Rate*2')).toBeUndefined();
    expect(nativeEngine.unsupportedReason('[1]Prices!Rate')).toEqual({ kind: 'externalLinks' });
    const e = nativeEngine.create(
      [
        sheet('Prices', [[5, engineFormula('Rate*2'), engineFormula('Prices!Rate*2')]]),
        sheet('Q1 Sales', [[7, engineFormula("'Q1 Sales'!Rate"), engineFormula('Prices!Rate+Prices!Tax'), engineFormula('Prices!Nothing'), engineFormula('Nowhere!Rate')]]),
      ],
      [
        { name: 'Rate', formula: 'Prices!$A$1', localSheet: 0 },
        { name: 'Rate', formula: "'Q1 Sales'!$A$1", localSheet: 1 },
        { name: 'Tax', formula: '0.5' },
      ],
      false,
    );
    expect([e.value(0, 0, 1), e.value(0, 0, 2)]).toEqual([10, 10]);
    expect([e.value(1, 0, 1), e.value(1, 0, 2), e.value(1, 0, 3), e.value(1, 0, 4)]).toEqual([7, 5.5, { error: '#NAME?' }, { error: '#REF!' }]);
    // It follows the cells the name refers to, and the name's definition.
    expect(e.write([{ sheet: 0, row: 0, col: 0, raw: 6 }])).toContainEqual({ sheet: 1, row: 0, col: 2, value: 6.5 });
    expect(e.setName('Rate', 0, 'Prices!$A$1*10').changes).toContainEqual({ sheet: 1, row: 0, col: 2, value: 60.5 });
    e.destroy();
  });

  it('reads a sheet that arrives after it', () => {
    const e = nativeEngine.create([sheet('S', [[engineFormula('Later!Rate+1')]])], [{ name: 'Rate', formula: 'Later!$A$1' }], false);
    expect(e.value(0, 0, 0)).toEqual({ error: '#REF!' });
    expect(e.addSheet(1, sheet('Later', [[41]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 42 }]);
    e.destroy();
  });
});

describe('CELL in the native engine', () => {
  it('is calculated for what the cell itself gives, and left to Excel for the rest', () => {
    const calculated = (f: string) => nativeEngine.unsupportedReason(f) === undefined;
    expect(['CELL("row",A1)', 'CELL( "Contents" , B2:C3)', 'CELL("type",Sheet2!A1)', 'CELL("address",A1)'].every(calculated)).toBe(true);
    // Formats, widths and protection; the last cell changed (no reference); an address with a file name; a kind not written out.
    expect(['CELL("format",A1)', 'CELL("width",A1)', 'CELL("protect",A1)', 'CELL("row")', 'CELL("address",Sheet2!A1)', 'CELL(B1,A1)'].some(calculated)).toBe(false);
  });
});
