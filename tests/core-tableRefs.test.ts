import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { resolveTableRefs, type TablePlace } from '../src/core/tableRefs';
import { english } from '../src/i18n';

// Orders on Sheet 'My Lines'!B2:E8: headers in row 2, five rows of data, totals in row 8.
const orders: TablePlace = { name: 'Orders', sheet: 1, sheetName: 'My Lines', rect: { r1: 1, c1: 1, r2: 7, c2: 4 }, headerRows: 1, totalsRows: 1, columns: ['Item', 'Qty', 'Unit Price', 'Total'] };

describe('table references', () => {
  const at = (row: number, col: number, sheet = 1) => ({ sheet, row, col });
  const resolve = (f: string, where = at(3, 4)) => resolveTableRefs(f, where, [orders]);

  it('are written as the cells they stand for', () => {
    expect(resolve('SUM(Orders[Total])')).toBe("SUM('My Lines'!$E$3:$E$7)");
    expect(resolve('[@Qty]*[@[Unit Price]]')).toBe("'My Lines'!$C$4*'My Lines'!$D$4");
    expect(resolve('Orders[[#Totals],[Qty]]')).toBe("'My Lines'!$C$8");
    expect(resolve('ROWS(Orders[#All])')).toBe("ROWS('My Lines'!$B$2:$E$8)");
    expect(resolve('Orders[[#Headers],[Item]:[Qty]]')).toBe("'My Lines'!$B$2:$C$2");
    expect(resolve('SUM(Orders[[#Data],[#Totals],[Qty]])')).toBe("SUM('My Lines'!$C$3:$C$8)");
    expect(resolve('COUNTA(Orders[])', at(0, 0, 0))).toBe("COUNTA('My Lines'!$B$3:$E$7)");
    expect(resolve('SUM(A1:A3)')).toBe('SUM(A1:A3)');
  });

  it('are left for the engine to refuse when they point at nothing', () => {
    expect(resolve('Others[Total]')).toBeUndefined();
    expect(resolve('Orders[Price]')).toBeUndefined();
    // This row, from outside the table's data.
    expect(resolve('[@Qty]', at(1, 2))).toBeUndefined();
    expect(resolve('[@Qty]', at(3, 9))).toBeUndefined();
    // Not a table reference: links to other workbooks are refused elsewhere.
    expect(resolve('[Book.xlsx]Sheet1!A1')).toBe('[Book.xlsx]Sheet1!A1');
  });
});

describe('a workbook with an Excel table', () => {
  it('calculates formulas with table references, and recalculates them as the table changes', () => {
    const s = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))), 'templates.xlsx', 'en-US', { recalc: 'on' });
    const report = s.ensureEngine();
    const notCalculated = report.samples.map((x) => `${x.sheet}!${x.ref} ${english.text(x.reason)}`);
    // Only the dynamic array reference (A2#), LET and LAMBDA remain.
    expect(notCalculated.filter((x) => /table references/i.test(x))).toEqual([]);
    const lines = s.model.sheets.findIndex((m) => m.info.name === 'Lines');
    const r = s.apply({ type: 'input', sheet: lines, row: 2, col: 2, text: '400' });
    const at = (sheet: string, row: number, col: number) => r.updates.find((u) => s.model.sheets[u.sheet].info.name === sheet && u.row === row && u.col === col)?.n;
    // Nuts: 400 at 0.10, and the order total and the table's totals row follow.
    expect(at('Lines', 2, 4)).toBeCloseTo(40);
    expect(at('Request', 5, 1)).toBeCloseTo(383.5);
    expect(at('Lines', 6, 4)).toBeCloseTo(383.5);
  });
});
