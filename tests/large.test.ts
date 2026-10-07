import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { CellFlag, has, open, sample, sheetIndex, updateAt, valueOf } from './helpers';

describe('live recalculation off', () => {
  it('keeps saved values, marks typed formulas, and leaves custom rules to Excel', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx', 'en-US', { recalc: 'off' });
    const report = s.ensureEngine();
    expect(report.off).toBe(true);
    expect(report.findings[0].title).toBe('Live recalculation is off');
    const summary = sheetIndex(s, 'Summary');
    // B9 sums B5:B8; changing B5 doesn't recalculate it here.
    const r = s.apply({ type: 'input', sheet: summary, row: 4, col: 1, text: '1' });
    expect(r.updates.map((u) => [u.row, u.col])).toEqual([[4, 1]]);
    const f = s.apply({ type: 'input', sheet: summary, row: 30, col: 3, text: '=B5*2' });
    const cell = updateAt(f.updates, s, 'Summary', 'D31');
    expect(has(cell?.flags, CellFlag.Unsupported)).toBe(true);
    expect(s.conditionalFormats(sheetIndex(s, 'Data')).cells.length).toBeGreaterThan(0);
  });

  it('can be turned on later, bringing in what was typed meanwhile', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx', 'en-US', { recalc: 'off' });
    const summary = sheetIndex(s, 'Summary');
    s.ensureEngine();
    s.apply({ type: 'input', sheet: summary, row: 30, col: 3, text: '=2+3' });
    const report = s.enableRecalc();
    expect(report.off).toBeUndefined();
    expect(report.liveCells).toBeGreaterThan(0);
    const cell = updateAt(report.updates, s, 'Summary', 'D31');
    expect(valueOf(cell)).toBe(5);
    expect(has(cell?.flags, CellFlag.Unsupported)).toBe(false);
  });

  it('is on by default for ordinary workbooks', () => {
    expect(open('kitchen-sink.xlsx').ensureEngine().off).toBeUndefined();
  });
});

describe('sheets the engine needs', () => {
  it('leaves out data no formula uses, and adds it when one does', async () => {
    const wb = new ExcelJS.Workbook();
    const calc = wb.addWorksheet('Calc');
    const used = wb.addWorksheet('Used');
    const spare = wb.addWorksheet('Spare');
    used.getCell('A1').value = 2;
    spare.getCell('A1').value = 40;
    calc.getCell('A1').value = { formula: 'Used!A1*10', result: 20 };
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 'n.xlsx');
    s.ensureEngine();
    const engine = (s as unknown as { engine: { has(i: number): boolean } }).engine;
    expect([0, 1, 2].map((i) => engine.has(i))).toEqual([true, true, false]);
    const r = s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: '=Spare!A1+2' });
    expect(valueOf(updateAt(r.updates, s, 'Calc', 'A2'))).toBe(42);
    expect(engine.has(2)).toBe(true);
  });
});

describe('sheet text', () => {
  it("isn't kept after opening, and is read again to save", () => {
    const s = open('kitchen-sink.xlsx');
    const pkg = (s as unknown as { pkg: { textCache: Map<string, string> } }).pkg;
    expect([...pkg.textCache.keys()].filter((k) => /^xl\/worksheets\/[^/]+\.xml$/.test(k))).toEqual([]);
    s.apply({ type: 'input', sheet: sheetIndex(s, 'Summary'), row: 30, col: 3, text: 'x' });
    expect(() => WorkbookSession.open(s.save(), 'k.xlsx')).not.toThrow();
  });
});
