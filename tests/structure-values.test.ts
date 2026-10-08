import ExcelJS from 'exceljs';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';

/**
 * A change of rows or columns rewrites the workbook and opens it again, with
 * each formula's saved value from before the change. Those are no longer
 * Excel's results, so the editor's own take over, except where Excel and the
 * editor were already known to differ.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  [1, 2, 3].forEach((n, i) => (ws.getCell(`C${i + 1}`).value = n));
  ws.getCell('B1').value = { formula: 'SUM(C1:C3)', result: 6 };
  // Saved by "Excel" as 3: the editor makes it 2, and Excel's value is kept.
  ws.getCell('A1').value = { formula: '1+1', result: 3 };
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const at = (session: WorkbookSession, row: number, col: number) =>
  session.ensureEngine().updates.find((u) => u.sheet === 0 && u.row === row && u.col === col);

describe('formula values after a change of rows', () => {
  it('recalculates a total whose range lost a row, rather than keeping the old total', async () => {
    const s = WorkbookSession.open(await workbook(), 'values.xlsx');
    s.ensureEngine();
    const next = s.restructure({ type: 'deleteRows', sheet: 0, at: 1, count: 1 });
    const report = next.ensureEngine();
    expect(at(next, 0, 1)).toMatchObject({ f: 'SUM(C1:C2)', n: 4 });
    expect(report.samples.map((x) => x.ref)).not.toContain('B1');
  });

  it('keeps Excel’s value where Excel and the editor were known to differ', async () => {
    const s = WorkbookSession.open(await workbook(), 'values.xlsx');
    expect(s.ensureEngine().mismatched).toBe(1);
    const next = s.restructure({ type: 'insertRows', sheet: 0, at: 0, count: 1 });
    const report = next.ensureEngine();
    // A1 moved to A2, and still shows Excel's 3 rather than the editor's 2.
    expect(report.samples.map((x) => [x.ref, english.text(x.excel), english.text(x.engine)])).toEqual([['A2', '3', '2']]);
    expect(report.mismatched).toBe(1);
  });
});

describe('precision', () => {
  /** Large amounts with cents, and formulas over them, as Excel saves them. */
  async function amounts(): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Ledger');
    ws.getCell('A1').value = 2234567890.12;
    ws.getCell('B1').value = { formula: 'A1*1', result: 2234567890.12 };
    ws.getCell('C1').value = { formula: 'A1/7', result: 2234567890.12 / 7 };
    ws.getCell('D1').value = { formula: 'ROUND(88765432.1234567,4)', result: 88765432.1235 };
    wb.addWorksheet('Other').getCell('A1').value = 1;
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }

  it('keeps the cents on amounts in the billions, through a change of rows and a save', async () => {
    const s = WorkbookSession.open(await amounts(), 'amounts.xlsx');
    s.ensureEngine();
    // A row inserted on another sheet makes the editor's results the saved ones.
    const next = s.restructure({ type: 'insertRows', sheet: 1, at: 0, count: 1 });
    next.ensureEngine();
    expect(at(next, 0, 1)?.n).toBe(2234567890.12);
    expect(at(next, 0, 2)?.n).toBeCloseTo(2234567890.12 / 7, 5);
    expect(at(next, 0, 3)?.n).toBe(88765432.1235);
    const sheetXml = new TextDecoder().decode(unzipSync(next.save())['xl/worksheets/sheet1.xml']);
    expect(sheetXml).toContain('<v>2234567890.12</v>');
    expect(sheetXml).toContain('<v>88765432.1235</v>');
  });
});
