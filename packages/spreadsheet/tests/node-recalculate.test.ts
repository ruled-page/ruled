import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { recalculate } from '../src/node';

/**
 * Inputs: B1 = 4.5 (a price), B2 = 3 (a quantity). Summary: B1 = price × quantity, saved as
 * 13.5; B2 = B1 + 1, saved as nothing; B3 = B1 × 2, saved wrongly as 999; B4 reads Inputs!B1
 * through CELL("filename"), which isn't calculated here. 'Q1 Sales'!A1 = Summary!B1.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const inputs = wb.addWorksheet('Inputs');
  inputs.getCell('A1').value = 'Price';
  inputs.getCell('B1').value = 4.5;
  inputs.getCell('A2').value = 'Quantity';
  inputs.getCell('B2').value = 3;
  const summary = wb.addWorksheet('Summary');
  summary.getCell('B1').value = { formula: 'Inputs!B1*Inputs!B2', result: 13.5 };
  summary.getCell('B2').value = { formula: 'B1+1' } as ExcelJS.CellFormulaValue;
  summary.getCell('B3').value = { formula: 'B1*2', result: 999 };
  summary.getCell('B4').value = { formula: 'CELL("filename",Inputs!B1)', result: 'book.xlsx' };
  wb.addWorksheet('Q1 Sales').getCell('A1').value = { formula: 'Summary!B1', result: 13.5 };
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** The values another reader finds in a saved workbook: ExcelJS, which calculates nothing itself. */
async function saved(bytes: Uint8Array, sheet: string, refs: string[]): Promise<unknown[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes) as never);
  return refs.map((ref) => {
    const v = wb.getWorksheet(sheet)!.getCell(ref).value;
    return v !== null && typeof v === 'object' && 'formula' in v ? v.result : v;
  });
}

describe('recalculating a saved workbook in Node', () => {
  it('sets the cells given and saves the formulas’ values worked out from them', async () => {
    const result = await recalculate(await workbook(), { set: { 'Inputs!B1': 10 } });
    expect(result.sheets).toEqual(['Inputs', 'Summary', 'Q1 Sales']);
    expect([result.get('Summary!B1'), result.get('Summary!B2'), result.get("'Q1 Sales'!A1")]).toEqual([30, 31, 30]);
    expect(result.getRange('Inputs!A1:B2')).toEqual([
      ['Price', 10],
      ['Quantity', 3],
    ]);
    expect(await saved(result.bytes, 'Summary', ['B1', 'B2'])).toEqual([30, 31]);
    expect(await saved(result.bytes, 'Q1 Sales', ['A1'])).toEqual([30]);
    // B3 kept the value the file saved, which the engine disagreed with, and CELL's file name isn't
    // known here: both keep the saved value, out of date once what they read changed.
    expect(result.get('Summary!B3')).toBe(999);
    expect(result.outOfDate).toEqual(['Summary!B3', 'Summary!B4']);
    expect(result.differs).toEqual(['Summary!B3']);
    const fresh = await recalculate(await workbook(), { set: { 'Inputs!B1': 10 }, ignoreSavedValues: true });
    expect(await saved(fresh.bytes, 'Summary', ['B3'])).toEqual([60]);
    expect(fresh.outOfDate).toEqual(['Summary!B4']);
  });

  it("works out formulas the file has no value for, and keeps Excel's where the engine's differs unless told otherwise", async () => {
    const kept = await recalculate(await workbook());
    expect([kept.get('Summary!B2'), kept.get('Summary!B3')]).toEqual([14.5, 999]);
    expect(kept.differs).toEqual(['Summary!B3']);
    expect(await saved(kept.bytes, 'Summary', ['B2'])).toEqual([14.5]);
    const fresh = await recalculate(await workbook(), { ignoreSavedValues: true });
    expect(fresh.get('Summary!B3')).toBe(27);
    expect(fresh.differs).toEqual([]);
    expect(await saved(fresh.bytes, 'Summary', ['B3'])).toEqual([27]);
  });

  it('stores what it is given: text as text, booleans, dates, formulas, and nothing', async () => {
    const result = await recalculate(await workbook(), {
      set: { 'Inputs!A1': '007', 'Inputs!A2': true, 'Inputs!C1': new Date('2026-10-13T12:30:00Z'), 'Inputs!C2': '=B1*100', 'Inputs!B2': null },
    });
    expect(result.getRange('Inputs!A1:C2')).toEqual([
      ['007', 4.5, 46308.520833333336],
      [true, null, 450],
    ]);
    expect(result.get('Summary!B1')).toBe(0);
  });

  it("refuses a cell the workbook won't let be changed", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Locked');
    ws.getCell('A1').value = 1;
    await ws.protect('', {});
    await expect(recalculate(new Uint8Array(await wb.xlsx.writeBuffer()), { set: { A1: 2 } })).rejects.toThrow("A1 can't be set");
  });

  it('says what is wrong with a reference', async () => {
    const bytes = await workbook();
    await expect(recalculate(bytes, { set: { 'Nowhere!A1': 1 } })).rejects.toThrow('no sheet named "Nowhere"');
    await expect(recalculate(bytes, { set: { 'Inputs!1A': 1 } })).rejects.toThrow("isn't a cell or range reference");
    const result = await recalculate(bytes);
    expect(() => result.get('Inputs!A1:B2')).toThrow('getRange');
    expect(result.get('A1')).toBe('Price');
  });
});
