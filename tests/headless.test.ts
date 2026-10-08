import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { schemaFromJSON, schemaToJSON, HeadlessWorkbook, validateWorkbook } from './headless';
import type { WorkbookSchema } from '../src/schema/types';
import { forecastSchema, priceListSchema } from '../src/demo/schemas';
import { sample } from './helpers';

describe('the editor without a browser', () => {
  it('checks a workbook, async checks included', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { fileName: 'price-list.xlsx', schemas: [forecastSchema, priceListSchema] });
    expect(wb.schema?.id).toBe('demo.price-list');
    const issues = await wb.issues();
    expect(issues).toContainEqual({ severity: 'warning', message: 'The catalogue lists this product as discontinued.', sheet: 'Prices', ref: 'A14' });
    expect(issues.filter((i) => i.severity === 'error').map((i) => i.ref)).toEqual(['A8', 'B11', 'C12', 'D13']);
  }, 10_000);

  it('types values in as a user would, running business logic', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const r = await wb.set('Prices', 'A16', 'SKU-2001');
    expect(r.rejected).toBeUndefined();
    expect(wb.cell('Prices', 'B16')).toBe('Barcode scanner');
    expect(wb.cell('Prices', 'C16')).toBe('Hardware');
    expect(wb.values()!.Prices.tables.prices.at(-1)).toMatchObject({ sku: 'SKU-2001', description: 'Barcode scanner', category: 'Hardware' });
    // As in the editor, a wrong value goes in and is flagged rather than refused.
    await wb.set('Prices', 'D2', 'cheap');
    await wb.set('Prices', 'E2', '1');
    const issues = await wb.issues();
    expect(issues.filter((i) => i.ref === 'D2' || i.ref === 'E2').map((i) => [i.ref, i.message])).toEqual([
      ['D2', 'Expected a number.'],
      ['E2', 'Should be calculated by a formula.'],
    ]);
  }, 10_000);

  it('reports why a workbook matches no schema', async () => {
    const wb = HeadlessWorkbook.open(sample('budget.xlsx'), { schemas: [priceListSchema] });
    expect(wb.schema).toBeUndefined();
    expect(wb.mismatch?.problems[0]).toMatch(/Prices/);
    const r = await validateWorkbook(sample('budget.xlsx'), { schemas: [priceListSchema] });
    expect(r.ok).toBe(false);
  });

  it('saves with the schema version stamped', async () => {
    // The JSON form, without the code rule that the sample's quarters break on purpose.
    const schema = { ...schemaFromJSON(schemaToJSON(forecastSchema)), version: '7' };
    const wb = HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [schema] });
    await wb.set('Forecast', 'E3', '5');
    const again = HeadlessWorkbook.open(wb.save(), { schemas: [{ ...schema, version: '8' }] });
    expect(again.note).toMatch(/version 7 .* version 8/);
    const r = await validateWorkbook(wb.save(), { schemas: [schema] });
    expect(r).toMatchObject({ ok: true, errors: 0, schema: { id: schema.id, version: '7' } });
  });
});

describe('names that are also JavaScript property names', () => {
  it('keeps a sheet, a table and a field called __proto__ as data', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('__proto__');
    ws.addRow(['__proto__', 'constructor']);
    ws.addRow([2, { formula: "'__proto__'!A2*3", result: 6 }]);
    const schema: WorkbookSchema = {
      id: 'odd',
      name: 'Odd names',
      sheets: [
        {
          name: '__proto__',
          tables: [
            {
              key: '__proto__',
              at: 'A1',
              fields: [
                { key: '__proto__', label: '__proto__', type: 'number' },
                { key: 'constructor', label: 'constructor', type: 'number' },
              ],
            },
          ],
        },
      ],
    };
    const book = HeadlessWorkbook.open(new Uint8Array(await wb.xlsx.writeBuffer()), { schemas: [schema], recalc: 'on' });
    expect(book.sheetNames).toEqual(['__proto__']);
    expect(book.cell('__proto__', 'B2')).toBe(6);
    // Written as JSON: in an object literal, __proto__ would set the prototype.
    expect(JSON.stringify(book.values())).toBe('{"__proto__":{"tables":{"__proto__":[{"__proto__":2,"constructor":6}]},"cells":{}}}');
  });
});
