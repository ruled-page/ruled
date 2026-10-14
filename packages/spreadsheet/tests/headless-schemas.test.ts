import { strToU8, zipSync } from 'fflate';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { forecastSchema, priceListSchema } from '@ruled/demo/schemas';
import { schemaFromJSON, schemaToJSON, HeadlessWorkbook, validateWorkbook, type WorkbookSchema } from './headless';
import { sample } from './helpers';

afterEach(() => vi.useRealTimers());

/** The forecast schema without its code rule (which the sample breaks on purpose), so the sample can be made clean. */
const plainForecast = (extra: Partial<WorkbookSchema> = {}): WorkbookSchema => ({ ...schemaFromJSON(schemaToJSON(forecastSchema)), ...extra });

/** Waits for `work`, moving fake timers on until it is done. */
async function settle<T>(work: Promise<T>): Promise<T> {
  let done = false;
  work.then(
    () => (done = true),
    () => (done = true),
  );
  for (let i = 0; !done && i < 100; i++) await vi.advanceTimersByTimeAsync(100);
  return work;
}

describe('opening a workbook on a server', () => {
  it('takes an ArrayBuffer as well as bytes', () => {
    const bytes = sample('orders.xlsx');
    const wb = HeadlessWorkbook.open(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    expect(wb.fileName).toBe('workbook.xlsx');
    expect(wb.sheetNames).toEqual(['Orders', 'Lists']);
  });

  it('turns away files that aren’t workbooks, saying why in English', () => {
    const cases: [Uint8Array, object, string][] = [
      [strToU8('just some text'), {}, 'This file is not an Excel workbook (.xlsx or .xlsm).'],
      [new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]), {}, 'This file is in the old .xls format. Only .xlsx and .xlsm files are supported.'],
      [zipSync({ 'hello.txt': strToU8('hi') }), {}, 'This file is not a valid Office Open XML package.'],
      [zipSync({ '[Content_Types].xml': strToU8('<Types/>') }), {}, 'The workbook part is missing from this file.'],
      [sample('price-list.xlsx'), { limits: { maxCells: 10 } }, 'This workbook has more than 10 cells, the most this editor can hold.'],
    ];
    for (const [bytes, options, message] of cases) {
      let thrown: unknown;
      try {
        HeadlessWorkbook.open(bytes, options);
      } catch (e) {
        thrown = e;
      }
      expect(thrown).toBeInstanceOf(Error);
      expect((thrown as Error).message).toBe(message);
      // A plain Error: nothing for a server to translate.
      expect((thrown as { text?: unknown }).text).toBeUndefined();
    }
  });

  it('names the closest schema when none fits, or none when no sheet matched', async () => {
    const old = HeadlessWorkbook.open(sample('price-list-old.xlsx'), { schemas: [forecastSchema, priceListSchema] });
    expect(old.schema).toBeUndefined();
    expect(old.mismatch).toEqual({ schema: 'Price list', problems: ['Prices!C1 should be "Category" (field label) but is "Group".'] });
    expect(HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [forecastSchema] }).mismatch).toEqual({ problems: ['The workbook has no sheet named "Forecast".'] });
  });

  it('works without a schema: no issues, no values, nothing to stop a save', async () => {
    const wb = HeadlessWorkbook.open(sample('orders.xlsx'));
    expect(wb.schema).toBeUndefined();
    expect(wb.mismatch).toBeUndefined();
    expect(await wb.issues()).toEqual([]);
    expect(wb.values()).toBeUndefined();
    expect(await wb.checkBeforeSave()).toEqual({ ok: true });
  });

  it('says in English what the file holds that the editor only keeps', () => {
    const findings = HeadlessWorkbook.open(sample('kitchen-sink.xlsx')).findings;
    expect(findings).toContainEqual(
      expect.objectContaining({ title: 'Chart "Revenue by region"', detail: 'Column chart. Kept in the file and redrawn by Excel from the edited data when the file is opened.' }),
    );
    for (const f of findings) expect([typeof f.title, typeof f.detail]).toEqual(['string', 'string']);
  });
});

describe('typing into a workbook on a server', () => {
  it('turns down input the schema refuses, and changes nothing', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const r = await wb.set('Prices', 'D2', '=1+1');
    expect(r).toMatchObject({ rejected: "Formulas aren't allowed in “Unit price”. Type a value instead.", updates: [], skipped: 0 });
    expect(wb.changes()).toEqual([]);
    expect(wb.cell('Prices', 'D2')).toBe(4.5);
  });

  it('refuses cells the schema locks, saying why', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const r = await wb.set('Prices', 'A1', 'Code');
    const why = 'Column names in this table are set by the “Price list” schema.';
    expect(r).toMatchObject({ rejected: why, skipped: 1, skippedReasons: [why], updates: [] });
    expect(wb.editedCount).toBe(0);
  });

  it('finds sheets whatever their case, and throws for a sheet or cell that isn’t there', async () => {
    const wb = HeadlessWorkbook.open(sample('orders.xlsx'));
    await wb.set('orders', 'Z1', 'note');
    expect(wb.cell('Orders', 'Z1')).toBe('note');
    await expect(wb.set('Nope', 'A1', 'x')).rejects.toThrow('There is no sheet called "Nope".');
    await expect(wb.set('Orders', 'A1:B2', 'x')).rejects.toThrow('"A1:B2" isn\'t a cell reference.');
  });

  it('refuses references off the sheet rather than writing somewhere else', async () => {
    const wb = HeadlessWorkbook.open(sample('orders.xlsx'));
    // XFE1 would be column 16,384: one past the last, which is A2 of the next row.
    for (const ref of ['XFE1', 'ZZZ1', 'A0', 'A1048577']) {
      await expect(wb.set('Orders', ref, 'x'), ref).rejects.toThrow(`"${ref}" isn't a cell reference.`);
    }
    expect(wb.changes()).toEqual([]);
    expect(wb.cell('Orders', 'A2')).toBe(5001);
    // The last cell of a sheet is still a cell.
    await wb.set('Orders', 'XFD1048576', 'corner');
    expect(wb.changes()).toEqual([{ sheet: 'Orders', ref: 'XFD1048576', before: '', after: 'corner' }]);
  });

  it('reports a formula it can’t read, in English', async () => {
    const wb = HeadlessWorkbook.open(sample('orders.xlsx'));
    expect((await wb.set('Orders', 'Z1', '=SUM(')).rejected).toBe("There's a problem with this formula: Invalid syntax");
    expect(wb.editedCount).toBe(0);
  });
});

describe('other edits on a server', () => {
  it('applies a paste, with its warnings in English, and lists the changes', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const notes = wb.sheetNames.indexOf('Notes');
    const r = await wb.edit({ type: 'pasteText', sheet: notes, row: 0, col: 5, values: [['=SUM(', '2']] });
    expect(r.warnings).toEqual(["1 pasted value looked like a formula but wasn't valid, so it was pasted as text."]);
    expect(wb.changes()).toEqual([
      { sheet: 'Notes', ref: 'F1', before: '', after: '=SUM(' },
      { sheet: 'Notes', ref: 'G1', before: '', after: '2' },
    ]);
    expect(wb.editedCount).toBe(2);
  });

  it('runs the schema’s business logic after a paste as after typing', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    await settle(wb.edit({ type: 'pasteText', sheet: 0, row: 15, col: 0, values: [['SKU-2002']] }));
    expect(wb.range('Prices', 'A16:C16')).toEqual([['SKU-2002', 'Label printer', 'Hardware']]);
  });

  it('changes nothing for an edit that touches nothing', async () => {
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const r = await wb.edit({ type: 'clear', sheet: 0, rects: [{ r1: 20, c1: 0, r2: 21, c2: 3 }] });
    expect(r.updates).toEqual([]);
    expect(wb.changes()).toEqual([]);
  });
});

describe('the last word before saving', () => {
  it('counts the schema’s errors, in English', async () => {
    // The price list's catalogue check takes 500 ms, and saving waits for it.
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
    const wb = HeadlessWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    expect(await settle(wb.checkBeforeSave())).toEqual({ ok: false, message: '4 cells break the “Price list” schema.', canOverride: false });
  });

  it('gives the schema’s own reason in English, or why its hook failed', async () => {
    const open = (beforeSave: WorkbookSchema['beforeSave'], extra: Partial<WorkbookSchema> = {}) =>
      HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [plainForecast({ beforeSave, ...extra })] });
    const fixed = async (wb: HeadlessWorkbook) => (await wb.set('Forecast', 'E3', '5'), wb);

    const reason = await fixed(open(() => ({ en: 'Finance hasn’t signed this off yet.', de: 'Die Finanzabteilung hat noch nicht zugestimmt.' })));
    expect(await reason.checkBeforeSave()).toEqual({ ok: false, message: 'Finance hasn’t signed this off yet.', canOverride: false });

    const failing = await fixed(
      open(() => {
        throw new Error('ledger offline');
      }),
    );
    expect(await failing.checkBeforeSave()).toEqual({ ok: false, message: 'The save check failed: ledger offline', canOverride: false });

    const slow = await fixed(open(() => new Promise(() => {}), { hookTimeoutMs: 20 }));
    expect(await slow.checkBeforeSave()).toEqual({ ok: false, message: "The checks before saving didn't finish within 20 ms. Try saving again.", canOverride: false });

    const fine = await fixed(open(() => true));
    expect(await fine.checkBeforeSave()).toEqual({ ok: true });
  });
});

describe('saving on a server', () => {
  it('stamps the schema version unless told not to', async () => {
    const wb = HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [plainForecast({ version: '7' })] });
    // An unchanged workbook is saved as it was opened, so change something.
    await wb.set('Forecast', 'E3', '5');
    const next = [plainForecast({ version: '8' })];
    expect(HeadlessWorkbook.open(wb.save(), { schemas: next }).note).toMatch(/version 7 .* version 8/);
    expect(HeadlessWorkbook.open(wb.save({ stamp: false }), { schemas: next }).note).toBeUndefined();
  });
});

describe('validateWorkbook', () => {
  it('accepts a clean workbook, naming the schema in English', async () => {
    const wb = HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [plainForecast({ version: '3' })] });
    await wb.set('Forecast', 'E3', '5');
    const r = await validateWorkbook(wb.save(), { schemas: [plainForecast({ version: '3' })] });
    expect(r).toMatchObject({ ok: true, schema: { id: 'demo.forecast', name: 'Quarterly forecast', version: '3' }, errors: 0, warnings: 0, issues: [] });
    expect(r.values!.Forecast.tables.quarters[3]).toMatchObject({ period: '2027-Q1', cost: 5 });
    expect('note' in r).toBe(false);
    expect('mismatch' in r).toBe(false);
  });

  it('counts warnings apart from errors, and warnings alone don’t fail it', async () => {
    const schema = plainForecast({
      validate: () => [
        { ref: 'Forecast!B2', severity: 'warning', message: 'Revenue looks high.' },
        { severity: 'warning', message: { en: 'Nobody has signed this off.', de: 'Niemand hat das freigegeben.' } },
      ],
    });
    const wb = HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [schema] });
    await wb.set('Forecast', 'E3', '5');
    const r = await validateWorkbook(wb.save(), { schemas: [schema] });
    expect(r).toMatchObject({ ok: true, errors: 0, warnings: 2 });
    // An issue about the whole workbook has no cell.
    expect(r.issues).toEqual([
      { severity: 'warning', message: 'Nobody has signed this off.', sheet: '' },
      { severity: 'warning', message: 'Revenue looks high.', sheet: 'Forecast', ref: 'B2' },
    ]);
  });

  it('fails a workbook that matches no schema, with why', async () => {
    const r = await validateWorkbook(sample('price-list-old.xlsx'), { schemas: [priceListSchema] });
    expect(r).toEqual({ ok: false, mismatch: { schema: 'Price list', problems: ['Prices!C1 should be "Category" (field label) but is "Group".'] }, errors: 0, warnings: 0, issues: [], values: undefined });
  });

  it('passes on the note about a file saved under another version', async () => {
    const wb = HeadlessWorkbook.open(sample('forecast.xlsx'), { schemas: [plainForecast({ version: '1' })] });
    await wb.set('Forecast', 'E3', '5');
    const r = await validateWorkbook(wb.save(), { schemas: [plainForecast({ version: '2' })] });
    expect(r).toMatchObject({ ok: true, schema: { version: '2' } });
    expect(r.note).toMatch(/version 1 .* version 2/);
  });

  it('accepts any workbook when no schema is given', async () => {
    expect(await validateWorkbook(sample('orders.xlsx'))).toEqual({ ok: true, errors: 0, warnings: 0, issues: [], values: undefined });
  });
});
