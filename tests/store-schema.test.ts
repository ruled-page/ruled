import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { VT } from '../src/core/types';
import { english } from '../src/i18n';
import { schemaFromJSON, schemaToJSON } from '../src/schema/json';
import type { WorkbookSchema } from '../src/schema/types';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Lines!A1:E4 = Item, Qty, Price, Total, Kind with three records; Kind also has an Excel list. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total', 'Kind']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }, 'Part']);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }, 'Part']);
  lines.addRow(['Saw', 1, 20, { formula: 'B4*C4', result: 20 }, 'Tool']);
  for (let r = 2; r <= 10; r++) {
    lines.getCell(`E${r}`).dataValidation = { type: 'list', formulae: ['"Part,Tool,Other"'], showErrorMessage: true } as ExcelJS.DataValidation;
  }
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const schema = (more: Partial<WorkbookSchema> = {}): WorkbookSchema => ({
  id: 'lines',
  name: { en: 'Order lines', de: 'Bestellzeilen' },
  sheets: [
    {
      name: 'Lines',
      tables: [
        {
          key: 'lines',
          at: 'A1',
          fields: [
            { key: 'item', label: 'Item', type: 'text' },
            { key: 'qty', label: 'Qty', type: 'integer', min: 0 },
            { key: 'price', label: 'Price', type: 'number', formula: 'forbid' },
            { key: 'total', label: 'Total', type: 'number', formula: 'require' },
            { key: 'kind', label: 'Kind', type: 'enum', values: ['Part', 'Tool'] },
          ],
          records: { max: 4 },
        },
      ],
    },
  ],
  ...more,
});

/** Fits no workbook here: it wants a sheet called Prices. */
const prices: WorkbookSchema = { id: 'prices', name: 'Prices', sheets: [{ name: 'Prices', cells: [{ key: 'price', ref: 'B2', type: 'number' }] }, { name: 'Lines' }] };

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
});

async function open(options: StoreOptions = { schemas: [schema()] }, file = bytes): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, file, 'lines.xlsx');
  return store;
}

const cell = (row: number, col: number) => store.view?.get(row, col);
const toasts = () => store.toasts.map((t) => english.text(t.text));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

describe('choosing a schema', () => {
  it('applies the first schema that fits, or says how close the nearest came', async () => {
    await open({ schemas: [prices] });
    expect(store.schema).toBeUndefined();
    expect(store.schemaMismatch).toMatchObject({ name: 'Prices', tried: 1 });
    expect(store.schemaMismatch?.problems.length).toBeGreaterThan(0);

    store.registerSchema(schema());
    expect(store.schema?.schema.id).toBe('lines');
    expect(store.schemaMismatch).toBeUndefined();
    expect(store.schemas.map((s) => s.id)).toEqual(['lines', 'prices']);

    store.registerSchema(schema({ description: 'Newer' }));
    expect(store.schemas.map((s) => s.description)).toEqual(['Newer', undefined]);

    store.removeSchema('lines');
    expect(store.schema).toBeUndefined();
    expect(store.schemaMismatch?.tried).toBe(1);
    store.removeSchema('prices');
    expect(store.schemaMismatch).toBeUndefined();
  });

  it('opens the list of issues when the workbook breaks the schema, except on a phone', async () => {
    const broken = schema({ sheets: [{ ...schema().sheets[0], tables: [{ ...schema().sheets[0].tables![0], records: { max: 2 } }] }] });
    await open({ schemas: [broken] });
    expect(store.schema?.summary().errors).toBeGreaterThan(0);
    expect(store.panel).toBe('issues');

    vi.stubGlobal('matchMedia', () => ({ matches: true }));
    await open({ schemas: [broken] });
    expect(store.panel).toBeNull();
  });

  it('makes cells outside the schema read-only', async () => {
    await open();
    await store.input(0, 7, 'note');
    expect(toasts()).toEqual([expect.stringMatching(/^Only the “Order lines” schema's tables and fields can be edited here/)]);
    expect(cell(0, 7)).toBeFalsy();
  });
});

describe('typing and pasting into schema fields', () => {
  it('stores typed text as the field’s type', async () => {
    await open();
    await store.input(1, 4, 'tool');
    expect(cell(1, 4)).toMatchObject({ t: VT.String, s: 'Tool' });
    await store.input(1, 0, '123');
    expect(cell(1, 0)).toMatchObject({ t: VT.String, s: '123' });
  });

  it('refuses formulas where a field forbids them, typed or pasted', async () => {
    await open();
    await store.input(1, 2, '=1+1');
    await store.paste({ row: 1, col: 1 }, [['5', '=2']], rect(1, 1));
    expect(toasts()).toEqual([
      "Formulas aren't allowed in “Price”. Type a value instead.",
      "Nothing was pasted. Formulas aren't allowed in “Price”. Type a value instead.",
    ]);
    expect([cell(1, 1)?.n, cell(1, 2)?.n]).toEqual([2, 3]);
    await store.paste({ row: 1, col: 3 }, [['=B2*C2*2', 'tool']], rect(1, 3));
    expect([cell(1, 3)?.n, cell(1, 4)?.s]).toEqual([12, 'Tool']);
  });

  it('leaves a field’s values to the schema rather than the workbook’s own list', async () => {
    await open();
    expect(store.listAt(0, 1, 4)).toBeUndefined();
    store.removeSchema('lines');
    expect(store.listAt(0, 1, 4)?.source).toEqual({ kind: 'values', values: ['Part', 'Tool', 'Other'] });
  });
});

describe('business logic', () => {
  const filling = schema({
    onChange: (ctx) => {
      for (const c of ctx.changes) {
        if (c.field === 'qty' && c.record !== undefined) ctx.set({ sheet: 'Lines', table: 'lines', record: c.record, field: 'item' }, `${c.after} pieces`);
      }
    },
  });

  it('fills in cells after an edit, undone together with it', async () => {
    await open({ schemas: [filling] });
    await store.input(2, 1, '8');
    await vi.waitFor(() => expect(cell(2, 0)?.s).toBe('8 pieces'));
    expect(store.undoStack).toHaveLength(1);
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Edit', 'Filled in by business logic']);
    await store.undo();
    expect([cell(2, 0)?.s, cell(2, 1)?.n]).toEqual(['Nuts', 4]);
  });
});

describe('adding and removing records', () => {
  it('inserts an empty record at the active cell, carrying calculated fields', async () => {
    await open();
    store.setSelection({ row: 2, col: 0, rects: [rect(2, 0)] });
    expect(store.recordTarget?.record).toBe(1);
    await store.insertRecord();
    expect(cell(2, 0)?.s ?? '').toBe('');
    expect(cell(3, 0)?.s).toBe('Nuts');
    expect(cell(2, 3)?.f).toBe('B3*C3');
    expect(cell(3, 3)?.f).toBe('B4*C4');

    await store.insertRecord();
    expect(toasts()).toEqual(['This table can have at most 4 rows.']);
  });

  it('deletes the selected records, closing the gap', async () => {
    await open();
    store.setSelection({ row: 1, col: 1, rects: [rect(1, 1, 2, 1)] });
    await store.deleteRecords();
    expect(cell(1, 0)?.s).toBe('Saw');
    expect(cell(1, 3)).toMatchObject({ f: 'B2*C2', n: 20 });
    expect(cell(2, 0)?.s ?? '').toBe('');
  });

  it('does nothing outside a table, or without a schema', async () => {
    await open();
    store.setSelection({ row: 8, col: 8, rects: [rect(8, 8)] });
    expect(store.recordTarget).toBeUndefined();
    await store.insertRecord();
    await store.deleteRecords();
    store.removeSchema('lines');
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0)] });
    await store.insertRecord();
    expect(store.history).toHaveLength(0);
  });

  it('keeps the records a table needs', async () => {
    const atLeastThree = schema({ sheets: [{ ...schema().sheets[0], tables: [{ ...schema().sheets[0].tables![0], records: { min: 3 } }] }] });
    await open({ schemas: [atLeastThree] });
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0)] });
    await store.deleteRecords();
    expect(toasts()).toEqual(['This table needs at least 3 rows.']);
  });
});

describe('a file saved with another version of the schema', () => {
  it('says so in the editor’s language, again when the language changes', async () => {
    let saved: Uint8Array | undefined;
    await open({ schemas: [schema({ version: '1' })], onSave: (f) => void (saved = f.bytes) });
    await store.input(1, 1, '3');
    await store.save();
    store.dispose();

    await open({ schemas: [schema({ version: '2' })] }, saved!);
    expect(english.text(store.schemaNote)).toBe('This file was last saved with version 1 of “Order lines”; it is checked against version 2.');
    store.setLanguage('de');
    expect(store.t.language).toBe('de');
    expect(store.t.text(store.schemaNote)).toMatch(/Bestellzeilen/);
  });
});

describe('schemas from the designer', () => {
  it("keep the code hooks of the host's schema they copy, and a schema that can't be read is passed over", async () => {
    const hooked = schema();
    const validate = vi.fn(() => 'Checked by the app.');
    hooked.sheets[0].tables![0].fields[0].validate = validate;
    await open({ schemas: [hooked] });
    expect(store.isHostSchema('lines')).toBe(true);

    // An edited copy, as JSON brings it back: no code.
    const copy = schemaFromJSON(schemaToJSON(hooked));
    copy.sheets[0].tables![0].fields[1].min = 5;
    store.applyDesignedSchema(copy);
    const fields = store.schema!.schema.sheets[0].tables![0].fields;
    expect(fields[0].validate).toBe(validate);
    expect(fields[1].min).toBe(5);

    // One that throws while binding doesn't stop the next being tried.
    const broken = { ...prices, id: 'broken', sheets: [{ name: 'Lines', tables: [{ key: 't', at: 'A1', fields: null }] }] } as unknown as WorkbookSchema;
    store.applyDesignedSchema(broken);
    expect(store.schema?.schema.id).toBe('lines');
  });
});
