import { describe, expect, it } from 'vitest';
import { parseA1 } from '../src/core/refs';
import { WorkbookSession } from '../src/core/session';
import type { EditOp } from '../src/core/types';
import { budgetSchema, forecastSchema, priceListSchema } from '../src/demo/schemas';
import { bindSchema } from '../src/schema/binding';
import { SchemaController } from '../src/schema/controller';
import { generateSchema } from '../src/schema/generate';
import { schemaFromJSON, schemaToJSON } from '../src/schema/json';
import type { WorkbookSchema } from '../src/schema/types';
import { SheetView } from '../src/state/sheetView';
import { sample } from './helpers';

/** Opens a sample the way the app does: worker session plus UI-side sheet views. */
function load(file: string) {
  const session = WorkbookSession.open(sample(file), file);
  const r = session.loadResult();
  const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  for (const u of session.ensureEngine().updates) views[u.sheet].apply(u);
  const apply = (op: EditOp) => {
    const res = session.apply(op);
    for (const u of res.updates) views[u.sheet].apply(u);
    return res;
  };
  return { session, views, styles: r.styles, apply };
}

function control(file: string, schema: WorkbookSchema) {
  const w = load(file);
  const bound = bindSchema(schema, w.views);
  if (!bound.binding) throw new Error(bound.problems.join('\n'));
  const ctl: SchemaController = new SchemaController(bound.binding, {
    views: () => w.views,
    styles: () => w.styles,
    date1904: () => false,
    followUp: async (op) => {
      w.apply(op);
      ctl.validate();
    },
    changed: () => {},
  });
  ctl.validate();
  const edit = (op: EditOp) => {
    const res = w.apply(op);
    ctl.validate();
    return res;
  };
  const issuesAt = (ref: string) => {
    const { row, col } = parseA1(ref)!;
    return ctl.cellIssues(0, row, col).map((i) => i.message);
  };
  return { ...w, ctl, edit, issuesAt };
}

describe('binding', () => {
  it('matches the price list and reports why an old layout does not match', () => {
    const ok = bindSchema(priceListSchema, load('price-list.xlsx').views);
    expect(ok.problems).toEqual([]);
    const old = bindSchema(priceListSchema, load('price-list-old.xlsx').views);
    expect(old.binding).toBeUndefined();
    expect(old.problems).toEqual(['Prices!C1 should be "Category" (field label) but is "Group".']);
    expect(bindSchema(priceListSchema, load('kitchen-sink.xlsx').views).problems[0]).toMatch(/no sheet named "Prices"/);
  });
});

describe('price list (records as rows)', () => {
  it('flags type, enum, uniqueness, precision and required problems', () => {
    const { issuesAt } = control('price-list.xlsx', priceListSchema);
    expect(issuesAt('A8')).toEqual(['Duplicate: A4 has the same SKU.']);
    expect(issuesAt('B11')).toEqual(['Required.']);
    expect(issuesAt('C12')).toEqual(['Must be one of: Hardware, Software, Services.']);
    expect(issuesAt('D13')).toEqual(['At most 2 decimal places.']);
    expect(issuesAt('D2')).toEqual([]);
  });

  it('runs async field checks and reports when they finish', async () => {
    const { ctl, issuesAt } = control('price-list.xlsx', priceListSchema);
    expect(ctl.summary().pending).toBeGreaterThan(0);
    await ctl.settled();
    expect(issuesAt('A14')).toEqual(['The catalogue lists this product as discontinued.']);
    expect(ctl.summary().pending).toBe(0);
  });

  it('locks labels and cells outside the table but leaves the next row open', () => {
    const { ctl } = control('price-list.xlsx', priceListSchema);
    const at = (ref: string) => {
      const { row, col } = parseA1(ref)!;
      return ctl.readOnlyReason(0, row, col);
    };
    expect(at('A1')).toMatch(/Column names/);
    expect(at('B2')).toBeUndefined();
    expect(at('A16')).toBeUndefined();
    expect(at('A17')).toMatch(/Only the “Price list” schema's tables/);
    expect(at('H2')).toMatch(/Only the/);
  });

  it('stores typed input as the field type', () => {
    const { ctl } = control('price-list.xlsx', priceListSchema);
    expect(ctl.prepareInput(0, 1, 0, '1234')).toEqual({ text: "'1234" });
    expect(ctl.prepareInput(0, 1, 2, 'software')).toEqual({ text: "'Software" });
    expect(ctl.prepareInput(0, 1, 6, 'yes')).toEqual({ text: 'TRUE' });
    expect(ctl.prepareInput(0, 1, 3, '=1+1')).toHaveProperty('reject');
  });

  it('inserts and deletes rows, carrying the calculated column', () => {
    const { ctl, edit, views } = control('price-list.xlsx', priceListSchema);
    const table = ctl.binding.sheets[0].tables[0];
    expect(table.recordCount()).toBe(14);
    edit(ctl.insertRecord(table, 2) as EditOp);
    // The new row carries the calculated column, so it stays part of the table.
    expect(table.recordCount()).toBe(15);
    expect(views[0].get(3, 0)?.t ?? 0).toBe(0);
    expect(views[0].get(3, 4)?.f).toBe('D4*1.2');
    expect(views[0].get(4, 0)?.s).toBe('SKU-1003');
    expect(views[0].get(4, 4)?.f).toBe('D5*1.2');
    edit(ctl.deleteRecords(table, 2, 2) as EditOp);
    expect(table.recordCount()).toBe(14);
    expect(views[0].get(3, 0)?.s).toBe('SKU-1003');
    expect(views[0].get(3, 4)?.f).toBe('D4*1.2');
  });

  it('runs business logic after an edit and fills in the new row', async () => {
    const { ctl, apply, views } = control('price-list.xlsx', priceListSchema);
    const before = ctl.valueAt(0, 15, 0);
    apply({ type: 'input', sheet: 0, row: 15, col: 0, text: "'SKU-2001" });
    await ctl.afterEdit([{ sheet: 0, row: 15, col: 0, before }]);
    expect(views[0].get(15, 1)?.s).toBe('Barcode scanner');
    expect(views[0].get(15, 2)?.s).toBe('Hardware');
    expect(views[0].get(15, 4)?.f).toBe('D16*1.2');
  });
});

describe('forecast (records as columns)', () => {
  it('validates down the rows and accepts a new quarter in the next column', () => {
    const { ctl, issuesAt } = control('forecast.xlsx', forecastSchema);
    expect(issuesAt('D1')).toEqual(['Quarters should follow on from each other; a quarter is missing before this one.']);
    expect(issuesAt('E3')).toEqual(['Must be at least 0.']);
    expect(ctl.readOnlyReason(0, 0, 5)).toBeUndefined();
    expect(ctl.readOnlyReason(0, 0, 6)).toMatch(/Only the/);
    expect(ctl.readOnlyReason(0, 1, 0)).toMatch(/Row names/);
  });
});

describe('budget (fixed grid)', () => {
  it('checks labels, single cells and custom rules', async () => {
    const { ctl, issuesAt } = control('budget.xlsx', budgetSchema);
    expect(issuesAt('D5')).toEqual(['Travel is over the monthly cap of 5,000.']);
    expect(issuesAt('B8')).toEqual(['Required.']);
    expect(ctl.readOnlyReason(0, 1, 0)).toMatch(/Row labels/);
    expect(ctl.readOnlyReason(0, 8, 1)).toMatch(/read-only/);
    expect(ctl.insertRecord(ctl.binding.sheets[0].tables[0], 1)).toMatch(/fixed/);
    await ctl.settled();
    expect(ctl.summary().pending).toBe(0);
  });

  it('asks before saving with errors when the schema allows it', async () => {
    const { ctl } = control('budget.xlsx', budgetSchema);
    const r = await ctl.checkBeforeSave();
    expect(r).toEqual({ ok: false, message: '1 cell breaks the “Half-year budget” schema.', canOverride: true });
  });
});

describe('JSON and generation', () => {
  it('exports without hooks and imports again', () => {
    const json = schemaToJSON(priceListSchema);
    expect(json).not.toMatch(/validate|onChange/);
    const back = schemaFromJSON(json);
    expect(bindSchema(back, load('price-list.xlsx').views).problems).toEqual([]);
    expect(() => schemaFromJSON('{"id":"x","name":"x","sheets":[{"name":"S","tables":[{"key":"t","at":"A1","fields":[{"key":"a","type":"colour"}]}]}]}')).toThrow(/unknown type "colour"/);
  });

  it('generates a schema that matches the workbook it came from', () => {
    const w = load('price-list.xlsx');
    const schema = generateSchema(w.views, w.styles, 'price-list.xlsx');
    const fields = schema.sheets[0].tables![0].fields;
    expect(fields.map((f) => [f.key, f.type])).toEqual([
      ['sku', 'text'],
      ['description', 'text'],
      ['category', 'enum'],
      ['unitPrice', 'number'],
      ['priceIncTax', 'number'],
      ['effectiveFrom', 'date'],
      ['active', 'boolean'],
    ]);
    expect(fields[0]).toMatchObject({ pattern: '^SKU-\\d{4}$', required: true });
    expect(fields[2].values).toEqual(['Hardware', 'Hardwear', 'Services', 'Software']);
    expect(fields[4].formula).toBe('require');
    expect(bindSchema(schemaFromJSON(schemaToJSON(schema)), w.views).problems).toEqual([]);
  });

  it('reads a numeric grid with row labels as a fixed table', () => {
    const w = load('budget.xlsx');
    const t = generateSchema(w.views, w.styles, 'budget.xlsx').sheets[0].tables![0];
    expect(t.records?.labels).toEqual(['Salaries', 'Rent', 'Marketing', 'Travel', 'Equipment']);
    expect(t.corner).toBe('Category');
    expect(t.fields.map((f) => f.label)).toEqual(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']);
  });
});
