import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createI18n, translator } from '@ruled/sheet/internal/i18n/index.ts';
import { HeadlessWorkbook } from '@ruled/sheet/testing/headless.ts';
import { supplierPriceList } from '../src/schema';
import type { FieldValue, IssueInput, RecordData, TableData, WorkbookData } from '@ruled/sheet';
import { sample } from '@ruled/sheet/testing/helpers.ts';

// The portal's services answer after 250–400 ms; fake timers play that out at once.
beforeEach(() => {
  vi.stubGlobal('localStorage', undefined);
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

/** Waits for `work`, moving the clock on until it is done. */
async function settle<T>(work: Promise<T>): Promise<T> {
  let done = false;
  work.then(
    () => (done = true),
    () => (done = true),
  );
  for (let i = 0; !done && i < 100; i++) await vi.advanceTimersByTimeAsync(100);
  return work;
}

const open = (file = 'supplier-price-list.xlsx') => HeadlessWorkbook.open(sample(file), { fileName: file, schemas: [supplierPriceList] });
const SHEET = 'Price list';

const [priceRule, currencyRule] = supplierPriceList.sheets[0].tables![0].rules!;

/** The price list's table as a rule sees it. */
function table(rows: Record<string, FieldValue>[]): TableData {
  const records: RecordData[] = rows.map((values, index) => ({ index, values, ref: () => '' }));
  return { key: 'lines', sheet: SHEET, records, fields: supplierPriceList.sheets[0].tables![0].fields };
}
const ctx = { language: 'en', signal: new AbortController().signal, sheet: () => undefined } as WorkbookData;
const runRule = async (rule: typeof priceRule, rows: Record<string, FieldValue>[]) => ((await settle(Promise.resolve(rule(table(rows), ctx)))) ?? []) as IssueInput[];

describe('the supplier price list', () => {
  it('flags the sample’s five errors and one warning, in English', async () => {
    const wb = open();
    expect(wb.schema?.id).toBe('portal.supplier-price-list');
    expect(await settle(wb.issues())).toEqual([
      { severity: 'error', message: 'Must be at least 1.', sheet: SHEET, ref: 'F7' },
      { severity: 'error', message: 'Must be one of: Hardware, Software, Services.', sheet: SHEET, ref: 'C8' },
      { severity: 'error', message: 'Use one currency for the whole list (USD, as in the first row).', sheet: SHEET, ref: 'E9' },
      { severity: 'error', message: 'SKU-1099 isn’t in our catalogue. Check the code, or ask us to add the product first.', sheet: SHEET, ref: 'A10' },
      { severity: 'error', message: 'At most 2 decimal places.', sheet: SHEET, ref: 'D11' },
      { severity: 'warning', message: '+30% on the last approved price of $7.25. We’ll check this before approving.', sheet: SHEET, ref: 'D3' },
    ]);
    expect(await settle(wb.checkBeforeSave())).toEqual({ ok: false, message: '5 cells break the “Supplier price list” schema.', canOverride: false });
  });

  it('clears every issue once the cells are fixed', async () => {
    const wb = open();
    for (const [ref, text] of [
      ['F7', '1'],
      ['C8', 'hardware'],
      ['E9', 'USD'],
      ['A10', 'SKU-1012'],
      ['D11', '1250.13'],
      ['D3', '7.60'],
    ]) {
      expect((await settle(wb.set(SHEET, ref, text))).rejected).toBeUndefined();
    }
    expect(await settle(wb.issues())).toEqual([]);
    expect(await settle(wb.checkBeforeSave())).toEqual({ ok: true });
    expect(wb.cell(SHEET, 'C8')).toBe('Hardware');
    // The order value is the workbook's own formula, recalculated.
    expect(wb.cell(SHEET, 'H7')).toBe(85);
  });

  it('turns down an empty template until a row is filled in', async () => {
    const wb = open('supplier-template.xlsx');
    expect(await settle(wb.issues())).toEqual([{ severity: 'error', message: 'This table needs at least 1 row.', sheet: SHEET, ref: 'A2' }]);
  });

  it('keeps the instructions sheet and the calculated column read-only', async () => {
    const wb = open();
    expect((await wb.set('How to fill this in', 'A1', 'x')).rejected).toBe('Only the sheets described by the “Supplier price list” schema can be edited.');
    expect((await wb.set(SHEET, 'H2', '5')).rejected).toBe('“Order value” is read-only.');
    expect(wb.editedCount).toBe(0);
  });
});

describe('the catalogue check', () => {
  it('passes a product in the catalogue and names one that isn’t', async () => {
    const wb = open();
    await settle(wb.set(SHEET, 'A10', 'SKU-1098'));
    const at = async (ref: string) => (await settle(wb.issues())).filter((i) => i.ref === ref).map((i) => i.message);
    expect(await at('A10')).toEqual(['SKU-1098 isn’t in our catalogue. Check the code, or ask us to add the product first.']);
    await settle(wb.set(SHEET, 'A10', 'SKU-1009'));
    expect(await at('A10')).toEqual([]);
  });
});

describe('the price-change check', () => {
  it('warns in English and German when a price moves more than 15%', async () => {
    const issues = await runRule(priceRule, [
      { sku: 'SKU-1002', price: 9.4, currency: 'USD' },
      { sku: 'SKU-1004', price: 180, currency: 'USD' },
    ]);
    expect(issues).toEqual([
      {
        record: 0,
        field: 'price',
        severity: 'warning',
        message: {
          en: '+30% on the last approved price of $7.25. We’ll check this before approving.',
          de: '+30 % gegenüber dem zuletzt genehmigten Preis von 7,25 $. Wir prüfen das vor der Genehmigung.',
        },
      },
      {
        record: 1,
        field: 'price',
        severity: 'warning',
        message: {
          en: '-25% on the last approved price of $240.00. We’ll check this before approving.',
          de: '-25 % gegenüber dem zuletzt genehmigten Preis von 240,00 $. Wir prüfen das vor der Genehmigung.',
        },
      },
    ]);
    // The editor shows the one for its language.
    expect(translator(createI18n('de-AT')).text(issues[0].message)).toMatch(/^\+30 % gegenüber/);
  });

  it('leaves alone new products, other currencies, small moves and rows it can’t read', async () => {
    expect(
      await runRule(priceRule, [
        { sku: 'SKU-1099', price: 99, currency: 'USD' },
        { sku: 'SKU-1008', price: 900, currency: 'EUR' },
        { sku: 'SKU-1001', price: 5.17, currency: 'USD' },
        // Exactly 15% down and up from 85.00.
        { sku: 'SKU-1006', price: 72.25, currency: 'USD' },
        { sku: 'SKU-1006', price: 97.75, currency: 'USD' },
        { sku: null, price: 1, currency: 'USD' },
        { sku: 'SKU-1001', price: 'cheap', currency: 'USD' },
        { sku: 'SKU-1001', price: 100, currency: null },
      ]),
    ).toEqual([]);
  });

  it('checks again when a product code changes', async () => {
    const wb = open();
    const warnings = async () => (await settle(wb.issues())).filter((i) => i.severity === 'warning').map((i) => i.ref);
    expect(await warnings()).toEqual(['D3']);
    // SKU-1013 has no approved price, so there is nothing to compare with.
    await settle(wb.set(SHEET, 'A3', 'SKU-1013'));
    expect(await warnings()).toEqual([]);
  });
});

describe('the one-currency rule', () => {
  it('follows the first row with a currency', async () => {
    expect(await runRule(currencyRule, [{ currency: null }, { currency: 'EUR' }, { currency: 'USD' }, { currency: 'EUR' }])).toEqual([
      { record: 2, field: 'currency', message: { en: 'Use one currency for the whole list (EUR, as in the first row).', de: 'Verwenden Sie eine Währung für die ganze Liste (EUR, wie in der ersten Zeile).' } },
    ]);
    expect(await runRule(currencyRule, [{ currency: null }, { currency: null }])).toEqual([]);
    expect(await runRule(currencyRule, [])).toEqual([]);
  });

  it('moves to every other row when the first row’s currency changes', async () => {
    const wb = open();
    await settle(wb.set(SHEET, 'E2', 'EUR'));
    const refs = (await settle(wb.issues())).filter((i) => i.message.startsWith('Use one currency')).map((i) => i.ref);
    expect(refs).toEqual(['E3', 'E4', 'E5', 'E6', 'E7', 'E8', 'E10', 'E11', 'E12']);
  });
});

describe('filling in a product from the catalogue', () => {
  it('fills in the description and category of a known product code typed into a new row', async () => {
    const wb = open();
    await settle(wb.set(SHEET, 'A13', 'SKU-1012'));
    expect(wb.cell(SHEET, 'B13')).toBe('Barcode scanner');
    expect(wb.cell(SHEET, 'C13')).toBe('Hardware');
    expect(wb.values()![SHEET].tables.lines.at(-1)).toMatchObject({ sku: 'SKU-1012', description: 'Barcode scanner', category: 'Hardware' });
    // With the order value's formula carried into the new row.
    expect(wb.changes().map((c) => [c.ref, c.after])).toEqual([
      ['A13', 'SKU-1012'],
      ['B13', 'Barcode scanner'],
      ['C13', 'Hardware'],
      ['H13', '=D13*F13'],
    ]);
  });

  it('leaves a description the supplier wrote', async () => {
    const wb = open();
    await settle(wb.set(SHEET, 'A2', 'SKU-1013'));
    expect(wb.cell(SHEET, 'B2')).toBe('Steel bracket, small');
    expect(wb.changes().map((c) => c.ref)).toEqual(['A2']);
  });

  it('fills in nothing for a code that isn’t in the catalogue', async () => {
    const wb = open();
    await settle(wb.set(SHEET, 'A13', 'SKU-9999'));
    expect(wb.cell(SHEET, 'B13')).toBeNull();
    const at13 = (await settle(wb.issues())).filter((i) => i.ref === 'A13').map((i) => i.message);
    expect(at13).toEqual(['SKU-9999 isn’t in our catalogue. Check the code, or ask us to add the product first.']);
  });
});
