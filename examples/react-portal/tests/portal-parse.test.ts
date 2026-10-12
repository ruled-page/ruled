import { afterEach, describe, expect, it, vi } from 'vitest';
import { HeadlessWorkbook } from '@sitepen/spreadsheet/testing/headless.ts';
import { CHANGE_THRESHOLD, priceChange, readPriceLines, summarise, type PriceLine } from '../src/parse';
import { SHEET, supplierPriceList, TABLE } from '../src/schema';
import { approvedPrices, type ApprovedPrice } from '../src/services';
import { sample } from '@sitepen/spreadsheet/testing/helpers.ts';

afterEach(() => vi.unstubAllGlobals());

const usd = (price: number): ApprovedPrice => ({ price, currency: 'USD', approvedAt: '2026-07-01' });

/** A save as the editor hands it over, with the price list's rows. */
const saved = (rows: Record<string, unknown>[]) => ({ values: { [SHEET]: { tables: { [TABLE]: rows }, cells: {} } } }) as never;

const line = (over: Partial<PriceLine>): PriceLine => ({
  sku: 'SKU-1001',
  description: '',
  category: 'Hardware',
  price: 10,
  currency: 'USD',
  minimumOrder: 1,
  validFrom: '2026-11-01',
  orderValue: 10,
  ...over,
});

describe('reading a saved price list', () => {
  it('turns each row into a price line, with the approved price it replaces', () => {
    const lines = readPriceLines(
      saved([
        { sku: 'SKU-1002', description: 'Steel bracket, large', category: 'Hardware', price: 9.4, currency: 'USD', minimumOrder: 50, validFrom: '2026-11-01', orderValue: 470 },
        { sku: 'SKU-1099', description: 'Rack shelf', category: 'Hardware', price: 55, currency: 'USD', minimumOrder: 5, validFrom: '2026-11-01', orderValue: 275 },
      ]),
      { 'SKU-1002': usd(7.25) },
    );
    expect(lines).toEqual([
      { sku: 'SKU-1002', description: 'Steel bracket, large', category: 'Hardware', price: 9.4, currency: 'USD', minimumOrder: 50, validFrom: '2026-11-01', orderValue: 470, approved: usd(7.25) },
      { sku: 'SKU-1099', description: 'Rack shelf', category: 'Hardware', price: 55, currency: 'USD', minimumOrder: 5, validFrom: '2026-11-01', orderValue: 275 },
    ]);
    // A new product has no approved price at all, not an undefined one.
    expect('approved' in lines[1]).toBe(false);
  });

  it('upper-cases product codes, so a lower-case SKU still finds its approved price', () => {
    const [l] = readPriceLines(saved([{ sku: 'sku-1002', price: 9.4, currency: 'USD' }]), { 'SKU-1002': usd(7.25) });
    expect(l.sku).toBe('SKU-1002');
    expect(l.approved).toEqual(usd(7.25));
  });

  it('fills fields that are missing with empty text and zero', () => {
    expect(readPriceLines(saved([{}]), {})).toEqual([
      { sku: '', description: '', category: '', price: 0, currency: '', minimumOrder: 0, validFrom: '', orderValue: 0 },
    ]);
    expect(readPriceLines(saved([{ sku: null, description: null, price: null }]), {})[0]).toMatchObject({ sku: '', description: '', price: 0 });
  });

  it('reads numbers given as text as zero, so totals never become NaN', () => {
    const [l] = readPriceLines(saved([{ price: '12.50', minimumOrder: '10', orderValue: Number.NaN }]), {});
    expect([l.price, l.minimumOrder, l.orderValue]).toEqual([0, 0, 0]);
    expect(summarise([l]).orderValue).toBe(0);
  });

  it('reads text fields given as numbers as text, and keeps only the date of a date-time', () => {
    const [l] = readPriceLines(saved([{ sku: 1001, description: 42, validFrom: '2026-11-01T00:00:00.000Z' }]), {});
    expect([l.sku, l.description, l.validFrom]).toEqual(['1001', '42', '2026-11-01']);
  });

  it('finds no lines in a save without the price list', () => {
    expect(readPriceLines({} as never, {})).toEqual([]);
    expect(readPriceLines({ values: { Other: { tables: {}, cells: {} } } } as never, {})).toEqual([]);
    expect(readPriceLines({ values: { [SHEET]: { tables: {}, cells: {} } } } as never, {})).toEqual([]);
  });
});

describe('the change from the approved price', () => {
  it('is a fraction of the approved price, up or down', () => {
    expect(priceChange(line({ price: 9.425, approved: usd(7.25) }))).toBeCloseTo(0.3, 10);
    expect(priceChange(line({ price: 180, approved: usd(240) }))).toBeCloseTo(-0.25, 10);
    expect(priceChange(line({ price: 85, approved: usd(85) }))).toBe(0);
  });

  it('is unknown without an approved price in the same currency', () => {
    expect(priceChange(line({}))).toBeUndefined();
    expect(priceChange(line({ currency: 'EUR', approved: usd(10) }))).toBeUndefined();
    expect(priceChange(line({ approved: usd(0) }))).toBeUndefined();
  });

  it('counts a move of exactly 15% as within the threshold, up or down', () => {
    // 85 / 100 − 1 is −0.15000000000000002 in floating point.
    const down = line({ price: 85, approved: usd(100) });
    const up = line({ price: 115, approved: usd(100) });
    expect(Math.abs(priceChange(down)!)).toBeLessThanOrEqual(CHANGE_THRESHOLD);
    expect(Math.abs(priceChange(up)!)).toBeLessThanOrEqual(CHANGE_THRESHOLD);
    expect(summarise([down, up, line({ price: 8.5, approved: usd(10) })]).bigChanges).toBe(0);
    expect(summarise([line({ price: 84.99, approved: usd(100) }), line({ price: 115.01, approved: usd(100) })]).bigChanges).toBe(2);
  });
});

describe('summarising a submission', () => {
  it('says nothing much about an empty list', () => {
    const s = summarise([]);
    expect(s).toStrictEqual({ products: 0, currency: '', orderValue: 0, byCategory: [], bigChanges: 0, newProducts: 0 });
  });

  it('totals by category, largest first, and averages the change over products that have one', () => {
    const s = summarise([
      line({ sku: 'SKU-1001', category: 'Hardware', orderValue: 100, price: 13, approved: usd(10) }),
      line({ sku: 'SKU-1002', category: 'Services', orderValue: 500, price: 9, approved: usd(10) }),
      line({ sku: 'SKU-1003', category: 'Hardware', orderValue: 200, price: 10, approved: usd(10) }),
      // In another currency: not compared, but not new either.
      line({ sku: 'SKU-1004', category: 'Software', orderValue: 50, price: 99, currency: 'EUR', approved: usd(10) }),
      line({ sku: 'SKU-1099', category: 'Software', orderValue: 25 }),
    ]);
    expect(s.products).toBe(5);
    expect(s.currency).toBe('USD');
    expect(s.orderValue).toBe(875);
    expect(s.byCategory).toEqual([
      { category: 'Services', products: 1, orderValue: 500 },
      { category: 'Hardware', products: 2, orderValue: 300 },
      { category: 'Software', products: 2, orderValue: 75 },
    ]);
    // (+30% − 10% + 0%) / 3
    expect(s.averageChange).toBeCloseTo(0.2 / 3, 10);
    expect(s.bigChanges).toBe(1);
    expect(s.newProducts).toBe(1);
  });

  it('reads what the portal’s schema produces from the supplier sample', () => {
    // No browser storage here: the approved prices are the seeded ones.
    vi.stubGlobal('localStorage', undefined);
    const wb = HeadlessWorkbook.open(sample('supplier-price-list.xlsx'), { schemas: [supplierPriceList] });
    const lines = readPriceLines({ values: wb.values() }, approvedPrices());
    expect(lines.map((l) => l.sku)).toEqual(['SKU-1001', 'SKU-1002', 'SKU-1003', 'SKU-1004', 'SKU-1005', 'SKU-1006', 'SKU-1007', 'SKU-1008', 'SKU-1099', 'SKU-1010', 'SKU-1011']);
    expect(lines[1]).toMatchObject({ description: 'Steel bracket, large', price: 9.4, minimumOrder: 50, validFrom: '2026-11-01', orderValue: 470, approved: { price: 7.25 } });
    const s = summarise(lines);
    expect(s).toMatchObject({ products: 11, currency: 'USD', bigChanges: 1, newProducts: 1 });
    expect(s.byCategory.map((c) => c.category)).toEqual(['Services', 'Hardware', 'Software', 'Hardwear']);
  });
});

