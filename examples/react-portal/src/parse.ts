/**
 * What the portal does with a save: the editor hands over the workbook's
 * values in the shape of the schema, as plain JSON, and the portal turns
 * them into its own records. This is the only place that knows the
 * schema's field names.
 */
import type { SavedWorkbook } from '@ruled/sheet';
import { SHEET, TABLE } from './schema';
import type { ApprovedPrice } from './services';

export interface PriceLine {
  sku: string;
  description: string;
  category: string;
  price: number;
  currency: string;
  minimumOrder: number;
  /** ISO date, as the editor hands dates over. */
  validFrom: string;
  /** Price × minimum order, as the workbook's formula worked it out. */
  orderValue: number;
  /** The last approved price when this was submitted, to show the change. */
  approved?: ApprovedPrice;
}

const text = (v: unknown) => (typeof v === 'string' ? v : v == null ? '' : String(v));
const num = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);

/** The rows of a saved price list, with the approved prices they replace. */
export function readPriceLines(saved: Pick<SavedWorkbook, 'values'>, approved: Record<string, ApprovedPrice>): PriceLine[] {
  const rows = saved.values?.[SHEET]?.tables[TABLE] ?? [];
  return rows.map((r) => {
    const sku = text(r.sku).toUpperCase();
    return {
      sku,
      description: text(r.description),
      category: text(r.category),
      price: num(r.price),
      currency: text(r.currency),
      minimumOrder: num(r.minimumOrder),
      validFrom: text(r.validFrom).slice(0, 10),
      orderValue: num(r.orderValue),
      ...(approved[sku] ? { approved: approved[sku] } : {}),
    };
  });
}

/** The change from the approved price, as a fraction (0.3 is 30% up); undefined when there's nothing to compare. */
export function priceChange(line: PriceLine): number | undefined {
  const a = line.approved;
  if (!a || a.currency !== line.currency || !a.price) return undefined;
  // Rounded, as the ratio of two prices isn't exact: 85 / 100 − 1 is −0.15000000000000002, past the threshold.
  return Math.round((line.price / a.price - 1) * 1e9) / 1e9;
}

export interface Summary {
  products: number;
  currency: string;
  orderValue: number;
  byCategory: { category: string; products: number; orderValue: number }[];
  /** Average change from approved prices, over the products that have one. */
  averageChange?: number;
  /** Products whose price moved more than the portal's 15% threshold. */
  bigChanges: number;
  newProducts: number;
}

export const CHANGE_THRESHOLD = 0.15;

export function summarise(lines: PriceLine[]): Summary {
  const categories = new Map<string, { products: number; orderValue: number }>();
  let changeSum = 0;
  let compared = 0;
  let bigChanges = 0;
  for (const l of lines) {
    const c = categories.get(l.category) ?? { products: 0, orderValue: 0 };
    c.products++;
    c.orderValue += l.orderValue;
    categories.set(l.category, c);
    const change = priceChange(l);
    if (change !== undefined) {
      changeSum += change;
      compared++;
      if (Math.abs(change) > CHANGE_THRESHOLD) bigChanges++;
    }
  }
  return {
    products: lines.length,
    currency: lines[0]?.currency ?? '',
    orderValue: lines.reduce((n, l) => n + l.orderValue, 0),
    byCategory: [...categories].map(([category, c]) => ({ category, ...c })).sort((a, b) => b.orderValue - a.orderValue),
    ...(compared ? { averageChange: changeSum / compared } : {}),
    bigChanges,
    newProducts: lines.filter((l) => !l.approved).length,
  };
}
