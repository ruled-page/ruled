/**
 * The portal's rules for a supplier price list. Three kinds of check meet
 * here: the types and limits of each column, the workbook's own Excel
 * validation (which the editor enforces without being told), and the
 * portal's business checks, which call its services: every product must be
 * in the catalogue, and a price that moves more than 15% from the last
 * approved one is flagged. Messages are in English and German.
 */
import { defineSchema } from '../editor';
import { lookupApprovedPrice, lookupProduct } from './services';

export const SHEET = 'Price list';
export const TABLE = 'lines';

const pct = (fraction: number, language: string) =>
  new Intl.NumberFormat(language, { style: 'percent', maximumFractionDigits: 0, signDisplay: 'exceptZero' }).format(fraction);
const money = (n: number, currency: string, language: string) =>
  new Intl.NumberFormat(language, { style: 'currency', currency }).format(n);

export const supplierPriceList = defineSchema({
  id: 'portal.supplier-price-list',
  name: { en: 'Supplier price list', de: 'Lieferantenpreisliste' },
  description: {
    en: 'One product per row. Products must be in our catalogue; prices that move more than 15% from the last approved price are flagged for a second look.',
    de: 'Ein Produkt pro Zeile. Produkte müssen in unserem Katalog stehen; Preise, die mehr als 15 % vom zuletzt genehmigten Preis abweichen, werden zur Prüfung markiert.',
  },
  version: '2026.2',
  otherSheets: 'readOnly',
  sheets: [
    {
      name: SHEET,
      tables: [
        {
          key: TABLE,
          at: 'A1',
          orientation: 'rows',
          records: { min: 1, max: 500 },
          fields: [
            {
              key: 'sku',
              label: 'SKU',
              type: 'text',
              required: true,
              unique: true,
              pattern: '^SKU-\\d{4}$',
              patternHint: { en: 'like SKU-1001', de: 'wie SKU-1001' },
              // Business check against the portal's catalogue API.
              validate: async (value, { signal }) => {
                if (await lookupProduct(String(value), signal)) return undefined;
                return { en: `${value} isn’t in our catalogue. Check the code, or ask us to add the product first.`, de: `${value} steht nicht in unserem Katalog. Prüfen Sie den Code oder bitten Sie uns, das Produkt zuerst aufzunehmen.` };
              },
            },
            { key: 'description', label: 'Description', type: 'text', required: true },
            { key: 'category', label: 'Category', type: 'enum', values: ['Hardware', 'Software', 'Services'], required: true },
            {
              key: 'price',
              label: 'Unit price',
              type: 'number',
              required: true,
              min: 0,
              decimals: 2,
              formula: 'forbid',
            },
            { key: 'currency', label: 'Currency', type: 'enum', values: ['USD', 'EUR', 'GBP'], required: true },
            { key: 'minimumOrder', label: 'Min. order qty', type: 'integer', required: true, min: 1, max: 10000 },
            { key: 'validFrom', label: 'Valid from', type: 'date', required: true, min: '2026-01-01' },
            {
              key: 'orderValue',
              label: 'Order value',
              type: 'number',
              formula: 'require',
              readOnly: true,
              description: { en: 'Worked out for you: unit price × minimum order.', de: 'Wird berechnet: Stückpreis × Mindestbestellmenge.' },
            },
          ],
          rules: [
            // Business check against the portal's price history, run on every edit so a
            // change of product code is checked too: prices more than 15% from the last approved one.
            async (table) => {
              const found = await Promise.all(
                table.records.map(async (r) => {
                  const { sku, price, currency } = r.values;
                  if (typeof sku !== 'string' || typeof price !== 'number' || typeof currency !== 'string') return undefined;
                  const approved = await lookupApprovedPrice(sku);
                  if (!approved || approved.currency !== currency || !approved.price) return undefined;
                  // Rounded, as the ratio of two prices isn't exact: 85 / 100 − 1 is −0.15000000000000002.
                  const change = Math.round((price / approved.price - 1) * 1e9) / 1e9;
                  if (Math.abs(change) <= 0.15) return undefined;
                  const say = (language: string) => [pct(change, language), money(approved.price, approved.currency, language)];
                  const [enPct, enMoney] = say('en');
                  const [dePct, deMoney] = say('de');
                  return {
                    record: r.index,
                    field: 'price',
                    severity: 'warning' as const,
                    message: {
                      en: `${enPct} on the last approved price of ${enMoney}. We’ll check this before approving.`,
                      de: `${dePct} gegenüber dem zuletzt genehmigten Preis von ${deMoney}. Wir prüfen das vor der Genehmigung.`,
                    },
                  };
                }),
              );
              return found.filter((x) => x !== undefined);
            },
            // A business rule in code: one currency for the whole list.
            (table) => {
              const first = table.records.find((r) => typeof r.values.currency === 'string')?.values.currency;
              if (typeof first !== 'string') return;
              return table.records
                .filter((r) => typeof r.values.currency === 'string' && r.values.currency !== first)
                .map((r) => ({
                  record: r.index,
                  field: 'currency',
                  message: { en: `Use one currency for the whole list (${first}, as in the first row).`, de: `Verwenden Sie eine Währung für die ganze Liste (${first}, wie in der ersten Zeile).` },
                }));
            },
          ],
        },
      ],
    },
  ],
  // Business logic: a product code that's in the catalogue fills in its description and category.
  onChange: async (ctx) => {
    for (const change of ctx.changes) {
      if (change.field !== 'sku' || typeof change.after !== 'string' || change.record === undefined) continue;
      const record = ctx.sheet(change.sheet)?.table(TABLE)?.records[change.record];
      if (!record || record.values.description) continue;
      const product = await lookupProduct(change.after);
      if (!product) continue;
      const target = { sheet: change.sheet, table: TABLE, record: change.record };
      ctx.set({ ...target, field: 'description' }, product.description);
      ctx.set({ ...target, field: 'category' }, product.category);
    }
  },
});
