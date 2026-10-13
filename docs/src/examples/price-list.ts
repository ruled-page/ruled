import { defineSchema } from '@sitepen/spreadsheet';
import { lookupSku } from './services';

export const priceList = defineSchema({
  id: 'docs.price-list',
  name: 'Price list',
  sheets: [
    {
      name: 'Prices',
      tables: [
        {
          key: 'prices',
          at: 'A1',
          records: { min: 1, max: 500 },
          fields: [
            {
              key: 'sku',
              label: 'SKU',
              type: 'text',
              required: true,
              unique: true,
              pattern: '^SKU-\\d{4}$',
              patternHint: 'like SKU-1234',
              // An async check against the catalogue service. The editor shows the cell
              // as being checked, aborts the call if the value changes again, and re-runs
              // it only when this cell's value changes.
              validate: async (value, { signal }) => {
                const item = await lookupSku(String(value), signal);
                if (item?.discontinued) return { message: 'The catalogue lists this product as discontinued.', severity: 'warning' };
              },
            },
            { key: 'description', label: 'Description', type: 'text', required: true },
            { key: 'category', label: 'Category', type: 'enum', values: ['Hardware', 'Software', 'Services'], required: true },
            { key: 'price', label: 'Unit price', type: 'number', required: true, min: 0, decimals: 2, formula: 'forbid' },
            { key: 'priceIncTax', label: 'Price inc. tax', type: 'number', formula: 'require', description: 'Unit price plus 20% tax.' },
            { key: 'from', label: 'Effective from', type: 'date', required: true, min: '2020-01-01' },
            { key: 'active', label: 'Active', type: 'boolean' },
          ],
        },
      ],
    },
  ],
  // Business logic after each edit: a new SKU fills in its description and category.
  onChange: async (ctx) => {
    for (const change of ctx.changes) {
      if (change.field !== 'sku' || typeof change.after !== 'string' || change.record === undefined) continue;
      const record = ctx.sheet(change.sheet)?.table('prices')?.records[change.record];
      if (!record || record.values.description) continue;
      const item = await lookupSku(change.after, ctx.signal);
      if (!item) continue;
      const at = { sheet: change.sheet, table: 'prices', record: change.record };
      ctx.set({ ...at, field: 'description' }, item.description);
      ctx.set({ ...at, field: 'category' }, item.category);
    }
  },
});
