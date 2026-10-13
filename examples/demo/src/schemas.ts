/**
 * Example schemas a consumer of the editor might register. They match the
 * sample workbooks in public/samples and show the three table shapes plus
 * async checks and business logic. The "services" are simulated with delays.
 *
 * Text the host writes for people (names, descriptions, messages) is given in
 * English and German, and the editor shows the one for its language. Field
 * labels are matched against the workbook's own headings, so they stay as
 * the workbook has them.
 */
import { defineSchema, type FieldValue } from '@ruled/sheet';

const wait = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(t);
      reject(new DOMException('Aborted', 'AbortError'));
    });
  });

/** A pretend product catalogue service. */
const CATALOGUE: Record<string, { description: string; category: string; discontinued?: boolean }> = {
  'SKU-1013': { description: 'Legacy adaptor', category: 'Hardware', discontinued: true },
  'SKU-2001': { description: 'Barcode scanner', category: 'Hardware' },
  'SKU-2002': { description: 'Label printer', category: 'Hardware' },
  'SKU-2003': { description: 'Analytics add-on', category: 'Software' },
  'SKU-2004': { description: 'On-site audit', category: 'Services' },
};

async function lookupSku(sku: string, signal?: AbortSignal) {
  await wait(500, signal);
  return CATALOGUE[sku.toUpperCase()];
}

export const priceListSchema = defineSchema({
  id: 'demo.price-list',
  name: { en: 'Price list', de: 'Preisliste' },
  description: {
    en: 'One product per row. Type a new SKU such as SKU-2001 in the empty row under the table to see the catalogue lookup fill it in.',
    de: 'Ein Produkt pro Zeile. Geben Sie in der leeren Zeile unter der Tabelle eine neue SKU wie SKU-2001 ein, um zu sehen, wie die Katalogsuche die Zeile ausfüllt.',
  },
  numbers: { display: 'full', copy: 'raw', paste: 'strict', reviewBeforeSave: true },
  sheets: [
    {
      name: 'Prices',
      tables: [
        {
          key: 'prices',
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
              patternHint: { en: 'like SKU-1234', de: 'wie SKU-1234' },
              // Async business check: the catalogue service flags discontinued products.
              validate: async (value, { signal }) => {
                const item = await lookupSku(String(value), signal);
                if (!item?.discontinued) return undefined;
                return {
                  message: { en: 'The catalogue lists this product as discontinued.', de: 'Laut Katalog wird dieses Produkt nicht mehr geführt.' },
                  severity: 'warning',
                };
              },
            },
            { key: 'description', label: 'Description', type: 'text', required: true },
            { key: 'category', label: 'Category', type: 'enum', values: ['Hardware', 'Software', 'Services'], required: true },
            { key: 'price', label: 'Unit price', type: 'number', required: true, min: 0, decimals: 2, formula: 'forbid' },
            {
              key: 'priceIncTax',
              label: 'Price inc. tax',
              type: 'number',
              formula: 'require',
              description: { en: 'Calculated: unit price plus 20% tax.', de: 'Berechnet: Stückpreis zuzüglich 20 % Steuer.' },
            },
            { key: 'from', label: 'Effective from', type: 'date', required: true, min: '2020-01-01' },
            { key: 'active', label: 'Active', type: 'boolean' },
          ],
        },
      ],
    },
  ],
  // Business logic: a new SKU fills its description and category from the catalogue.
  onChange: async (ctx) => {
    for (const change of ctx.changes) {
      if (change.field !== 'sku' || typeof change.after !== 'string' || change.record === undefined) continue;
      const record = ctx.sheet(change.sheet)?.table('prices')?.records[change.record];
      if (!record || record.values.description) continue;
      const item = await lookupSku(change.after);
      if (!item) continue;
      const target = { sheet: change.sheet, table: 'prices', record: change.record };
      ctx.set({ ...target, field: 'description' }, item.description);
      ctx.set({ ...target, field: 'category' }, item.category);
      if (!record.values.from) ctx.set({ ...target, field: 'from' }, new Date());
      if (record.values.active === null) ctx.set({ ...target, field: 'active' }, true);
    }
  },
});

const QUARTER = /^(\d{4})-Q([1-4])$/;

export const forecastSchema = defineSchema({
  id: 'demo.forecast',
  name: { en: 'Quarterly forecast', de: 'Quartalsprognose' },
  description: {
    en: 'One quarter per column. Type a period in the empty column after the table to add a quarter.',
    de: 'Ein Quartal pro Spalte. Geben Sie in der leeren Spalte nach der Tabelle einen Zeitraum ein, um ein Quartal hinzuzufügen.',
  },
  sheets: [
    {
      name: 'Forecast',
      tables: [
        {
          key: 'quarters',
          at: 'A1',
          orientation: 'columns',
          records: { min: 1, max: 12 },
          fields: [
            {
              key: 'period',
              label: 'Period',
              type: 'text',
              required: true,
              unique: true,
              pattern: QUARTER.source,
              patternHint: { en: 'a quarter like 2027-Q1', de: 'ein Quartal wie 2027-Q1' },
            },
            { key: 'revenue', label: 'Revenue', type: 'number', required: true, min: 0, decimals: 2 },
            { key: 'cost', label: 'Cost', type: 'number', required: true, min: 0, decimals: 2 },
            { key: 'margin', label: 'Margin', type: 'number', formula: 'require' },
            { key: 'marginPct', label: 'Margin %', type: 'percent', formula: 'require', min: 0 },
          ],
          rules: [
            // Business rule in code: quarters must follow on from each other.
            (table) => {
              const issues = [];
              let prev: number | undefined;
              for (const r of table.records) {
                const m = QUARTER.exec(String(r.values.period ?? ''));
                if (!m) continue;
                const q = Number(m[1]) * 4 + Number(m[2]);
                if (prev !== undefined && q !== prev + 1) {
                  issues.push({
                    record: r.index,
                    field: 'period',
                    message: {
                      en: 'Quarters should follow on from each other; a quarter is missing before this one.',
                      de: 'Die Quartale sollten aufeinander folgen; vor diesem fehlt ein Quartal.',
                    },
                  });
                }
                prev = q;
              }
              return issues;
            },
          ],
        },
      ],
    },
  ],
});

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

export const budgetSchema = defineSchema({
  id: 'demo.budget',
  name: { en: 'Half-year budget', de: 'Halbjahresbudget' },
  description: {
    en: 'A fixed grid: the categories and months are set by the schema; only the amounts can change.',
    de: 'Ein festes Raster: Die Kategorien und Monate legt das Schema fest; nur die Beträge können sich ändern.',
  },
  saveWithErrors: 'confirm',
  sheets: [
    {
      name: 'Budget',
      tables: [
        {
          key: 'budget',
          at: 'A1',
          orientation: 'rows',
          corner: 'Category',
          records: { labels: ['Salaries', 'Rent', 'Marketing', 'Travel', 'Equipment'] },
          fields: MONTHS.map((m) => ({ key: m.toLowerCase(), label: m, type: 'number' as const, required: true, min: 0, decimals: 2, formula: 'forbid' as const })),
          rules: [
            (table) =>
              table.records
                .filter((r) => r.label === 'Travel')
                .flatMap((r) =>
                  MONTHS.filter((m) => Number(r.values[m.toLowerCase()]) > 5000).map((m) => ({
                    record: r.index,
                    field: m.toLowerCase(),
                    severity: 'warning' as const,
                    message: { en: 'Travel is over the monthly cap of 5,000.', de: '„Travel“ liegt über der monatlichen Obergrenze von 5.000.' },
                  })),
                ),
          ],
        },
      ],
      cells: [
        {
          key: 'approvedBy',
          ref: 'B8',
          label: { ref: 'A8', text: 'Approved by' },
          type: 'text',
          required: true,
          description: { en: 'Who signed off this budget.', de: 'Wer dieses Budget freigegeben hat.' },
        },
        { key: 'total', ref: 'B9', label: { ref: 'A9', text: 'Total' }, type: 'number', formula: 'require', readOnly: true },
      ],
    },
  ],
  // Async workbook-wide check against a pretend budget service.
  validate: async ({ sheet, signal }) => {
    await wait(800, signal);
    const total = sheet('Budget')?.cell('total') as FieldValue;
    if (typeof total !== 'number' || total <= 400_000) return [];
    // Each language writes the total its own way, so the message stays right when the editor's language changes.
    const message = {
      en: `The budget service flags totals over 400,000 for review (this is ${total.toLocaleString('en')}).`,
      de: `Der Budgetdienst markiert Summen über 400.000 zur Überprüfung (diese Summe beträgt ${total.toLocaleString('de')}).`,
    };
    return [{ ref: 'Budget!B9', severity: 'warning', message }];
  },
});

export const DEMO_SCHEMAS = [priceListSchema, forecastSchema, budgetSchema];
