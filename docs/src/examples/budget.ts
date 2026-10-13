import { defineSchema } from '@sitepen/spreadsheet';
import { needsReview } from './services';

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

export const budget = defineSchema({
  id: 'docs.budget',
  name: 'Half-year budget',
  // Errors ask before saving rather than stopping it.
  saveWithErrors: 'confirm',
  sheets: [
    {
      name: 'Budget',
      tables: [
        {
          key: 'budget',
          at: 'A1',
          corner: 'Category',
          // Fixed records: the categories can't be added to, removed or renamed.
          records: { labels: ['Salaries', 'Rent', 'Marketing', 'Travel', 'Equipment'] },
          fields: MONTHS.map((m) => ({ key: m.toLowerCase(), label: m, type: 'number' as const, required: true, min: 0, decimals: 2, formula: 'forbid' as const })),
          // A table rule in code: travel is capped each month.
          rules: [
            (table) =>
              table.records
                .filter((r) => r.label === 'Travel')
                .flatMap((r) =>
                  MONTHS.filter((m) => Number(r.values[m.toLowerCase()]) > 5000).map((m) => ({
                    record: r.index,
                    field: m.toLowerCase(),
                    severity: 'warning' as const,
                    message: 'Travel is over the monthly cap of 5,000.',
                  })),
                ),
          ],
        },
      ],
      cells: [
        { key: 'approvedBy', ref: 'B8', label: { ref: 'A8', text: 'Approved by' }, type: 'text', required: true, description: 'Who signed off this budget.' },
        { key: 'total', ref: 'B9', label: { ref: 'A9', text: 'Total' }, type: 'number', formula: 'require', readOnly: true },
      ],
    },
  ],
  // A workbook-wide async check against the budget service.
  validate: async ({ sheet, signal }) => {
    const total = sheet('Budget')?.cell('total');
    if (typeof total === 'number' && (await needsReview(total, signal))) {
      return [{ ref: 'Budget!B9', severity: 'warning', message: `Totals over 400,000 need review (this is ${total.toLocaleString('en')}).` }];
    }
  },
});
