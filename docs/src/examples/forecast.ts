import { defineSchema } from '@ruled/sheet';

const QUARTER = /^(\d{4})-Q([1-4])$/;

export const forecast = defineSchema({
  id: 'docs.forecast',
  name: 'Quarterly forecast',
  sheets: [
    {
      name: 'Forecast',
      tables: [
        {
          key: 'quarters',
          at: 'A1',
          // One record per column: each quarter runs down its own column.
          orientation: 'columns',
          records: { min: 1, max: 12 },
          fields: [
            { key: 'period', label: 'Period', type: 'text', required: true, unique: true, pattern: QUARTER.source, patternHint: 'a quarter like 2027-Q1' },
            { key: 'revenue', label: 'Revenue', type: 'number', required: true, min: 0, decimals: 2 },
            { key: 'cost', label: 'Cost', type: 'number', required: true, min: 0, decimals: 2 },
            { key: 'margin', label: 'Margin', type: 'number', formula: 'require' },
            { key: 'marginPct', label: 'Margin %', type: 'percent', formula: 'require', min: 0 },
          ],
          rules: [
            // Quarters must follow on from each other.
            (table) => {
              const issues = [];
              let previous: number | undefined;
              for (const r of table.records) {
                const m = QUARTER.exec(String(r.values.period ?? ''));
                if (!m) continue;
                const q = Number(m[1]) * 4 + Number(m[2]);
                if (previous !== undefined && q !== previous + 1) {
                  issues.push({ record: r.index, field: 'period', message: 'A quarter is missing before this one.' });
                }
                previous = q;
              }
              return issues;
            },
          ],
        },
      ],
    },
  ],
});
