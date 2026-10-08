import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { de } from '../src/demo/locales/de';
import { en } from '../src/demo/locales/en';
import { budgetSchema, DEMO_SCHEMAS, forecastSchema, priceListSchema } from '../src/demo/schemas';
import { DEMO_THEMES } from '../src/demo/themes';
import type { EditorTheme } from '../src/editor/theme';
import { HeadlessWorkbook } from './headless';
import type { FieldValue, WorkbookData, WorkbookSchema } from '../src/schema/types';
import { sample } from './helpers';

// The demo's pretend services answer after 500–800 ms; fake timers play that out at once.
beforeEach(() => {
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
  vi.setSystemTime(new Date(2026, 9, 7, 12));
});
afterEach(() => vi.useRealTimers());

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

const open = (file: string, schema: WorkbookSchema) => {
  const wb = HeadlessWorkbook.open(sample(file), { fileName: file, schemas: [schema] });
  const issuesAt = async (ref: string) => (await settle(wb.issues())).filter((i) => i.ref === ref).map((i) => i.message);
  const set = (sheet: string, ref: string, text: string) => settle(wb.set(sheet, ref, text));
  return { wb, issuesAt, set };
};

describe('the demo schemas', () => {
  it('each match their own sample and no other', () => {
    const samples = ['price-list.xlsx', 'forecast.xlsx', 'budget.xlsx'];
    for (const [i, file] of samples.entries()) {
      expect(HeadlessWorkbook.open(sample(file), { schemas: DEMO_SCHEMAS }).schema).toBe(DEMO_SCHEMAS[i]);
    }
  });
});

describe('the price list demo', () => {
  it('fills a new product in from the catalogue, effective today and active', async () => {
    const { wb, set } = open('price-list.xlsx', priceListSchema);
    await set('Prices', 'A16', 'SKU-2004');
    expect(wb.range('Prices', 'A16:G16')).toEqual([['SKU-2004', 'On-site audit', 'Services', null, 0, '2026-10-07', true]]);
  });

  it('keeps an effective date and active flag already typed', async () => {
    const { wb, set } = open('price-list.xlsx', priceListSchema);
    await set('Prices', 'F16', '2026-12-01');
    await set('Prices', 'G16', 'no');
    await set('Prices', 'A16', 'SKU-2001');
    expect(wb.range('Prices', 'A16:G16')).toEqual([['SKU-2001', 'Barcode scanner', 'Hardware', null, 0, '2026-12-01', false]]);
  });

  it('leaves a description already written, and fills nothing for an unknown code', async () => {
    const { wb, set } = open('price-list.xlsx', priceListSchema);
    await set('Prices', 'A2', 'SKU-2003');
    expect(wb.range('Prices', 'A2:C2')).toEqual([['SKU-2003', 'Steel bracket, small', 'Hardware']]);
    await set('Prices', 'A16', 'SKU-2999');
    expect(wb.range('Prices', 'A16:C16')).toEqual([['SKU-2999', null, null]]);
  });

  it('fills in a discontinued product too, with a warning', async () => {
    const { wb, set, issuesAt } = open('price-list.xlsx', priceListSchema);
    await set('Prices', 'A16', 'SKU-1013');
    expect(wb.cell('Prices', 'B16')).toBe('Legacy adaptor');
    expect(await issuesAt('A16')).toEqual(['Duplicate: A14 has the same SKU.', 'The catalogue lists this product as discontinued.']);
  });
});

describe('the forecast demo', () => {
  const gap = 'Quarters should follow on from each other; a quarter is missing before this one.';

  it('takes the next quarter in the column after the table', async () => {
    const { wb, set, issuesAt } = open('forecast.xlsx', forecastSchema);
    await set('Forecast', 'F1', '2027-Q2');
    expect(await issuesAt('F1')).toEqual([]);
    expect(wb.values()!.Forecast.tables.quarters.map((q) => q.period)).toEqual(['2026-Q1', '2026-Q2', '2026-Q4', '2027-Q1', '2027-Q2']);
  });

  it('wants a period for a quarter started without one, and doesn’t call it a gap', async () => {
    const { set, issuesAt } = open('forecast.xlsx', forecastSchema);
    await set('Forecast', 'F2', '160000');
    expect(await issuesAt('F1')).toEqual(['Required.']);
  });

  it('flags a quarter that skips one', async () => {
    const { set, issuesAt } = open('forecast.xlsx', forecastSchema);
    await set('Forecast', 'F1', '2027-Q4');
    expect(await issuesAt('F1')).toEqual([gap]);
  });

  it('moves the gap on when a quarter is corrected', async () => {
    const { set, issuesAt } = open('forecast.xlsx', forecastSchema);
    expect(await issuesAt('D1')).toEqual([gap]);
    await set('Forecast', 'D1', '2026-Q3');
    expect(await issuesAt('D1')).toEqual([]);
    expect(await issuesAt('E1')).toEqual([gap]);
  });

  it('passes over a period that isn’t a quarter, comparing the next with the last quarter', async () => {
    const { set, issuesAt } = open('forecast.xlsx', forecastSchema);
    await set('Forecast', 'D1', 'Q3');
    expect(await issuesAt('D1')).toEqual(['Should be a quarter like 2027-Q1.']);
    expect(await issuesAt('E1')).toEqual([gap]);
  });
});

describe('the budget demo', () => {
  const cap = 'Travel is over the monthly cap of 5,000.';

  it('warns about travel over the monthly cap, but not at it', async () => {
    const { set, issuesAt } = open('budget.xlsx', budgetSchema);
    await set('Budget', 'G5', '5000');
    expect(await issuesAt('G5')).toEqual([]);
    await set('Budget', 'G5', '5000.01');
    expect(await issuesAt('G5')).toEqual([cap]);
    await set('Budget', 'D5', '4000');
    expect(await issuesAt('D5')).toEqual([]);
    // Only travel has a cap.
    await set('Budget', 'B4', '9000');
    expect(await issuesAt('B4')).toEqual([]);
  });

  it('asks the budget service about a total over 400,000', async () => {
    const { wb, set, issuesAt } = open('budget.xlsx', budgetSchema);
    expect(await issuesAt('B9')).toEqual([]);
    await set('Budget', 'B2', '62000');
    expect(wb.cell('Budget', 'B9')).toBe(409650.5);
    expect(await issuesAt('B9')).toEqual(['The budget service flags totals over 400,000 for review (this is 409,650.5).']);
  });

  it('writes the total each language’s own way', async () => {
    const check = (total: FieldValue, signal = new AbortController().signal) =>
      settle(Promise.resolve(budgetSchema.validate!({ language: 'de', signal, sheet: () => ({ name: 'Budget', table: () => undefined, cell: () => total }) } as WorkbookData)));
    expect(await check(409650.5)).toEqual([
      {
        ref: 'Budget!B9',
        severity: 'warning',
        message: {
          en: 'The budget service flags totals over 400,000 for review (this is 409,650.5).',
          de: 'Der Budgetdienst markiert Summen über 400.000 zur Überprüfung (diese Summe beträgt 409.650,5).',
        },
      },
    ]);
    expect(await check(400_000)).toEqual([]);
    expect(await check({ error: '#REF!' })).toEqual([]);
  });

  it('stops waiting for the budget service when the workbook changes again', async () => {
    const controller = new AbortController();
    const pending = Promise.resolve(budgetSchema.validate!({ language: 'en', signal: controller.signal, sheet: () => undefined }));
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('can be saved with errors after asking, and without asking once they are fixed', async () => {
    const { wb, set } = open('budget.xlsx', budgetSchema);
    expect(await settle(wb.checkBeforeSave())).toEqual({ ok: false, message: '1 cell breaks the “Half-year budget” schema.', canOverride: true });
    await set('Budget', 'B8', 'Dana');
    expect(await settle(wb.checkBeforeSave())).toEqual({ ok: true });
  });
});

/** WCAG contrast ratio of two #rrggbb colours. */
function contrast(a: string, b: string): number {
  const luminance = (hex: string) => {
    const [r, g, bl] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * bl;
  };
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

describe('the demo themes', () => {
  it('are named in the demo’s English and German, with ids of their own', () => {
    for (const t of DEMO_THEMES) {
      for (const catalogue of [en, de] as Record<string, string>[]) expect(catalogue[t.label], t.id).toBeTruthy();
    }
    expect(new Set(DEMO_THEMES.map((t) => t.id)).size).toBe(DEMO_THEMES.length);
    expect(DEMO_THEMES[0]).toEqual({ id: 'default', label: 'theme.default' });
    expect(DEMO_THEMES.find((t) => t.id === 'dark')?.theme?.colorScheme).toBe('dark');
  });

  it('keep text readable on the background it is drawn on', () => {
    const pairs: [keyof EditorTheme, keyof EditorTheme][] = [
      ['text', 'surface'],
      ['textMuted', 'surface'],
      ['textSubtle', 'surface'],
      ['gridText', 'gridBg'],
      ['gridTextMuted', 'gridBg'],
      ['gridHeaderText', 'gridHeaderBg'],
      ['gridHeaderSelectedText', 'gridHeaderSelectedBg'],
      ['onAccent', 'accent'],
      ['infoBannerText', 'infoBannerBg'],
      ['warningText', 'warningSoft'],
      ['successText', 'successSoft'],
      ['fieldText', 'fieldSoft'],
      ['noteText', 'noteSoft'],
      ['infoText', 'infoSoft'],
    ];
    let checked = 0;
    for (const { id, theme } of DEMO_THEMES) {
      for (const [fg, bg] of pairs) {
        const [a, b] = [theme?.[fg], theme?.[bg]];
        if (!a?.startsWith('#') || !b?.startsWith('#')) continue;
        expect(contrast(a, b), `${id}: ${fg} on ${bg}`).toBeGreaterThanOrEqual(4.5);
        checked++;
      }
    }
    expect(checked).toBeGreaterThan(10);
  });
});
