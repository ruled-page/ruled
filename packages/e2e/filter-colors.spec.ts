import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { goTo, openFile, shown } from './helpers';

/**
 * Names filled red, none, green and red (Bob's text blue), scores with
 * conditional formatting's three arrows (Ann's and Di's down, Bob's level,
 * Cy's up), and dates: today, a February day and two others, with filter
 * buttons on A1:C5.
 */
async function coloured(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Colours');
  const now = new Date();
  const today = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  ws.addRow(['Name', 'Score', 'When']);
  const rows: [string, number, Date, string | null][] = [
    ['Ann', 10, new Date(Date.UTC(2020, 1, 14)), 'FFFF0000'],
    ['Bob', 50, today, null],
    ['Cy', 90, new Date(Date.UTC(2019, 6, 1)), 'FF00B050'],
    ['Di', 20, new Date(Date.UTC(2018, 9, 3)), 'FFFF0000'],
  ];
  for (const [name, score, when, fill] of rows) {
    const r = ws.addRow([name, score, when]);
    if (fill) r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    r.getCell(3).numFmt = 'yyyy-mm-dd';
  }
  ws.getCell('A3').font = { color: { argb: 'FF0070C0' } };
  ws.addConditionalFormatting({
    ref: 'B2:B5',
    rules: [{ type: 'iconSet', iconSet: '3Arrows', priority: 1, cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: 30 }, { type: 'num', value: 70 }] }],
  });
  ws.autoFilter = 'A1:C5';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const status = (page: Page) => page.locator('.wbe-status-filter');

/** Opens a column's filter menu from its heading. */
async function menu(page: Page, ref: string, heading: string) {
  await goTo(page, ref);
  await page.keyboard.press('Alt+ArrowDown');
  return page.getByRole('dialog', { name: `Show rows where “${heading}” is` });
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'colours.xlsx', await coloured());
});

test('filters a column by a fill colour, by having none, and by an icon', async ({ page }) => {
  let name = await menu(page, 'A1', 'Name');
  await name.getByRole('button', { name: 'Filter by colour' }).click();
  await expect(name.getByRole('button', { name: /^Colour / })).toHaveCount(3);
  await name.getByRole('button', { name: 'Colour #FF0000' }).click();
  await expect(name).toHaveCount(0);
  await expect(status(page)).toContainText('Filtered: 2 rows hidden');
  await expect(shown(page, 'A2')).toHaveText('Ann');
  await expect(shown(page, 'A3')).toHaveCount(0);

  name = await menu(page, 'A1', 'Name');
  await name.getByRole('button', { name: 'Filter by colour' }).click();
  await name.getByRole('button', { name: 'No fill' }).click();
  await expect(status(page)).toContainText('Filtered: 3 rows hidden');
  await expect(shown(page, 'A3')).toHaveText('Bob');

  // Another column's filter must hold too: Bob's arrow is the level one.
  const score = await menu(page, 'B1', 'Score');
  await score.getByRole('button', { name: 'Filter by colour' }).click();
  await score.getByRole('button', { name: 'Icon 1' }).click();
  await expect(status(page)).toContainText('Filtered: 4 rows hidden');
});

test("filters dates by Excel's periods: today, and a month of any year", async ({ page }) => {
  let when = await menu(page, 'C1', 'When');
  await when.getByRole('button', { name: 'Date filters' }).click();
  await when.getByRole('button', { name: 'Today', exact: true }).click();
  await expect(status(page)).toContainText('Filtered: 3 rows hidden');
  await expect(shown(page, 'A3')).toHaveText('Bob');

  when = await menu(page, 'C1', 'When');
  await when.getByRole('button', { name: 'Date filters' }).click();
  await when.getByRole('button', { name: 'All dates in the period' }).click();
  await when.getByRole('button', { name: 'February' }).click();
  await expect(shown(page, 'A2')).toHaveText('Ann');
  // Today in February is February too.
  await expect(status(page)).toContainText(new Date().getMonth() === 1 ? 'Filtered: 2 rows hidden' : 'Filtered: 3 rows hidden');
});

test('sorts by a colour on top from the filter menu, and at the bottom from the Sort dialog', async ({ page }) => {
  const name = await menu(page, 'A1', 'Name');
  await name.getByRole('button', { name: 'Sort by colour' }).click();
  await name.getByRole('button', { name: 'Colour #00B050' }).click();
  await expect(shown(page, 'A2')).toHaveText('Cy');
  await expect(shown(page, 'A3')).toHaveText('Ann');

  await goTo(page, 'A2');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Custom sort…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Sort' });
  await dialog.getByRole('combobox', { name: 'Sort by: Order' }).selectOption({ label: 'Cell colour #FF0000 on bottom' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  // The red rows go last, the rest keeping their order.
  for (const [ref, text] of [['A2', 'Cy'], ['A3', 'Bob'], ['A4', 'Ann'], ['A5', 'Di']]) await expect(shown(page, ref)).toHaveText(text);
});
