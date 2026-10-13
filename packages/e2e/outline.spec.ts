import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import { focusedGrid, goTo, grid as gridSurface, openFile, readWorkbook, shown } from './helpers';

/**
 * tests/corpus/outline.xlsx: on "Report", rows 2-7 grouped with their total
 * in row 8, North's lines 2-3 and South's 5-6 grouped inside them (South's
 * collapsed), and columns B:D and F:H grouped beside their totals in E and I
 * (F:H collapsed).
 */
const outline = () => new Uint8Array(readFileSync(new URL('../spreadsheet/tests/corpus/outline.xlsx', import.meta.url)));

const rowBar = (page: Page) => page.getByRole('group', { name: 'Row outline' });
const colBar = (page: Page) => page.getByRole('group', { name: 'Column outline' });
/** The sheet's numbers for the grid's first rows, as its table for screen readers numbers them. */
const rowNumbers = (page: Page, n: number) =>
  page
    .locator('table[role="grid"]')
    .first()
    .locator('tbody tr')
    .evaluateAll((rows, count) => rows.slice(0, count).map((r) => r.getAttribute('aria-rowindex')), n);

/** Opens the cell menu from the keyboard and chooses an item, in a submenu when one is named. */
async function choose(page: Page, item: string | RegExp, submenu?: string) {
  await page.keyboard.press('Shift+F10');
  if (submenu) {
    await page.getByRole('menuitem', { name: submenu }).focus();
    await page.keyboard.press('Enter');
  }
  await page.getByRole('menuitem', { name: item }).focus();
  await page.keyboard.press('Enter');
}

/** Whether a button's middle is y pixels below the top of the grid (a pixel either way, for the frozen pane's edge). */
async function at(page: Page, name: string, y: number): Promise<boolean> {
  const top = (await gridSurface(page).first().boundingBox())!.y;
  const button = (await page.getByRole('button', { name }).boundingBox())!;
  return Math.abs(button.y + button.height / 2 - top - y) <= 1;
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'outline.xlsx', outline());
});

test('shows the outline bars, and collapses and expands groups and shows levels with their buttons', async ({ page }) => {
  await expect(rowBar(page).getByRole('button', { name: /^Show row level/ })).toHaveCount(3);
  await expect(colBar(page).getByRole('button', { name: /^Show column level/ })).toHaveCount(2);
  const south = rowBar(page).getByRole('button', { name: 'Expand rows 5 to 6' });
  await expect(south).toHaveAttribute('aria-expanded', 'false');
  await expect(rowBar(page).getByRole('button', { name: 'Collapse rows 2 to 7' })).toHaveAttribute('aria-expanded', 'true');
  await expect(colBar(page).getByRole('button', { name: 'Expand columns F to H' })).toBeVisible();
  await expect.poll(() => rowNumbers(page, 6)).toEqual(['1', '2', '3', '4', '7', '8']);

  await goTo(page, 'A1');
  await south.click();
  await expect(rowBar(page).getByRole('button', { name: 'Collapse rows 5 to 6' })).toHaveAttribute('aria-expanded', 'true');
  await expect.poll(() => rowNumbers(page, 8)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8']);
  await rowBar(page).getByRole('button', { name: 'Show row level 1' }).click();
  await expect(rowBar(page).getByRole('button', { name: 'Expand rows 2 to 7' })).toBeVisible();
  await expect.poll(() => rowNumbers(page, 2)).toEqual(['1', '8']);

  await colBar(page).getByRole('button', { name: 'Show column level 2' }).click();
  await expect(colBar(page).getByRole('button', { name: 'Collapse columns F to H' })).toBeVisible();
  await expect(shown(page, 'F1')).toHaveText('Apr');

  // The buttons leave the focus in the grid, so undo is a key away.
  await expect(focusedGrid(page)).toHaveCount(1);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(colBar(page).getByRole('button', { name: 'Expand columns F to H' })).toBeVisible();

  const scan = await new AxeBuilder({ page }).include('.wbe-outline').analyze();
  expect(scan.violations).toEqual([]);
});

test('groups, hides the detail of and ungroups rows and columns from the keyboard, and saves the outline', async ({ page }) => {
  // For cells, the shortcut asks whether to group their rows or their columns.
  await goTo(page, 'A10:B11');
  await page.keyboard.press('Alt+Shift+ArrowRight');
  const dialog = page.getByRole('alertdialog', { name: 'Group' });
  await expect(dialog.getByRole('button', { name: 'Rows' })).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(rowBar(page).getByRole('button', { name: 'Collapse rows 10 to 11' })).toBeVisible();

  await choose(page, 'Hide detail', 'Rows');
  await expect(rowBar(page).getByRole('button', { name: 'Expand rows 10 to 11' })).toBeVisible();
  // The active cell moves out of the rows hidden, to the group's total.
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('A12');
  await choose(page, 'Show detail', 'Rows');
  await expect(rowBar(page).getByRole('button', { name: 'Collapse rows 10 to 11' })).toBeVisible();

  await goTo(page, 'K1:L1');
  await choose(page, /^Group columns/, 'Columns');
  await expect(colBar(page).getByRole('button', { name: 'Collapse columns K to L' })).toBeVisible();
  await goTo(page, 'L1');
  await page.keyboard.press('Alt+Shift+ArrowLeft');
  await page.getByRole('alertdialog', { name: 'Ungroup' }).getByRole('button', { name: 'Columns' }).click();
  await expect(colBar(page).getByRole('button', { name: 'Collapse column K' })).toBeVisible();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const report = (await readWorkbook(await readFile(await (await download).path()))).getWorksheet('Report')!;
  expect([report.getRow(10).outlineLevel, report.getRow(11).outlineLevel, report.getRow(12).outlineLevel]).toEqual([1, 1, 0]);
  expect([report.getColumn(11).outlineLevel, report.getColumn(12).outlineLevel]).toEqual([1, 0]);
  expect(report.getRow(6).hidden).toBe(true);
});

test('keeps the row outline in line with the rows as they scroll, and with frozen rows', async ({ page }) => {
  // Row 8, the 6th row shown, is 20 pixels high under the 24 of the column letters.
  const total = 'Collapse rows 2 to 7';
  await expect.poll(() => at(page, total, 24 + 5 * 20 + 10)).toBe(true);
  await goTo(page, 'A2');
  await choose(page, 'Freeze top row', 'Freeze panes');
  // Glide's grid puts frozen rows in a grid of their own; Ruled's in a group of rows.
  const frozen = { name: /^Frozen rows of /, includeHidden: true };
  await expect(page.getByRole('grid', frozen).or(page.getByRole('rowgroup', frozen))).toHaveCount(1);
  await expect.poll(() => at(page, total, 24 + 5 * 20 + 10)).toBe(true);
  // Scrolled down two rows, the button follows its row; scrolled well past it, it goes.
  const grid = (await gridSurface(page).last().boundingBox())!;
  await page.mouse.move(grid.x + grid.width / 2, grid.y + grid.height / 2);
  await page.mouse.wheel(0, 40);
  await expect.poll(() => at(page, total, 24 + 3 * 20 + 10)).toBe(true);
  await page.mouse.wheel(0, 400);
  await expect(page.getByRole('button', { name: total })).toHaveCount(0);
});
