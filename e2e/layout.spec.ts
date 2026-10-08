import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook, shown } from './helpers';

/** Opens the cell menu from the keyboard and chooses an item, in a submenu when one is named. */
async function choose(page: Page, item: string, submenu?: string) {
  await page.keyboard.press('Shift+F10');
  if (submenu) {
    await page.getByRole('menuitem', { name: submenu }).focus();
    await page.keyboard.press('Enter');
  }
  await page.getByRole('menuitem', { name: item }).focus();
  await page.keyboard.press('Enter');
}

async function save(page: Page) {
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  return readWorkbook(await readFile(await (await download).path()));
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test('hides and unhides a column, sets a row height, and saves both', async ({ page }) => {
  await goTo(page, 'B2');
  await choose(page, 'Hide column', 'Columns');
  // The grid's second column is now C, with C2's 3 where B2's 2 was.
  await expect(shown(page, 'B2')).toHaveText('3');
  await expect(page.getByRole('button', { name: '1 change' })).toBeVisible();
  // The column right of it finds it.
  await goTo(page, 'C2');
  await choose(page, 'Unhide columns', 'Columns');
  await expect(shown(page, 'B2')).toHaveText('2');

  await goTo(page, 'A3');
  await choose(page, 'Row height…', 'Rows');
  const dialog = page.getByRole('dialog', { name: 'Height of row 3' });
  await dialog.getByRole('textbox').fill('30');
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  await goTo(page, 'B2');
  await choose(page, 'Hide column', 'Columns');

  const wb = await save(page);
  const lines = wb.getWorksheet('Lines')!;
  expect(lines.getRow(3).height).toBe(30);
  expect(lines.getColumn(2).hidden).toBe(true);
});

test('merges cells after asking about the values that go, and freezes the top row', async ({ page }) => {
  await goTo(page, 'A2:B3');
  await choose(page, 'Merge cells');
  const confirm = page.getByRole('alertdialog', { name: 'Merge cells?' });
  await expect(confirm).toContainText('3 other cells will be cleared');
  await confirm.getByRole('button', { name: 'Merge and clear' }).click();
  await expect(shown(page, 'A2')).toHaveText('Bolts');
  // The merged cell's formulas went with the values, so the total changes.
  await expect(shown(page, 'D2')).toHaveText('0');

  await goTo(page, 'C3');
  await choose(page, 'Freeze top row', 'Freeze panes');
  await expect(page.locator('.wbe-pane-top')).toHaveCount(1);

  const wb = await save(page);
  const lines = wb.getWorksheet('Lines')!;
  expect(lines.getCell('B3').isMerged).toBe(true);
  expect(lines.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
});

test('resizes a row by dragging the edge of its number', async ({ page }) => {
  await goTo(page, 'A2');
  const canvas = page.locator('canvas[data-testid="data-grid-canvas"]').first();
  const box = (await canvas.boundingBox())!;
  // Row 2's number, then the edge below it: rows are 20 pixels under a 24 pixel heading.
  const x = box.x + 10;
  const edge = box.y + 24 + 20 * 2;
  await page.mouse.move(x, edge);
  await page.mouse.down();
  await page.mouse.move(x, edge + 20, { steps: 4 });
  await page.mouse.up();
  await expect(page.getByRole('button', { name: '1 change' })).toBeVisible();
  // The press didn't select the row number's cell.
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('A2');
  const wb = await save(page);
  expect(wb.getWorksheet('Lines')!.getRow(2).height).toBe(30);
});

test("keeps a column's width when its edge is only pressed, and sets it when dragged", async ({ page }) => {
  await goTo(page, 'A2');
  const canvas = page.locator('canvas[data-testid="data-grid-canvas"]').first();
  const box = (await canvas.boundingBox())!;
  // Column A's right edge, in the heading: a 40 pixel row number column, then A at 64 pixels.
  const edge = box.x + 40 + 64;
  const width = async () => {
    await choose(page, 'Column width…', 'Columns');
    const dialog = page.getByRole('dialog', { name: 'Width of column A' });
    const value = await dialog.getByRole('textbox').inputValue();
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    return value;
  };
  await page.mouse.move(edge, box.y + 12);
  await page.mouse.down();
  await page.mouse.up();
  await goTo(page, 'A2');
  expect(await width()).toBe('9.14');
  await expect(page.getByRole('button', { name: /change/ })).toHaveCount(0);

  await page.mouse.move(edge, box.y + 12);
  await page.mouse.down();
  await page.mouse.move(edge + 36, box.y + 12, { steps: 4 });
  await page.mouse.up();
  await expect(page.getByRole('button', { name: '1 change' })).toBeVisible();
  await goTo(page, 'A2');
  expect(await width()).toBe('14.29');
});
