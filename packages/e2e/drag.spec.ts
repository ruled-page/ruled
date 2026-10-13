import { expect, test } from '@playwright/test';
import { goTo, grid, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test("moves cells by dragging the selection's border, and copies them with Ctrl (Option on a Mac)", async ({ page }) => {
  const canvas = grid(page).first();
  const box = (await canvas.boundingBox())!;
  // Rows are 20 pixels under a 24 pixel heading; columns 64 pixels after the 40 pixel row numbers.
  const x = (col: number) => box.x + 40 + 64 * col + 32;
  const bottom = (row: number) => box.y + 24 + 20 * (row + 1);
  const middle = (row: number) => box.y + 24 + 20 * row + 10;

  await goTo(page, 'B2');
  await page.mouse.move(x(1), bottom(1));
  await expect(page.locator('.wbe-grid-host').first()).toHaveCSS('cursor', 'move');
  await page.mouse.down();
  await page.mouse.move(x(1), middle(5), { steps: 5 });
  await page.mouse.up();
  await expect(shown(page, 'B6')).toHaveText('2');
  await expect(shown(page, 'B2')).toHaveText('');
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B6');
  // The total that read it follows it.
  await goTo(page, 'D2');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=B6*C2');
  await expect(shown(page, 'D2')).toHaveText('6');

  const copy = (await page.evaluate(() => /Mac|iPhone|iPad/.test(navigator.platform))) ? 'Alt' : 'Control';
  await goTo(page, 'B6');
  await page.mouse.move(x(1), bottom(5));
  await page.mouse.down();
  await page.keyboard.down(copy);
  await page.mouse.move(x(4), middle(5), { steps: 5 });
  await page.mouse.up();
  await page.keyboard.up(copy);
  await expect(shown(page, 'E6')).toHaveText('2');
  await expect(shown(page, 'B6')).toHaveText('2');
});

test('moves cells in between others with Shift held, a line showing where they go', async ({ page }) => {
  const canvas = grid(page).first();
  const box = (await canvas.boundingBox())!;
  const x = box.x + 40 + 32;
  await goTo(page, 'A3:D3');
  await page.mouse.move(x, box.y + 24 + 20 * 3);
  await page.mouse.down();
  await page.keyboard.down('Shift');
  // Near the top edge of row 2: in between the headings and Bolts.
  await page.mouse.move(x, box.y + 24 + 20 + 3, { steps: 5 });
  await expect(page.locator('.wbe-insert-guide')).toBeVisible();
  await page.mouse.up();
  await page.keyboard.up('Shift');
  await expect(page.locator('.wbe-insert-guide')).toHaveCount(0);
  await expect(shown(page, 'A2')).toHaveText('Nuts');
  await expect(shown(page, 'A3')).toHaveText('Bolts');
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('A2:D2');
  await goTo(page, 'D2');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=B2*C2');
  // The grand total still takes in both rows.
  await goTo(page, 'Summary!B1');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=SUM(Lines!D2:D3)');
  await expect(shown(page, 'B1')).toHaveText('8');
});
