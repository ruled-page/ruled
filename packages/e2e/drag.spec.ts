import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test("moves cells by dragging the selection's border, and copies them with Ctrl (Option on a Mac)", async ({ page }) => {
  const canvas = page.locator('canvas[data-testid="data-grid-canvas"]').first();
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
