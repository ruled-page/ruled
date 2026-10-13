import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test('adds filter buttons with Ctrl+Shift+L, and a heading opens its filter menu with Alt+Down', async ({ page }) => {
  await goTo(page, 'B2');
  await page.keyboard.press('Control+Shift+L');
  await goTo(page, 'A1');
  await page.keyboard.press('Alt+ArrowDown');
  const menu = page.getByRole('dialog', { name: 'Show rows where “Item” is' });
  await expect(menu).toBeVisible();
  // Sorting from the menu sorts the filter's range by the column.
  await menu.getByRole('button', { name: 'Z to A' }).click();
  await expect(shown(page, 'A2')).toHaveText('Nuts');

  await page.keyboard.press('Alt+ArrowDown');
  await menu.getByRole('checkbox', { name: 'Bolts' }).uncheck();
  await menu.getByRole('button', { name: 'Show' }).click();
  await expect(page.locator('.wbe-status-filter')).toContainText('Filtered: 1 row hidden');

  // Turned off, the buttons go and every row shows again.
  await goTo(page, 'B2');
  await page.keyboard.press('Control+Shift+L');
  await expect(shown(page, 'A3')).toHaveText('Bolts');
});

test("opens a heading's filter menu from its button", async ({ page }) => {
  await goTo(page, 'B2');
  await page.keyboard.press('Control+Shift+L');
  // Turning the filter on is a step that can be undone: once it can, the buttons are there to click.
  await expect(page.getByRole('button', { name: 'Undo' })).toHaveAttribute('aria-disabled', 'false');
  const canvas = page.locator('canvas[data-testid="data-grid-canvas"]').first();
  const box = (await canvas.boundingBox())!;
  // The button sits at the right of the heading cell: column B ends 40 + 2 × 64 pixels in, under a 24 pixel heading.
  await page.mouse.click(box.x + 40 + 2 * 64 - 10, box.y + 24 + 10);
  await expect(page.getByRole('dialog', { name: 'Show rows where “Qty” is' })).toBeVisible();
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B2');
});
