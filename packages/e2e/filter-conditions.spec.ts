import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await goTo(page, 'B2');
  await page.keyboard.press('Control+Shift+L');
});

test("filters a column by a condition through Excel's Number Filters and Custom AutoFilter", async ({ page }) => {
  await goTo(page, 'B1');
  await page.keyboard.press('Alt+ArrowDown');
  const menu = page.getByRole('dialog', { name: 'Show rows where “Qty” is' });
  await menu.getByRole('button', { name: 'Number filters' }).click();
  await menu.getByRole('button', { name: 'Greater than…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Custom AutoFilter' });
  await expect(dialog.getByRole('combobox', { name: 'First rule: comparison' })).toHaveValue('greaterThan');
  // Something that isn't a number can't be compared with a column of numbers.
  await dialog.getByRole('combobox', { name: 'First rule: value' }).fill('lots');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByText('Type a number to compare with.')).toBeVisible();
  await dialog.getByRole('combobox', { name: 'First rule: value' }).fill('3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.wbe-status-filter')).toContainText('Filtered: 1 row hidden');
  // Bolts' row is hidden, and Nuts' row shown.
  await expect(shown(page, 'A2')).toHaveCount(0);
  await expect(shown(page, 'A3')).toHaveText('Nuts');

  // Opened again, the dialog shows the rule; Clear shows every row.
  await goTo(page, 'B1');
  await page.keyboard.press('Alt+ArrowDown');
  await menu.getByRole('button', { name: 'Number filters' }).click();
  await menu.getByRole('button', { name: 'Custom filter…' }).click();
  await expect(dialog.getByRole('combobox', { name: 'First rule: value' })).toHaveValue('3');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await goTo(page, 'B1');
  await page.keyboard.press('Alt+ArrowDown');
  await menu.getByRole('button', { name: 'Clear', exact: true }).click();
  await expect(page.locator('.wbe-status-filter')).toHaveCount(0);
});

test('keeps the top numbers with Top 10, and text that begins with what is typed', async ({ page }) => {
  await goTo(page, 'C1');
  await page.keyboard.press('Alt+ArrowDown');
  const price = page.getByRole('dialog', { name: 'Show rows where “Price” is' });
  await price.getByRole('button', { name: 'Number filters' }).click();
  await price.getByRole('button', { name: 'Top 10…' }).click();
  const top = page.getByRole('dialog', { name: 'Top 10 AutoFilter' });
  await top.getByRole('textbox', { name: 'How many' }).fill('1');
  await top.getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.wbe-status-filter')).toContainText('Filtered: 1 row hidden');
  await expect(shown(page, 'A2')).toHaveText('Bolts');

  await goTo(page, 'A1');
  await page.keyboard.press('Alt+ArrowDown');
  const item = page.getByRole('dialog', { name: 'Show rows where “Item” is' });
  await item.getByRole('button', { name: 'Text filters' }).click();
  await item.getByRole('button', { name: 'Begins with…' }).click();
  const custom = page.getByRole('dialog', { name: 'Custom AutoFilter' });
  await custom.getByRole('combobox', { name: 'First rule: value' }).fill('nu');
  await custom.getByRole('button', { name: 'OK' }).click();
  // Both columns' conditions must hold: the largest price is Bolts', which doesn't begin with "nu".
  await expect(page.locator('.wbe-status-filter')).toContainText('Filtered: 2 rows hidden');
});
