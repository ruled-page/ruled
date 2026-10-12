import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test("sorts by levels through Excel's Sort dialog, from the cell menu", async ({ page }) => {
  await goTo(page, 'B2');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Custom sort…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Sort' });
  // The headings name the columns, and the active cell's column comes first.
  const sortBy = dialog.getByRole('combobox', { name: 'Sort by: Column' });
  await expect(sortBy).toBeFocused();
  await expect(sortBy.locator('option:checked')).toHaveText('Qty');
  await expect(dialog.getByRole('checkbox', { name: 'My data has headers' })).toBeChecked();

  // Sorting by the same column twice is refused, as in Excel.
  await dialog.getByRole('button', { name: 'Add level' }).click();
  await dialog.getByRole('combobox', { name: 'Then by: Column' }).selectOption({ label: 'Qty' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByRole('alert')).toHaveText('Qty is being sorted by more than once. Delete the duplicate level and try again.');
  await dialog.getByRole('button', { name: 'Delete level 2' }).click();

  await sortBy.selectOption({ label: 'Total' });
  await dialog.getByRole('combobox', { name: 'Sort by: Order' }).selectOption({ label: 'Smallest to Largest' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(shown(page, 'A2')).toHaveText('Nuts');
  await expect(shown(page, 'A3')).toHaveText('Bolts');
  await goTo(page, 'D2');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=B2*C2');
});
