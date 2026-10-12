import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test('inserts cells from the keyboard, asking which way to move the others as Excel does', async ({ page }) => {
  await goTo(page, 'B2');
  await page.keyboard.press('Control+Shift+Equal');
  const dialog = page.getByRole('dialog', { name: 'Insert' });
  // One cell moves the others down unless asked otherwise.
  await expect(dialog.getByRole('radio', { name: 'Shift cells down' })).toBeChecked();
  await expect(dialog.getByRole('radio', { name: 'Shift cells down' })).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(shown(page, 'B2')).toHaveText('');
  await expect(shown(page, 'B3')).toHaveText('2');
  await goTo(page, 'D2');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=B3*C2');
});

test("deletes cells from the cell's menu, the cells to the right moving back", async ({ page }) => {
  await goTo(page, 'B2:B3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Delete cells…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Delete' });
  // A selection taller than it is wide closes the gap from the side.
  await expect(dialog.getByRole('radio', { name: 'Shift cells left' })).toBeChecked();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(shown(page, 'B2')).toHaveText('3');
  await expect(shown(page, 'B3')).toHaveText('0.5');
});

test('inserts and deletes selected rows straight away', async ({ page }) => {
  await goTo(page, 'A2');
  await page.keyboard.press('Shift+Space');
  await page.keyboard.press('Control+Minus');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(shown(page, 'A2')).toHaveText('Nuts');
  await page.keyboard.press('Control+Shift+Equal');
  await expect(shown(page, 'A2')).toHaveText('');
  await expect(shown(page, 'A3')).toHaveText('Nuts');
});
