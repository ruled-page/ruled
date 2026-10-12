import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook, setCell } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test("sets a rule with Excel's Data Validation dialog, enforces it, and saves it as Excel does", async ({ page }) => {
  await goTo(page, 'B2:B3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Data validation…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Data Validation' });
  await dialog.getByLabel('Allow').selectOption({ label: 'Whole number' });
  await dialog.getByLabel('Data').selectOption({ label: 'less than or equal to' });
  // A bound that isn't a whole number is refused.
  await dialog.getByLabel('Maximum').fill('9.5');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByText('Type a whole number, or a formula.')).toBeVisible();
  await dialog.getByLabel('Maximum').fill('10');
  await dialog.getByRole('tab', { name: 'Input message' }).click();
  await dialog.getByLabel('Title').fill('Quantity');
  await dialog.getByRole('textbox', { name: 'Input message' }).fill('Ten at most');
  await dialog.getByRole('tab', { name: 'Error alert' }).click();
  await dialog.getByLabel('Error message').fill('Too many');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);

  // The message shows on the cell, and a value breaking the rule is turned down.
  await goTo(page, 'B3');
  await expect(page.getByText('Quantity: Ten at most')).toBeVisible();
  await setCell(page, 'B3', '12');
  await expect(page.getByText('Too many')).toBeVisible();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const wb = await readWorkbook(await readFile(await (await download).path()));
  expect(wb.getWorksheet('Lines')!.getCell('B3').dataValidation).toMatchObject({ type: 'whole', operator: 'lessThanOrEqual', formulae: [10], prompt: 'Ten at most', error: 'Too many' });
});

test('clears the rule of the selected cells with Clear all', async ({ page }) => {
  await goTo(page, 'A2:A3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Data validation…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Data Validation' });
  await dialog.getByLabel('Allow').selectOption({ label: 'List' });
  await dialog.getByLabel('Source').fill('Bolts, Nuts, Washers');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await goTo(page, 'A2');
  await expect(page.getByText('Choose from the list: Bolts, Nuts, Washers')).toBeVisible();

  await goTo(page, 'A2:A3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Data validation…' }).click();
  await expect(dialog.getByLabel('Source')).toHaveValue('Bolts,Nuts,Washers');
  await dialog.getByRole('button', { name: 'Clear all' }).click();
  await goTo(page, 'A3');
  await expect(page.getByText('Choose from the list: Bolts, Nuts, Washers')).toHaveCount(0);
});
