import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

/** Opens an item of the cell menu's Conditional formatting menu. */
async function cfMenu(page: Page, name: string): Promise<void> {
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Conditional formatting' }).click();
  await page.getByRole('menuitem', { name }).click();
}

test("adds a rule from Excel's Highlight Cells menu, manages it, and saves it as Excel does", async ({ page }) => {
  await goTo(page, 'B2:B3');
  await cfMenu(page, 'Greater than…');
  const dialog = page.getByRole('dialog', { name: 'Greater Than' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByRole('alert')).toHaveText('Type a value, or a formula.');
  await dialog.getByRole('textbox', { name: 'Value' }).fill('3');
  await dialog.getByRole('combobox', { name: 'with' }).selectOption({ label: 'Green fill with dark green text' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);

  await cfMenu(page, 'Manage rules…');
  const manager = page.getByRole('dialog', { name: 'Conditional Formatting Rules Manager' });
  await expect(manager.getByRole('radio', { name: 'Cell value > 3' })).toBeChecked();
  await expect(manager.getByRole('textbox', { name: 'Applies to: Cell value > 3' })).toHaveValue('B2:B3');
  await manager.getByRole('textbox', { name: 'Applies to: Cell value > 3' }).fill('B2:B4');
  await manager.getByRole('button', { name: 'OK' }).click();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const wb = await readWorkbook(await readFile(await (await download).path()));
  // ExcelJS reads conditional formatting, though its types leave it out.
  const rules = (wb.getWorksheet('Lines') as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings;
  expect(rules).toHaveLength(1);
  expect(rules[0]).toMatchObject({ ref: 'B2:B4', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['3'] }] });
});

test('edits and deletes a rule in the Rules Manager, and clears rules from the sheet', async ({ page }) => {
  await goTo(page, 'C2:C3');
  await cfMenu(page, 'Top 10 items…');
  const top = page.getByRole('dialog', { name: 'Top 10 Items' });
  await top.getByRole('textbox', { name: 'Value' }).fill('1');
  await top.getByRole('button', { name: 'OK' }).click();
  await cfMenu(page, 'Data bars, color scales and icon sets…');
  await page.getByRole('dialog', { name: 'Data bars, color scales and icon sets…' }).getByRole('button', { name: 'Green data bar' }).click();

  await cfMenu(page, 'Manage rules…');
  const manager = page.getByRole('dialog', { name: 'Conditional Formatting Rules Manager' });
  // The newest rule comes first; the data bar can't stop the rules after it.
  await expect(manager.getByRole('radio')).toHaveCount(2);
  await expect(manager.getByRole('checkbox', { name: 'Stop if true: Data bar' })).toBeDisabled();
  await manager.getByRole('radio', { name: 'Top 1' }).check();
  await manager.getByRole('button', { name: 'Edit rule' }).click();
  const edit = page.getByRole('dialog', { name: 'Top 10 Items' });
  await expect(edit.getByRole('textbox', { name: 'Value' })).toHaveValue('1');
  await edit.getByRole('textbox', { name: 'Value' }).fill('2');
  await edit.getByRole('button', { name: 'OK' }).click();
  await expect(manager.getByRole('radio', { name: 'Top 2' })).toBeVisible();
  await manager.getByRole('radio', { name: 'Data bar' }).check();
  await manager.getByRole('button', { name: 'Delete rule' }).click();
  await expect(manager.getByRole('radio')).toHaveCount(1);
  await manager.getByRole('button', { name: 'OK' }).click();

  await cfMenu(page, 'Clear rules from entire sheet');
  await cfMenu(page, 'Manage rules…');
  await expect(manager.getByText('No rules.')).toBeVisible();
});
