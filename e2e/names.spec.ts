import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { openFile, ordersWorkbook, readWorkbook, shown } from './helpers';

/** Everything here is done with the keyboard. */

const grid = (page: Page) => page.locator('canvas[data-testid="data-grid-canvas"]:focus-within');
const nameBox = (page: Page) => page.getByRole('combobox', { name: 'Go to cell' });
const formulaBar = (page: Page) => page.getByRole('combobox', { name: 'Cell contents' });

/** From the sheet, types into the name box and presses Enter. */
async function nameBoxEnter(page: Page, text: string): Promise<void> {
  await expect(grid(page)).toHaveCount(1);
  await page.keyboard.press('Shift+F6');
  await expect(nameBox(page)).toBeFocused();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

/** Opens the cell menu from the keyboard and chooses an item. */
async function choose(page: Page, item: string): Promise<void> {
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: item }).focus();
  await page.keyboard.press('Enter');
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  // From the name box, which F6 reaches from anywhere, to the sheet.
  await nameBox(page).focus();
  await page.keyboard.press('Enter');
  await expect(grid(page)).toHaveCount(1);
});

test('names the selection from the name box, goes to the name, uses it in a formula and edits it in the name manager', async ({ page }) => {
  await nameBoxEnter(page, 'D2:D3');
  await expect(grid(page)).toHaveCount(1);
  await nameBoxEnter(page, 'Totals');
  await expect(grid(page)).toHaveCount(1);
  // The selection is exactly the named range, so the name box shows the name.
  await expect(nameBox(page)).toHaveValue('Totals');
  await expect(page.getByRole('button', { name: '1 change' })).toBeVisible();

  await nameBoxEnter(page, 'A1');
  await expect(nameBox(page)).toHaveValue('A1');
  // The list offers the names as they are typed; the arrow keys pick one.
  await expect(grid(page)).toHaveCount(1);
  await page.keyboard.press('Shift+F6');
  await expect(nameBox(page)).toBeFocused();
  await page.keyboard.type('tot');
  const list = page.getByRole('listbox', { name: 'Names' });
  await expect(list.getByRole('option')).toHaveText(['Totals']);
  await expect(nameBox(page)).toHaveAttribute('aria-expanded', 'true');
  await page.keyboard.press('ArrowDown');
  await expect(list.getByRole('option', { name: 'Totals' })).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Enter');
  await expect(grid(page)).toHaveCount(1);
  await expect(nameBox(page)).toHaveValue('Totals');
  await expect(formulaBar(page)).toHaveValue('=B2*C2');

  await nameBoxEnter(page, 'Summary!B2');
  await expect(grid(page)).toHaveCount(1);
  await page.keyboard.type('=SUM(Totals)*2', { delay: 30 });
  await page.keyboard.press('Enter');
  await expect(shown(page, 'B2')).toHaveText('16');

  // Rename it and point it at one cell; the formula follows the new name.
  await choose(page, 'Name manager…');
  const manager = page.getByRole('dialog', { name: 'Name manager' });
  await expect(manager.getByRole('radio', { name: 'Totals' })).toBeChecked();
  await expect(manager.getByRole('row', { name: /Totals/ })).toContainText('{6;2}');
  await page.keyboard.press('Tab');
  await expect(manager.getByRole('button', { name: 'Edit…' })).toBeFocused();
  await page.keyboard.press('Enter');
  const edit = page.getByRole('dialog', { name: 'Edit name' });
  await expect(edit.getByRole('textbox', { name: 'Name' })).toBeFocused();
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.type('LineTotals');
  await page.keyboard.press('Tab');
  await page.keyboard.type('The line totals');
  await page.keyboard.press('Tab');
  await expect(edit.getByRole('textbox', { name: 'Refers to' })).toBeFocused();
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.type('=Lines!$D$2');
  await page.keyboard.press('Enter');
  await expect(edit).toBeHidden();
  await expect(manager.getByRole('button', { name: 'Edit…' })).toBeFocused();
  await expect(manager.getByRole('row', { name: /LineTotals/ })).toContainText('=Lines!$D$2');
  await expect(manager.getByRole('row', { name: /LineTotals/ })).toContainText('The line totals');
  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await expect(shown(page, 'B2')).toHaveText('12');
  await nameBoxEnter(page, 'B2');
  await expect(formulaBar(page)).toHaveValue('=SUM(LineTotals)*2');

  const download = page.waitForEvent('download');
  await page.keyboard.press('ControlOrMeta+s');
  const wb = await readWorkbook(await readFile(await (await download).path()));
  expect(wb.definedNames.getRanges('LineTotals').ranges).toEqual(['Lines!$D$2']);
  expect(wb.getWorksheet('Summary')!.getCell('B2').formula).toBe('SUM(LineTotals)*2');
});

test('the name dialog says what is wrong with a name, on the field, and undo takes the name away', async ({ page }) => {
  await nameBoxEnter(page, 'B2:B3');
  await expect(grid(page)).toHaveCount(1);
  await choose(page, 'Define name…');
  const dialog = page.getByRole('dialog', { name: 'New name' });
  const name = dialog.getByRole('textbox', { name: 'Name' });
  await expect(name).toBeFocused();
  await expect(dialog.getByRole('textbox', { name: 'Refers to' })).toHaveValue('=Lines!$B$2:$B$3');
  await page.keyboard.type('B2');
  await page.keyboard.press('Enter');
  await expect(name).toHaveAttribute('aria-invalid', 'true');
  await expect(name).toHaveAccessibleDescription('"B2" is a cell reference, so it can\'t be a name.');
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.type('Quantities');
  // The scope can be one sheet.
  await page.keyboard.press('Tab');
  await page.keyboard.type('Lines');
  await expect(dialog.getByRole('combobox', { name: 'Scope' })).toHaveValue('0');
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  await expect(grid(page)).toHaveCount(1);
  await expect(nameBox(page)).toHaveValue('Quantities');

  // Text that is neither a reference nor a name is refused, and the name box keeps it to correct.
  await nameBoxEnter(page, 'two words');
  await expect(page.locator('.wbe-toasts')).toContainText('"two words" isn\'t a cell, a range or a name in this workbook.');
  await expect(nameBox(page)).toBeFocused();
  // Alt+Down lists the names; Escape closes the list, then leaves the box as it was.
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Alt+ArrowDown');
  await expect(page.getByRole('listbox', { name: 'Names' }).getByRole('option')).toHaveText(['Quantities']);
  await page.keyboard.press('Escape');
  await expect(nameBox(page)).toHaveAttribute('aria-expanded', 'false');
  await expect(nameBox(page)).toBeFocused();
  await page.keyboard.type('B2:B3');
  await page.keyboard.press('Enter');
  await expect(nameBox(page)).toHaveValue('Quantities');
  await expect(grid(page)).toHaveCount(1);

  await page.keyboard.press('ControlOrMeta+z');
  await expect(nameBox(page)).toHaveValue('B2:B3');
  await expect(page.getByRole('button', { name: '1 change' })).toBeHidden();
});
