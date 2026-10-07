import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, openSample, ordersWorkbook, shown } from './helpers';

/** Everything here is done without the mouse. */

const grid = (page: Page) => page.locator('canvas[data-testid="data-grid-canvas"]:focus-within');
const nameBox = (page: Page) => page.getByRole('textbox', { name: 'Go to cell' });

test.describe('with only a keyboard', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await goTo(page, 'A1');
  });

  test('F6 moves between the parts of the editor and back to the sheet', async ({ page }) => {
    await page.keyboard.press('F6');
    await expect(page.getByRole('navigation', { name: 'Sheets' }).getByRole('button', { name: 'Lines' })).toBeFocused();
    await page.keyboard.press('F6');
    await expect(page.locator('.toolbar :focus')).toHaveCount(1);
    await page.keyboard.press('F6');
    await expect(nameBox(page)).toBeFocused();
    await page.keyboard.press('F6');
    await expect(grid(page)).toHaveCount(1);
    await page.keyboard.press('Shift+F6');
    await expect(nameBox(page)).toBeFocused();
  });

  test('Ctrl+G goes to a cell', async ({ page }) => {
    await page.keyboard.press('Control+g');
    await expect(nameBox(page)).toBeFocused();
    await page.keyboard.type('C3');
    await page.keyboard.press('Enter');
    await expect(grid(page)).toHaveCount(1);
    await expect(page.getByRole('textbox', { name: 'Cell contents' })).toHaveValue('0.5');
  });

  test('Enter edits a cell, Escape cancels, typing replaces', async ({ page }) => {
    await goTo(page, 'B2');
    await page.keyboard.press('Enter');
    await page.keyboard.press('End');
    await page.keyboard.type('0', { delay: 20 });
    await page.keyboard.press('Escape');
    await expect(shown(page, 'B2')).toHaveText('2');
    await page.keyboard.type('7', { delay: 20 });
    await page.keyboard.press('Enter');
    await expect(shown(page, 'D2')).toHaveText('21');
  });

  test('Shift+F10 opens the cell menu, which works with the arrow keys', async ({ page }) => {
    await goTo(page, 'A2');
    await page.keyboard.press('Shift+F10');
    const menu = page.getByRole('menu');
    await expect(menu.getByRole('menuitem', { name: 'Cut' })).toBeFocused();
    // Cut, Copy, Paste, then the first row item.
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expect(menu.getByRole('menuitem', { name: 'Insert 1 row above' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(menu).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
    await expect(shown(page, 'A3')).toHaveText('Bolts');
    // The formulas moved with their rows.
    await expect(shown(page, 'D3')).toHaveText('6');
  });

  test('Escape closes the cell menu and the focus goes back to the grid', async ({ page }) => {
    await page.keyboard.press('Shift+F10');
    await expect(page.getByRole('menu')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('menu')).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
  });

  test('Ctrl+D fills down, as dragging the fill handle would', async ({ page }) => {
    await goTo(page, 'C2:C3');
    await page.keyboard.press('Control+d');
    await expect(shown(page, 'C3')).toHaveText('3');
    await expect(shown(page, 'D3')).toHaveText('12');
  });

  test('sets a column width from the menu', async ({ page }) => {
    await goTo(page, 'B2');
    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Column width…' }).focus();
    await page.keyboard.press('Enter');
    const dialog = page.getByRole('dialog', { name: 'Width of column B' });
    const input = dialog.getByRole('textbox');
    await expect(input).toBeFocused();
    await page.keyboard.type('25');
    await page.keyboard.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(grid(page)).toHaveCount(1);

    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Column width…' }).focus();
    await page.keyboard.press('Enter');
    await expect(input).toHaveValue('25');
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });

  test('Ctrl+/ lists the shortcuts and Escape gives the focus back', async ({ page }) => {
    await page.keyboard.press('Control+/');
    const dialog = page.getByRole('dialog', { name: 'Keyboard shortcuts' });
    await expect(dialog).toBeVisible();
    await expect(dialog.getByText('Shift+F10, Menu key')).toBeVisible();
    // Tab stays inside the dialog.
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await expect(dialog.locator(':focus')).toHaveCount(1);
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
  });

  test('switches sheets from the sheet tabs', async ({ page }) => {
    await page.keyboard.press('F6');
    await page.keyboard.press('Tab');
    await expect(page.getByRole('button', { name: 'Summary' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('button', { name: 'Summary' })).toHaveAttribute('aria-current', 'true');
  });
});

test('picks from a dropdown list with the keyboard', async ({ page }) => {
  await openSample(page, 'Price list');
  await goTo(page, 'C3');
  await page.keyboard.press('Enter');
  const box = page.getByRole('combobox', { name: 'Filter values' });
  await expect(box).toBeFocused();
  await page.keyboard.type('serv');
  await expect(page.getByRole('option', { name: 'Services' })).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Enter');
  await expect(shown(page, 'C3')).toHaveText('Services');
});

test('picks a cell for a schema from the keyboard', async ({ page }) => {
  await openSample(page, 'Price list');
  await page.getByRole('button', { name: 'Schemas' }).focus();
  await page.keyboard.press('Enter');
  const designer = page.getByRole('dialog', { name: 'Schemas' });
  await designer.getByRole('button', { name: /Single cell/ }).focus();
  await page.keyboard.press('Enter');

  // The sheet takes the focus for picking; F6 reaches the bar.
  await expect(page.locator('canvas[data-testid="data-grid-canvas"]:focus-within')).toHaveCount(1);
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('F6');
  const add = page.getByRole('dialog', { name: 'Pick cells for the schema' }).getByRole('button', { name: 'Add cell B2' });
  await expect(add).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(designer).toBeVisible();
  await expect(designer.getByRole('textbox', { name: 'Cell', exact: true }).last()).toHaveValue('B2');
});
