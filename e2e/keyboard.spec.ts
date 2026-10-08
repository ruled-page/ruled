import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, openSample, ordersWorkbook, shown } from './helpers';

/** Everything here is done without the mouse. */

const grid = (page: Page) => page.locator('canvas[data-testid="data-grid-canvas"]:focus-within');
const nameBox = (page: Page) => page.getByRole('combobox', { name: 'Go to cell' });

test.describe('with only a keyboard', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await goTo(page, 'A1');
  });

  test('F6 moves between the parts of the editor and back to the sheet', async ({ page }) => {
    await page.keyboard.press('F6');
    await expect(page.getByRole('tab', { name: 'Lines' })).toBeFocused();
    await page.keyboard.press('F6');
    await expect(page.locator('.wbe-toolbar :focus')).toHaveCount(1);
    await page.keyboard.press('F6');
    await expect(page.getByRole('combobox', { name: 'Font', exact: true })).toBeFocused();
    await page.keyboard.press('F6');
    await expect(nameBox(page)).toBeFocused();
    await page.keyboard.press('F6');
    await expect(grid(page)).toHaveCount(1);
    await page.keyboard.press('Shift+F6');
    await expect(nameBox(page)).toBeFocused();
  });

  test('Shift+F6 reaches the name box, which goes to a cell', async ({ page }) => {
    await page.keyboard.press('Shift+F6');
    await expect(nameBox(page)).toBeFocused();
    await page.keyboard.type('C3');
    await page.keyboard.press('Enter');
    await expect(grid(page)).toHaveCount(1);
    await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('0.5');
  });

  test('fills down from the cell menu, as dragging the fill handle would', async ({ page }) => {
    await goTo(page, 'C2:C3');
    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Fill down' }).focus();
    await page.keyboard.press('Enter');
    await expect(shown(page, 'C3')).toHaveText('3');
    await expect(shown(page, 'D3')).toHaveText('12');
  });

  test('leaves the browser its own shortcuts: reload, bookmark, find next, history', async ({ page }) => {
    // Runs after the editor's own listener, so it sees whether the editor took the key.
    await page.evaluate(() => {
      (window as unknown as { taken: string[] }).taken = [];
      window.addEventListener('keydown', (e) => {
        if (e.defaultPrevented) (window as unknown as { taken: string[] }).taken.push(`${e.ctrlKey ? 'Ctrl+' : ''}${e.metaKey ? 'Meta+' : ''}${e.key}`);
      });
    });
    for (const key of ['Control+r', 'Meta+r', 'Control+d', 'Meta+d', 'Control+g', 'Meta+g', 'Meta+y']) await page.keyboard.press(key);
    expect(await page.evaluate(() => (window as unknown as { taken: string[] }).taken)).toEqual([]);
    await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('Item');
  });

  test('⌘G finds the next match, and ⇧⌘G the one before, while finding', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.keyboard.type('s');
    await page.keyboard.press('Enter');
    await expect(nameBox(page)).toHaveValue('A2');
    await page.keyboard.press('Control+g');
    await expect(nameBox(page)).toHaveValue('A3');
    await page.keyboard.press('Control+Shift+g');
    await expect(nameBox(page)).toHaveValue('A2');
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
    // Cut, Copy, Paste, Paste special, Fill down, Fill right, Fill series, New note, New comment, then the rows' submenu.
    for (let i = 0; i < 9; i++) await page.keyboard.press('ArrowDown');
    const rows = page.getByRole('menuitem', { name: 'Rows' });
    await expect(rows).toBeFocused();
    await page.keyboard.press('ArrowRight');
    const sub = page.getByRole('menu', { name: 'Rows' });
    await expect(sub.getByRole('menuitem', { name: 'Insert 1 row above' })).toBeFocused();
    await expect(rows).toHaveAttribute('aria-expanded', 'true');
    // Left closes it, back on the item that opened it; Enter opens it again.
    await page.keyboard.press('ArrowLeft');
    await expect(sub).toBeHidden();
    await expect(rows).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(sub.getByRole('menuitem', { name: 'Insert 1 row above' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('menu')).toHaveCount(0);
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

  test('sets a column width from the menu', async ({ page }) => {
    await goTo(page, 'B2');
    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Columns' }).focus();
    await page.keyboard.press('Enter');
    await page.getByRole('menuitem', { name: 'Column width…' }).focus();
    await page.keyboard.press('Enter');
    const dialog = page.getByRole('dialog', { name: 'Width of column B' });
    const input = dialog.getByRole('textbox');
    await expect(input).toBeFocused();
    await page.keyboard.type('25');
    await page.keyboard.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
    // Said once the width is set.
    await expect(page.getByRole('status').filter({ hasText: 'Column width' })).toHaveCount(1);

    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Columns' }).focus();
    await page.keyboard.press('Enter');
    await page.getByRole('menuitem', { name: 'Column width…' }).focus();
    await page.keyboard.press('Enter');
    await expect(input).toHaveValue('25');
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });

  test('⌘/ lists the shortcuts too, as do keyboards that type / with Shift', async ({ page }) => {
    await page.keyboard.press('Meta+/');
    await expect(page.getByRole('dialog', { name: 'Keyboard shortcuts' })).toBeVisible();
    await page.keyboard.press('Escape');
    // A German keyboard: / is Shift+7, but the key in the US / place still counts.
    await page.locator('canvas[data-testid="data-grid-canvas"]').first().dispatchEvent('keydown', { key: '-', code: 'Slash', metaKey: true, bubbles: true });
    await expect(page.getByRole('dialog', { name: 'Keyboard shortcuts' })).toBeVisible();
  });

  test('Ctrl+/ lists the shortcuts and Escape gives the focus back', async ({ page }) => {
    await page.keyboard.press('Control+/');
    const dialog = page.getByRole('dialog', { name: 'Keyboard shortcuts' });
    await expect(dialog).toBeVisible();
    // On the grid and on the sheet tabs.
    await expect(dialog.getByText('Shift+F10, Menu key')).toHaveCount(2);
    // Tab stays inside the dialog.
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await expect(dialog.locator(':focus')).toHaveCount(1);
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
  });

  test('switches sheets from the sheet tabs with the arrow keys', async ({ page }) => {
    await page.keyboard.press('F6');
    await expect(page.getByRole('tab', { name: 'Lines' })).toBeFocused();
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('tab', { name: 'Summary' })).toBeFocused();
    // Moving the focus doesn't switch sheet; Enter does.
    await expect(page.getByRole('tab', { name: 'Lines' })).toHaveAttribute('aria-selected', 'true');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('tab', { name: 'Lines' })).toBeFocused();
    await page.keyboard.press('End');
    await expect(page.getByRole('tab', { name: 'Summary' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('tab', { name: 'Summary' })).toHaveAttribute('aria-selected', 'true');
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
  await expect(designer.getByRole('textbox', { name: /^Cell of / }).last()).toHaveValue('B2');
});

test('screen readers hear cells by their reference, rows as the sheet numbers them, and what was done', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  // Inside the canvas, which Playwright takes as hidden.
  await expect(page.locator('table[role="grid"]')).toHaveAttribute('aria-label', 'Cells of Lines');
  const d2 = shown(page, 'D2');
  // The cell's text is its value; it is named by where it is and what it holds.
  await expect(d2).toHaveText('6');
  await expect(d2).toHaveAttribute('aria-label', 'D2, 6, formula');
  await expect(d2.locator('xpath=..')).toHaveAttribute('aria-rowindex', '2');
  await expect(d2).toHaveAttribute('aria-colindex', '4');

  // Enter in the formula bar gives the keyboard back to the grid.
  await goTo(page, 'B2');
  const input = page.getByRole('combobox', { name: 'Cell contents' });
  await input.fill('5');
  await input.press('Enter');
  await expect(grid(page)).toHaveCount(1);
  await expect(shown(page, 'D2')).toHaveText('15');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.getByRole('status').filter({ hasText: 'Undone: Edit' })).toHaveCount(1);
  await expect(shown(page, 'D2')).toHaveText('6');
});
