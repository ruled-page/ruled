import { expect, test, type Page } from '@playwright/test';
import { goTo, grid, openFile, ordersWorkbook, shown } from './helpers';

/**
 * Lines!A1:D3 of the orders workbook: Item, Qty, Price, Total, with Total a
 * formula (B*C) in rows 2 and 3. A selection is checked by what Delete clears.
 */

const nameBox = (page: Page) => page.getByRole('combobox', { name: 'Go to cell' });

/** Where to press: columns are 64 pixels after the 40 pixel row numbers, rows 20 pixels under the 24 pixel letters. */
async function at(page: Page) {
  const box = (await grid(page).first().boundingBox())!;
  return {
    letter: (col: number) => ({ x: box.x + 40 + 64 * col + 32, y: box.y + 12 }),
    number: (row: number) => ({ x: box.x + 20, y: box.y + 24 + 20 * (row - 1) + 10 }),
    corner: { x: box.x + 20, y: box.y + 12 },
  };
}

async function press(page: Page, p: { x: number; y: number }, modifier?: 'Shift' | 'ControlOrMeta') {
  if (modifier) await page.keyboard.down(modifier);
  await page.mouse.click(p.x, p.y);
  if (modifier) await page.keyboard.up(modifier);
}

async function drag(page: Page, from: { x: number; y: number }, to: { x: number; y: number }) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
}

/** Clears the selection, agreeing to lose formulas when asked, checks what the cells then show, and undoes it. */
async function clearShows(page: Page, expected: Record<string, string>): Promise<void> {
  await page.keyboard.press('Delete');
  const ask = page.getByRole('alertdialog');
  if (await ask.isVisible().catch(() => false)) await ask.getByRole('button', { name: /^Replace formula/ }).click();
  await expect(page.getByRole('button', { name: /change/ })).toBeVisible();
  for (const [ref, text] of Object.entries(expected)) await expect(shown(page, ref), ref).toHaveText(text);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.getByRole('button', { name: /change/ })).toHaveCount(0);
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await goTo(page, 'A1');
});

test('selects whole columns by their letters: a click, a drag, Shift+click and Ctrl+click', async ({ page }) => {
  const p = await at(page);
  await press(page, p.letter(2));
  await expect(nameBox(page)).toHaveValue('C1');
  await clearShows(page, { B2: '2', C2: '', D2: '0' });

  await drag(page, p.letter(1), p.letter(2));
  await clearShows(page, { A2: 'Bolts', B2: '', C2: '', D2: '0' });

  await press(page, p.letter(0));
  await press(page, p.letter(2), 'ControlOrMeta');
  await clearShows(page, { A2: '', B2: '2', C2: '' });

  await press(page, p.letter(1));
  await press(page, p.letter(2), 'Shift');
  await clearShows(page, { A2: 'Bolts', B2: '', C2: '' });
});

test('selects whole rows by their numbers: a click, a drag, Shift+click and Ctrl+click', async ({ page }) => {
  const p = await at(page);
  await press(page, p.number(2));
  await expect(nameBox(page)).toHaveValue('A2');
  await clearShows(page, { A2: '', D2: '', A3: 'Nuts' });

  await drag(page, p.number(2), p.number(3));
  await clearShows(page, { A1: 'Item', A2: '', A3: '' });

  await press(page, p.number(1));
  await press(page, p.number(3), 'ControlOrMeta');
  await clearShows(page, { A1: '', A2: 'Bolts', A3: '' });

  await press(page, p.number(1));
  await press(page, p.number(2), 'Shift');
  await clearShows(page, { A1: '', A2: '', A3: 'Nuts' });
});

test('selects the whole sheet from the corner or Ctrl+A, and whole lines with Shift+Space and Ctrl+Space', async ({ page }) => {
  const p = await at(page);
  await press(page, p.corner);
  await clearShows(page, { A1: '', D3: '' });

  await goTo(page, 'B2');
  await page.keyboard.press('Shift+Space');
  await clearShows(page, { A2: '', C2: '', A3: 'Nuts' });

  await goTo(page, 'C3');
  await page.keyboard.press('Control+Space');
  await clearShows(page, { C2: '', C3: '', B3: '4' });

  await goTo(page, 'C3');
  await page.keyboard.press('ControlOrMeta+a');
  // The active cell stays where it was.
  await expect(nameBox(page)).toHaveValue('C3');
  await clearShows(page, { A1: '', D3: '' });
});
