import { expect, test, type Page } from '@playwright/test';
import type { MenuOptions } from '@ruled/sheet';
import { goTo, grid, serverUrl } from './helpers';

/** The <workbook-editor> element's page, with the menus kept as given. */
async function open(page: Page, menus: MenuOptions): Promise<void> {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await setMenus(page, menus);
}

const setMenus = (page: Page, menus: MenuOptions) =>
  page.evaluate((m) => (document.getElementById('editor') as HTMLElement & { setMenus(m: MenuOptions): void }).setMenus(m), menus);

const menu = (page: Page) => page.locator('.wbe-context-menu').first();

/** The open menu's entries, separators as `-`. */
const entries = (page: Page) =>
  menu(page)
    .locator(':scope > [role="menuitem"], :scope > [role="separator"]')
    .evaluateAll((els) => els.map((el) => (el.getAttribute('role') === 'separator' ? '-' : (el.querySelector('span')?.textContent ?? ''))));

test('keeps the cell menu entries listed, in English and German', async ({ page }) => {
  await open(page, { cell: ['menu.copy', 'menu.rows', 'menu.find'], header: [] });
  await goTo(page, 'B2');
  await page.keyboard.press('Shift+F10');
  await expect.poll(() => entries(page)).toEqual(['Copy', '-', 'Rows', '-', 'Find and replace…']);
  await page.keyboard.press('Escape');

  await page.evaluate(() => (document.getElementById('editor') as HTMLElement & { setLanguage(l: string): void }).setLanguage('de'));
  await page.keyboard.press('Shift+F10');
  await expect.poll(() => entries(page)).toEqual(['Kopieren', '-', 'Zeilen', '-', 'Suchen und Ersetzen…']);
});

test('turns an entry’s shortcut off with it, unless kept by id', async ({ page }) => {
  await open(page, { cell: [], header: [] });
  await goTo(page, 'B2');
  // No menu to open.
  await page.keyboard.press('Shift+F10');
  await expect(menu(page)).toHaveCount(0);
  await page.keyboard.press('ControlOrMeta+f');
  await expect(page.getByRole('search')).toHaveCount(0);

  await setMenus(page, { cell: [], header: [], shortcuts: ['menu.find'] });
  await goTo(page, 'B2');
  await page.keyboard.press('ControlOrMeta+f');
  await expect(page.getByRole('search')).toBeVisible();
});

test('opens the cell and heading menus by right-click, as soon as the page opens', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  const sheet = grid(page).first();
  const nameBox = page.getByRole('combobox', { name: 'Go to cell' });
  // Rows are 20 pixels under a 24 pixel heading, after 40 pixels of row numbers; column B starts within 150 pixels.
  await sheet.click({ button: 'right', position: { x: 150, y: 74 } });
  await expect.poll(() => entries(page)).toEqual(expect.arrayContaining(['Cut', 'Copy', 'Paste', 'New note', 'Insert cells…', 'Rows', 'Columns', 'Find and replace…']));
  await expect(nameBox).toHaveValue('B3');
  await page.keyboard.press('Escape');
  await expect(menu(page)).toHaveCount(0);

  await sheet.click({ button: 'right', position: { x: 20, y: 74 } });
  await expect.poll(() => entries(page)).toEqual(expect.arrayContaining(['Copy', 'Insert 1 row above', 'Insert 1 row below', 'Delete 1 row', 'Hide row']));
  expect(await entries(page)).not.toContain('Insert 1 column left');
  await page.keyboard.press('Escape');

  await sheet.click({ button: 'right', position: { x: 150, y: 12 } });
  await expect.poll(() => entries(page)).toEqual(expect.arrayContaining(['Copy', 'Insert 1 column left', 'Insert 1 column right', 'Delete 1 column', 'Hide column']));
  expect(await entries(page)).not.toContain('Insert 1 row above');
  await expect(nameBox).toHaveValue('B1');
});
