import { expect, test, type Page } from '@playwright/test';
import type { MenuOptions } from '@ruled/sheet';
import { goTo, serverUrl } from './helpers';

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
