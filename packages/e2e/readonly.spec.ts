import { expect, test, type Page } from '@playwright/test';
import { goTo, grid, serverUrl, shown } from './helpers';

type Editor = HTMLElement & { readOnly: boolean; editedCount: number; unsaved: boolean };

/** The <workbook-editor> element's page, its sample open and made read-only by its attribute. */
async function open(page: Page): Promise<void> {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await page.evaluate(() => document.getElementById('editor')!.setAttribute('read-only', ''));
}

const edited = (page: Page) => page.evaluate(() => ({ count: (document.getElementById('editor') as Editor).editedCount, unsaved: (document.getElementById('editor') as Editor).unsaved }));

const menu = (page: Page) => page.locator('.wbe-context-menu').first();

const entries = (page: Page) =>
  menu(page)
    .locator(':scope > [role="menuitem"], :scope > [role="separator"]')
    .evaluateAll((els) => els.map((el) => (el.getAttribute('role') === 'separator' ? '-' : (el.querySelector('span')?.textContent ?? ''))));

test('keys that edit change nothing, and say so only to screen readers', async ({ page }) => {
  await open(page);
  await goTo(page, 'A2');
  const before = await shown(page, 'A2').getAttribute('aria-label');
  // Typing starts no edit; Delete, F2, paste, cut, fill and undo change nothing.
  await page.keyboard.type('xyz');
  await expect(page.locator('.wbe-cell-editor')).toHaveCount(0);
  for (const key of ['Delete', 'Backspace', 'F2', 'ControlOrMeta+x', 'ControlOrMeta+v', 'ControlOrMeta+d', 'ControlOrMeta+r', 'ControlOrMeta+z', 'ControlOrMeta+y', 'ControlOrMeta+b', 'ControlOrMeta+Shift+=', 'ControlOrMeta+-']) {
    await page.keyboard.press(key);
  }
  await expect(page.locator('.wbe-cell-editor')).toHaveCount(0);
  await expect(page.locator('.wbe-toast')).toHaveCount(0);
  await expect(shown(page, 'A2')).toHaveAttribute('aria-label', before!);
  await expect(page.getByText('This workbook is read-only.')).toHaveCount(1);
  expect(await edited(page)).toEqual({ count: 0, unsaved: false });
});

test('the cell menu keeps only what does not edit, and the toolbar its viewing parts', async ({ page }) => {
  await open(page);
  await goTo(page, 'B2');
  await page.keyboard.press('Shift+F10');
  await expect.poll(() => entries(page)).toEqual(['Copy', '-', 'Filter by values…', 'Add filter buttons', '-', 'Find and replace…']);
  // Filtering stays, for viewing only: nothing is changed or marked unsaved.
  await menu(page).getByRole('menuitem', { name: 'Add filter buttons' }).click();
  await page.keyboard.press('Shift+F10');
  await expect.poll(() => entries(page)).toContain('Remove filter buttons');
  await page.keyboard.press('Escape');
  expect(await edited(page)).toEqual({ count: 0, unsaved: false });
  await expect(page.locator('.wbe-toast')).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Undo' })).toHaveCount(0);
  await expect(page.getByRole('toolbar', { name: 'Formatting' })).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Add sheet' })).toHaveCount(0);
  await expect(grid(page)).toBeVisible();

  // Taking the attribute away makes it editable again.
  await page.evaluate(() => document.getElementById('editor')!.removeAttribute('read-only'));
  await page.keyboard.press('Shift+F10');
  await expect.poll(async () => (await entries(page)).length).toBeGreaterThan(10);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('button', { name: 'Undo' })).toHaveCount(1);
});

test('the readOnly property and option turn it on and off', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  expect(await page.evaluate(() => (document.getElementById('editor') as Editor).readOnly)).toBe(false);
  await page.evaluate(() => ((document.getElementById('editor') as Editor).readOnly = true));
  await expect(page.getByRole('button', { name: 'Undo' })).toHaveCount(0);
  await page.evaluate(() => {
    const el = document.getElementById('editor') as Editor & { options: object };
    el.options = { ...el.options, readOnly: false };
  });
  await expect(page.getByRole('button', { name: 'Undo' })).toHaveCount(1);
});
