import AxeBuilder from '@axe-core/playwright';
import { devices, expect, test, type Page } from '@playwright/test';
import { goTo, openSample } from './helpers';

/** A phone, with touch: the editor, the demo and the portal at that size. */
test.use({ ...devices['Pixel 7'] });

/** No sideways scrolling of the page itself. */
async function fitsWidth(page: Page): Promise<void> {
  expect(await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth)).toBeLessThanOrEqual(0);
}

test('the start screen and an open workbook fit the screen', async ({ page }) => {
  await page.goto('/');
  await fitsWidth(page);
  await openSample(page, 'Price list');
  await fitsWidth(page);
  // The toolbar is one row of actions that scrolls sideways, not several rows.
  const toolbar = await page.locator('.toolbar-actions').boundingBox();
  expect(toolbar!.height).toBeLessThan(60);
});

test('the checks panel waits to be opened, and going to a problem shows its cell', async ({ page }) => {
  await openSample(page, 'Price list');
  await expect(page.getByRole('tabpanel')).toHaveCount(0);
  await page.getByRole('button', { name: /^Price list/ }).tap();
  const panel = page.getByRole('tabpanel');
  await expect(panel.getByText('Required.')).toBeVisible();
  // The toolbar stays usable above the panel.
  await expect(page.getByRole('button', { name: 'Undo' })).toBeInViewport();
  await panel.getByText('Required.').tap();
  await expect(page.getByRole('tabpanel')).toHaveCount(0);
  await expect(page.getByRole('textbox', { name: 'Go to cell' })).toHaveValue('B11');
  // The problem shows under the formula bar, as there's no hovering on a phone.
  await expect(page.locator('.cell-note')).toContainText('Required.');
});

test('a tap selects a cell, and the formula bar edits it', async ({ page }) => {
  await openSample(page, 'Price list');
  const grid = (await page.locator('canvas[data-testid="data-grid-canvas"]').first().boundingBox())!;
  // Row 3, column B: past the row numbers and the headings.
  await page.touchscreen.tap(grid.x + 120, grid.y + 24 + 34 * 2 + 10);
  await expect(page.getByRole('textbox', { name: 'Go to cell' })).not.toHaveValue('A1');
  const name = await page.getByRole('textbox', { name: 'Go to cell' }).inputValue();
  const input = page.getByRole('textbox', { name: 'Cell contents' });
  await input.tap();
  await input.fill('Bracket, edited on a phone');
  await input.press('Enter');
  await goTo(page, name);
  await expect(input).toHaveValue('Bracket, edited on a phone');
});

test('the cell menu opens from the formula bar, without right-clicking', async ({ page }) => {
  await openSample(page, 'Price list');
  await goTo(page, 'A3');
  await page.getByRole('button', { name: 'Cell menu' }).tap();
  const menu = page.getByRole('menu');
  await expect(menu).toBeInViewport();
  await menu.getByRole('menuitem', { name: 'Insert 1 row above' }).tap();
  await goTo(page, 'A4');
  await expect(page.getByRole('textbox', { name: 'Cell contents' })).toHaveValue('SKU-1002');
});

test('dialogs fit the screen', async ({ page }) => {
  await openSample(page, 'Price list');
  await page.getByRole('button', { name: 'Keyboard shortcuts' }).tap();
  const dialog = page.getByRole('dialog', { name: 'Keyboard shortcuts' });
  const box = (await dialog.boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(page.viewportSize()!.width);
  await page.getByRole('button', { name: 'Close' }).last().tap();
  await page.getByRole('button', { name: 'Schemas' }).tap();
  const designer = page.getByRole('dialog', { name: 'Schemas' });
  await expect(designer.getByRole('button', { name: /Apply/ })).toBeInViewport();
});

test('the portal works on a phone', async ({ page }) => {
  await page.goto('/embedded.html');
  await fitsWidth(page);
  await page.getByRole('button', { name: 'Try a sample with problems' }).tap();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await fitsWidth(page);
  await expect(page.getByRole('button', { name: /Submit price list/ })).toBeVisible();
});

test('meets the accessibility checks on a phone', async ({ page }) => {
  await openSample(page, 'Price list');
  await page.getByRole('button', { name: /^Price list/ }).tap();
  await page.waitForTimeout(300);
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  expect(violations.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(' ')).join(', ')}`)).toEqual([]);
});
