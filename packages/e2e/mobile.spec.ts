import AxeBuilder from '@axe-core/playwright';
import { devices, expect, test, type Locator, type Page } from '@playwright/test';
import { goTo, grid as gridSurface, openFile, openSample, ordersWorkbook, serverUrl, shown } from './helpers';

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
  // The toolbar's actions are one row that fits: nothing scrolls out of it, Save included.
  const actions = page.locator('.wbe-toolbar-actions');
  expect((await actions.boundingBox())!.height).toBeLessThan(60);
  expect(await actions.evaluate((el) => el.scrollWidth - el.clientWidth)).toBeLessThanOrEqual(0);
  await expect(page.getByRole('button', { name: 'Save as Excel' })).toBeInViewport({ ratio: 1 });
});

test('the less used controls are behind "More", and the format bar shows when asked for', async ({ page }) => {
  await openSample(page, 'Price list');
  await expect(page.getByRole('group', { name: 'Format', exact: true })).toBeHidden();
  const more = page.getByRole('button', { name: 'More' });
  await more.tap();
  await expect(more).toHaveAttribute('aria-expanded', 'true');
  for (const name of ['Schemas', 'Samples', 'Keyboard shortcuts']) await expect(page.getByRole('button', { name, exact: true })).toBeInViewport({ ratio: 1 });
  await page.getByRole('checkbox', { name: 'Formatting' }).tap();
  await expect(page.getByRole('group', { name: 'Format', exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(more).toHaveAttribute('aria-expanded', 'false');
  await expect(more).toBeFocused();
  await expect(page.getByRole('button', { name: 'Schemas', exact: true })).toBeHidden();
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
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B11');
  // The problem shows under the formula bar, as there's no hovering on a phone.
  await expect(page.locator('.wbe-cell-note')).toContainText('Required.');
});

test('a tap selects a cell, and the formula bar edits it', async ({ page }) => {
  await openSample(page, 'Price list');
  const grid = (await gridSurface(page).first().boundingBox())!;
  // Row 3, column B: past the row numbers and the headings.
  await page.touchscreen.tap(grid.x + 120, grid.y + 24 + 34 * 2 + 10);
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).not.toHaveValue('A1');
  const name = await page.getByRole('combobox', { name: 'Go to cell' }).inputValue();
  const input = page.getByRole('combobox', { name: 'Cell contents' });
  await input.tap();
  await input.fill('Bracket, edited on a phone');
  await input.press('Enter');
  await goTo(page, name);
  await expect(input).toHaveValue('Bracket, edited on a phone');
});

/** A finger dragged from one point to another: Playwright taps but doesn't drag, so the touches go through Chromium's DevTools protocol. */
async function touchDrag(page: Page, from: { x: number; y: number }, to: { x: number; y: number }): Promise<void> {
  const cdp = await page.context().newCDPSession(page);
  const at = (t: number) => [{ x: from.x + (to.x - from.x) * t, y: from.y + (to.y - from.y) * t }];
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: at(0) });
  for (let i = 1; i <= 10; i++) await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: at(i / 10) });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await cdp.detach();
}

async function centre(locator: Locator): Promise<{ x: number; y: number }> {
  const b = (await locator.boundingBox())!;
  return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
}

/** Opens the orders workbook, touches its grid, and selects B2, whose handles then show. */
async function touchedOrders(page: Page): Promise<void> {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  const grid = (await gridSurface(page).first().boundingBox())!;
  await page.touchscreen.tap(grid.x + 120, grid.y + 60);
  await goTo(page, 'B2');
}

test("a finger drags the selection's corner to select a range", async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', "touches are sent through Chromium's DevTools protocol");
  await touchedOrders(page);
  const handles = page.locator('.wbe-touch-handle');
  await expect(handles).toHaveCount(2);
  const [corner, far] = [await centre(handles.first()), await centre(handles.last())];
  const height = far.y - corner.y;
  // From B2's bottom-right corner into column C, two rows down.
  await touchDrag(page, far, { x: far.x + 10, y: far.y + height * 1.5 });
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B2:C4');
});

test('filling by dragging, from the cell menu, continues the selection down', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', "touches are sent through Chromium's DevTools protocol");
  await touchedOrders(page);
  const height = (await centre(page.locator('.wbe-touch-handle').last())).y - (await centre(page.locator('.wbe-touch-handle').first())).y;
  await page.getByRole('button', { name: 'Cell menu' }).tap();
  await page.getByRole('menuitem', { name: 'Fill by dragging' }).tap();
  const arrows = page.locator('.wbe-touch-fill');
  await expect(arrows).toHaveCount(2);
  const down = await centre(arrows.first());
  await touchDrag(page, down, { x: down.x, y: down.y + height * 1.5 });
  await expect(shown(page, 'B3')).toHaveText('2');
  await expect(shown(page, 'B4')).toHaveText('2');
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B2:B4');
  await expect(arrows).toHaveCount(0);
});

test('the cell menu opens from the formula bar, without right-clicking', async ({ page }) => {
  await openSample(page, 'Price list');
  await goTo(page, 'A3');
  await page.getByRole('button', { name: 'Cell menu' }).tap();
  const menu = page.getByRole('menu');
  await expect(menu).toBeInViewport();
  await menu.getByRole('menuitem', { name: 'Rows' }).tap();
  await page.getByRole('menu', { name: 'Rows' }).getByRole('menuitem', { name: 'Insert 1 row above' }).tap();
  await goTo(page, 'A4');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('SKU-1002');
});

test('dialogs fit the screen', async ({ page }) => {
  await openSample(page, 'Price list');
  await page.getByRole('button', { name: 'More' }).tap();
  await page.getByRole('button', { name: 'Keyboard shortcuts' }).tap();
  const dialog = page.getByRole('dialog', { name: 'Keyboard shortcuts' });
  const box = (await dialog.boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(page.viewportSize()!.width);
  // Taller than the screen, it scrolls within it, its last line within reach.
  expect(box.y + box.height).toBeLessThanOrEqual(page.viewportSize()!.height);
  await dialog.evaluate((el) => (el.scrollTop = el.scrollHeight));
  await expect(dialog.getByRole('row').last()).toBeInViewport();
  await page.getByRole('button', { name: 'Close' }).last().tap();
  await page.getByRole('button', { name: 'More' }).tap();
  await page.getByRole('button', { name: 'Schemas' }).tap();
  const designer = page.getByRole('dialog', { name: 'Schemas' });
  await expect(designer.getByRole('button', { name: /Apply/ })).toBeInViewport();
});

test('the portal works on a phone', async ({ page }) => {
  await page.goto(serverUrl('portal'));
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

test('holding a finger on a cell opens the cell menu', async ({ page }) => {
  test.skip(process.env.E2E_GRID !== 'ruled', "Glide's grid opens no menu for a long press");
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  const box = (await gridSurface(page).first().boundingBox())!;
  // B3: 40 pixel row numbers, then 64 pixel columns, under a 24 pixel heading.
  const at = { x: Math.round(box.x + 40 + 64 + 32), y: Math.round(box.y + 24 + 2 * 20 + 10) };
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [at] });
  await page.waitForTimeout(800);
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await expect(page.getByRole('menu')).toBeVisible();
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('B3');
});
