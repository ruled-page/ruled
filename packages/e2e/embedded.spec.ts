import AxeBuilder from '@axe-core/playwright';
import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { goTo, readWorkbook, serverUrl, setCell } from './helpers';

/**
 * The editor embedded in a host app (embedded.html): a procurement portal
 * where suppliers check and submit price lists, and the portal reads the
 * values into its own view.
 */

/** A cell's contents, read from the formula bar (the template freezes its heading row, so the grid is two panes). */
async function contents(page: Page, ref: string): Promise<string> {
  await goTo(page, ref);
  return page.getByRole('combobox', { name: 'Cell contents' }).inputValue();
}

async function openSample(page: Page): Promise<void> {
  await page.goto(serverUrl('portal'));
  await page.getByRole('button', { name: 'Try a sample with problems' }).click();
  await expect(page.getByRole('heading', { name: 'Check your price list' })).toBeVisible();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
}

const chip = (page: Page) => page.getByRole('button', { name: /^Supplier price list/ });
const panel = (page: Page) => page.getByRole('tabpanel');

/** The sample's five errors, fixed. */
async function fixErrors(page: Page): Promise<void> {
  await setCell(page, 'F7', '10');
  await setCell(page, 'C8', 'Hardware');
  await setCell(page, 'E9', 'USD');
  await setCell(page, 'A10', 'SKU-1009');
  await setCell(page, 'D11', '1250');
}

test('shows every kind of check: types, the workbook’s Excel rules, and the portal’s business checks', async ({ page }) => {
  await openSample(page);
  await expect(chip(page)).toContainText('5 errors');
  // Types and limits.
  await expect(panel(page).getByText('At most 2 decimal places.')).toBeVisible();
  await expect(panel(page).getByText('Must be one of: Hardware, Software, Services.')).toBeVisible();
  // Business checks that call the portal's services, and a rule across the table.
  await expect(panel(page).getByText('SKU-1099 isn’t in our catalogue. Check the code, or ask us to add the product first.')).toBeVisible();
  await expect(panel(page).getByText('+30% on the last approved price of $7.25. We’ll check this before approving.')).toBeVisible();
  await expect(panel(page).getByText('Use one currency for the whole list (USD, as in the first row).')).toBeVisible();

  // Excel's own rule for the column turns a value away, with the workbook's message.
  await setCell(page, 'F6', '0');
  await expect(page.getByText('Minimum order: Enter a whole number from 1 to 10,000.')).toBeVisible();
  expect(await contents(page, 'F6')).toBe('1');
});

test('fills in a known product from the catalogue as it is typed', async ({ page }) => {
  await openSample(page);
  await setCell(page, 'A13', 'SKU-1012');
  await expect.poll(() => contents(page, 'B13')).toBe('Barcode scanner');
  expect(await contents(page, 'C13')).toBe('Hardware');
});

test('a supplier fixes the list, downloads a copy, submits it, and the portal shows and approves it', async ({ page }) => {
  await openSample(page);
  await fixErrors(page);
  // Price changes are only warnings, so the list can go in with them. The new code in
  // A10 brings its own approved price, so its price is checked again.
  await expect(chip(page)).toContainText('2 warnings');
  await expect(panel(page).getByText('+529% on the last approved price of $8.75. We’ll check this before approving.')).toBeVisible();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download a copy' }).click();
  const copy = await readWorkbook(await readFile(await (await download).path()));
  expect(copy.getWorksheet('Price list')!.getCell('A10').value).toBe('SKU-1009');

  await page.getByRole('button', { name: /Submit price list/ }).click();
  await expect(page.getByRole('heading', { name: 'supplier-price-list.xlsx' })).toBeVisible();
  await expect(page.getByText('Waiting for approval')).toBeVisible();

  // The portal's own view, from the values the editor handed over.
  const lines = page.getByRole('region', { name: 'Products' });
  const row = (sku: string) => lines.getByRole('row').filter({ has: page.getByRole('rowheader', { name: sku }) });
  await expect(row('SKU-1010')).toContainText('$1,250.00');
  await expect(row('SKU-1002')).toContainText('+29.7%');
  await expect(row('SKU-1006')).toContainText('10');
  await expect(page.getByText('prices moved more than 15%')).toBeVisible();

  await page.getByRole('button', { name: 'Approve prices' }).click();
  await expect(page.getByText('Approved. These prices are now the ones the next price list is compared with.')).toBeVisible();

  // Approved prices are what the next list is checked against: the big change is now the approved price.
  await page.getByRole('link', { name: 'Price lists' }).first().click();
  await expect(page.getByRole('row', { name: /supplier-price-list\.xlsx/ })).toContainText('Approved');
  await page.getByRole('button', { name: 'Try a sample with problems' }).click();
  await expect(chip(page)).toContainText('5 errors');
  await expect(panel(page).getByText(/on the last approved price of \$7\.25/)).toHaveCount(0);
});

test('edits a submitted list again and resubmits it in place', async ({ page }) => {
  await openSample(page);
  await fixErrors(page);
  await page.getByRole('button', { name: /Submit price list/ }).click();
  await expect(page.getByRole('heading', { name: 'supplier-price-list.xlsx' })).toBeVisible();
  const url = page.url();

  await page.getByRole('button', { name: 'Edit again' }).click();
  await expect(page.getByRole('heading', { name: 'Change your price list' })).toBeVisible();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  expect(await contents(page, 'A10')).toBe('SKU-1009');
  await setCell(page, 'D2', '4.6');
  await page.getByRole('button', { name: /Submit price list/ }).click();
  await expect(page.getByRole('heading', { name: 'supplier-price-list.xlsx' })).toBeVisible();
  expect(page.url()).toBe(url);
  await expect(page.getByRole('region', { name: 'Products' }).getByRole('row').filter({ hasText: 'SKU-1001' })).toContainText('$4.60');
});

test('the portal and the editor change language together', async ({ page }) => {
  await openSample(page);
  await page.getByRole('combobox', { name: 'Language' }).selectOption('de');
  await expect(page.getByRole('heading', { name: 'Preisliste prüfen' })).toBeVisible();
  await expect(page.getByRole('button', { name: /^Lieferantenpreisliste/ })).toContainText('5 Fehler');
  await expect(panel(page).getByText('SKU-1099 steht nicht in unserem Katalog.', { exact: false })).toBeVisible();
  await expect(page.getByRole('button', { name: /Preisliste einreichen/ })).toBeVisible();
});

test('every page of the portal meets the accessibility checks', async ({ page }) => {
  const scan = async () => {
    await page.waitForTimeout(300);
    const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
    return violations.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(' ')).join(', ')}`);
  };
  await page.goto(serverUrl('portal'));
  expect(await scan()).toEqual([]);
  await page.getByRole('button', { name: 'Try a sample with problems' }).click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  expect(await scan()).toEqual([]);
  await fixErrors(page);
  await page.getByRole('button', { name: /Submit price list/ }).click();
  await expect(page.getByRole('heading', { name: 'supplier-price-list.xlsx' })).toBeVisible();
  expect(await scan()).toEqual([]);
});

test('selecting cells never moves the toolbar, and closing the panel gives the grid the room', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 });
  await openSample(page);
  const toolbar = page.locator('.wbe-toolbar');
  const height = async () => (await toolbar.boundingBox())!.height;
  // Measured once the fonts have loaded and the toolbar has folded for them.
  await page.evaluate(async () => {
    await document.fonts.ready;
    await new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)));
  });
  const before = await height();
  // The row buttons are there, but can't be used until a row of the table is selected.
  await expect(page.getByRole('button', { name: 'Insert row' })).toBeDisabled();
  await goTo(page, 'B3');
  await expect(page.getByRole('button', { name: 'Insert row' })).toBeEnabled();
  expect(await height()).toBe(before);
  await goTo(page, 'K20');
  expect(await height()).toBe(before);

  const workspace = (await page.locator('.wbe .wbe-workspace').boundingBox())!;
  await page.locator('.wbe-report-head').getByRole('button', { name: 'Close' }).click();
  const area = (await page.locator('.wbe .wbe-sheet-area').boundingBox())!;
  expect(area.x).toBe(workspace.x);
  expect(area.width).toBe(workspace.width);
});

test('the portal leaves out the report on what the file holds', async ({ page }) => {
  await openSample(page);
  await expect(page.getByRole('tab', { name: 'Checks' })).toBeVisible();
  await expect(page.getByRole('tab', { name: "What's in this file" })).toHaveCount(0);
  await expect(page.getByRole('button', { name: /What's in this file/ })).toHaveCount(0);
});
