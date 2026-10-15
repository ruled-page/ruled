import { expect, test, type Page } from '@playwright/test';
import { goTo, serverUrl, setCell } from './helpers';

/** The <workbook-editor> element with a plugin whose file source keeps one workbook in memory (fixtures/plugins.html). */
const drive = (page: Page) => page.evaluate(() => (window as unknown as { drive: { version: number; writes: string[]; copies: string[] } }).drive);

test('opens from a plugin’s source, saves back, catches a conflict and saves a copy', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/packages/e2e/fixtures/plugins.html`);

  // With a source registered, Open is a menu: this device first, then the source.
  await page.getByRole('button', { name: 'Open', exact: true }).click();
  const openMenu = page.getByRole('menu', { name: 'Open from' });
  await expect(openMenu.getByRole('menuitem')).toHaveText(['This device', 'Fixture Drive']);
  await openMenu.getByRole('menuitem', { name: 'Fixture Drive' }).click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await expect(page.locator('.wbe-file-name')).toHaveAttribute('title', 'Team/Planning/forecast.xlsx');

  // Without a host's Save, the primary Save writes back to the drive, and says when it hasn't yet.
  await setCell(page, 'B2', '125000');
  await page.getByRole('button', { name: 'Save to Fixture Drive (changes not saved there yet)' }).click();
  await expect(page.getByText('Saved forecast.xlsx to Fixture Drive.')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Save to Fixture Drive', exact: true })).toBeVisible();
  expect(await drive(page)).toMatchObject({ version: 2, writes: ['forecast.xlsx'] });

  // Someone else saves the file meanwhile: the changes stay, and nothing is written over theirs.
  await setCell(page, 'B3', '1');
  await page.evaluate(() => (window as unknown as { drive: { version: number } }).drive.version++);
  await page.keyboard.press('ControlOrMeta+s');
  await expect(page.getByRole('alertdialog', { name: 'This workbook was changed elsewhere' })).toBeVisible();
  await page.getByRole('button', { name: 'Keep editing' }).click();
  expect((await drive(page)).writes).toHaveLength(1);
  await expect(page.getByRole('button', { name: 'Save to Fixture Drive (changes not saved there yet)' })).toBeVisible();

  // "Save a copy": this device, or the drive, as a new file.
  await page.getByRole('button', { name: 'Save a copy' }).click();
  const copyMenu = page.getByRole('menu', { name: 'Save a copy' });
  await expect(copyMenu.getByRole('menuitem')).toHaveText(['This device', 'Fixture Drive']);
  await copyMenu.getByRole('menuitem', { name: 'Fixture Drive' }).click();
  await expect(page.getByText('Saved a copy, forecast copy.xlsx, to Fixture Drive.')).toBeVisible();
  expect((await drive(page)).copies).toEqual(['forecast.xlsx']);
  // The workbook still belongs where it came from.
  await expect(page.locator('.wbe-file-name')).toHaveText('forecast.xlsx');
});

test('runs a plugin’s cell-menu command with the selection', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/packages/e2e/fixtures/plugins.html`);
  await page.getByRole('button', { name: 'Open', exact: true }).click();
  await page.getByRole('menu', { name: 'Open from' }).getByRole('menuitem', { name: 'Fixture Drive' }).click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();

  await goTo(page, 'B2:C3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Where is this?' }).click();
  await expect(page.getByText('Forecast!B2 (B2:C3)')).toBeVisible();
});
