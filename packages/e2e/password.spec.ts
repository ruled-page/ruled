import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { setCell, shown, XLSX } from './helpers';

/** tests/corpus/outline-password.xlsx: outline.xlsx encrypted by msoffcrypto-tool with this password. */
const PASSWORD = 'Ruled-2026';
const file = () => readFileSync(new URL('../spreadsheet/tests/corpus/outline-password.xlsx', import.meta.url));

/** Chooses the protected workbook with the Open button's file picker, as a user would. */
async function choose(page: Page) {
  await page.goto('/');
  await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name: 'outline-password.xlsx', mimeType: XLSX, buffer: file() });
  return page.getByRole('dialog', { name: 'Password' });
}

test('asks for the password, again after a wrong one, then opens the workbook and saves it encrypted', async ({ page }) => {
  const dialog = await choose(page);
  await expect(dialog).toContainText('“outline-password.xlsx” is password-protected.');
  const password = dialog.getByLabel('Password', { exact: true });
  await expect(password).toBeFocused();
  await password.fill('ruled-2026');
  await password.press('Enter');
  await expect(dialog.getByRole('alert')).toHaveText("That password isn't right. Passwords are case-sensitive.");
  await expect(password).toHaveValue('');
  await password.fill(PASSWORD);
  await dialog.getByRole('button', { name: 'Open' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await expect(shown(page, 'A2')).toHaveText('North A');

  await setCell(page, 'B2', '99');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const saved = await readFile(await (await download).path());
  // A compound file holding the encrypted package, not a plain .xlsx zip.
  expect([...saved.subarray(0, 4)]).toEqual([0xd0, 0xcf, 0x11, 0xe0]);
  expect(saved.includes(Buffer.from('E\0n\0c\0r\0y\0p\0t\0e\0d\0P\0a\0c\0k\0a\0g\0e\0'))).toBe(true);
});

test('cancelled, leaves the workbook unopened and says why', async ({ page }) => {
  const dialog = await choose(page);
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByText('This workbook is password-protected. Enter its password to open it.')).toBeVisible();
});
