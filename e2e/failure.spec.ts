import { expect, test } from '@playwright/test';
import { openFile, ordersWorkbook, shown, XLSX } from './helpers';

test('a worker that can’t load is reported, and the file reopens once it can', async ({ page }) => {
  // Stands in for a worker blocked by a content security policy or missing from the server.
  await page.route('**/workbook.worker*', (route) => route.abort());
  await page.goto('/');
  await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name: 'orders.xlsx', mimeType: XLSX, buffer: Buffer.from(await ordersWorkbook()) });
  const alert = page.getByRole('alert').filter({ hasText: 'The editor stopped working' });
  await expect(alert).toBeVisible();

  await page.unroute('**/workbook.worker*');
  await alert.getByRole('button', { name: 'Reopen orders.xlsx' }).click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await expect(shown(page, 'D2')).toHaveText('6');
});

test('a workbook that opens normally offers no reopening', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await expect(page.getByRole('button', { name: /^Reopen/ })).toHaveCount(0);
});

test('an error while drawing shows a way to carry on, rather than an empty page', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  // The status bar formats the sum of a selection as it draws: 2 + 3 + 4 + 0.5 for B2:C3. Make that
  // fail until told otherwise (React tries drawing a second time before giving up).
  await page.evaluate(() => {
    const format = Number.prototype.toLocaleString;
    const w = window as unknown as { breakSum: boolean };
    w.breakSum = true;
    Number.prototype.toLocaleString = function (this: number, ...args: Parameters<typeof format>) {
      if (Number(this) === 9.5 && w.breakSum) throw new Error('drawing failed');
      return format.apply(this, args);
    };
  });
  await page.getByRole('combobox', { name: 'Go to cell' }).fill('B2:C3');
  await page.getByRole('combobox', { name: 'Go to cell' }).press('Enter');
  const alert = page.getByRole('alert').filter({ hasText: 'Something went wrong showing the workbook' });
  await expect(alert).toBeVisible();
  await page.evaluate(() => ((window as unknown as { breakSum: boolean }).breakSum = false));
  await alert.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await expect(shown(page, 'D2')).toHaveText('6');
});
