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
