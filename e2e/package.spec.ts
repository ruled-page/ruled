import { expect, test } from '@playwright/test';
import { ordersWorkbook, serverUrl, setCell, shown, XLSX } from './helpers';

/**
 * The React component as a host gets it: installed from the packed library
 * into the app in e2e/host-app, built by that app's own Vite and served
 * under /app/ (scripts/host-app.mjs). Run with E2E_HOST_APP=1.
 */
test('the packaged React component opens, recalculates and edits a workbook in a host app', async ({ page }) => {
  const failed: string[] = [];
  page.on('requestfailed', (r) => failed.push(r.url()));
  page.on('response', (r) => {
    if (r.status() >= 400) failed.push(`${r.status()} ${r.url()}`);
  });
  await page.goto(`${serverUrl('hostApp')}/app/`);
  // The app's own React, not this repository's.
  await expect(page.locator('html')).toHaveAttribute('data-react', /^18\./);
  await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name: 'orders.xlsx', mimeType: XLSX, buffer: Buffer.from(await ordersWorkbook()) });
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await expect(shown(page, 'D2')).toHaveText('6');
  // A result that only the worker's engine can work out.
  await setCell(page, 'B2', '10');
  await expect(shown(page, 'D2')).toHaveText('30');
  expect(failed).toEqual([]);
});
