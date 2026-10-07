import { expect, test } from '@playwright/test';
import { setCell } from './helpers';

/** The <workbook-editor> element in a plain HTML page, saving to the host through its save event. */
test('hands the host the workbook and its values on save', async ({ page }) => {
  await page.goto('http://localhost:4174/examples/html/');
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();

  await setCell(page, 'B2', '125000');
  await expect(page.locator('#status')).toHaveText('1 change, not saved');

  // The sample has a negative cost in E3, which the host's schema doesn't allow.
  const save = page.getByRole('button', { name: 'Save from the host' });
  await save.click();
  await expect(page.getByText('Not saved. 1 cell breaks the “Forecast (from the host page)” schema.')).toBeVisible();
  await expect(page.locator('#received')).toHaveText('What the host receives on save appears here.');

  await setCell(page, 'E3', '30000');
  await save.click();
  await expect(page.locator('#status')).toHaveText('2 changes, saved');

  const received = JSON.parse((await page.locator('#received').textContent()) ?? '{}');
  expect(received.fileName).toBe('forecast.xlsx');
  expect(received.bytes).toBeGreaterThan(1000);
  expect(received.schema).toMatchObject({ id: 'example.forecast', version: '1', errors: 0 });
  expect(received.values.Forecast.tables.quarters[0]).toMatchObject({ revenue: 125000 });
  expect(received.values.Forecast.tables.quarters[3]).toMatchObject({ cost: 30000 });
  expect(received.changes.map((c: { ref: string }) => c.ref)).toEqual(['B2', 'E3']);
});
