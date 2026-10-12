import { expect, test } from '@playwright/test';
import { goTo, openSample, setCell, shown } from './helpers';

test('a pivot table follows its source as it is edited, grows with a new item, and refreshes on demand', async ({ page }) => {
  await openSample(page, 'Pivot table');
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'B5')).toHaveText(/^1,?315$/);
  const north = Number((await shown(page, 'B5').textContent())!.replace(/,/g, ''));

  // Data!C2 is one of North's rows: 100 more units there.
  await page.getByRole('tab', { name: 'Data' }).click();
  await goTo(page, 'C2');
  const units = Number(await page.getByRole('combobox', { name: 'Cell contents' }).inputValue());
  await setCell(page, 'C2', String(units + 100));
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'B5')).toHaveText(new RegExp(`^${(north + 100).toLocaleString('en-US').replace(',', ',?')}$`));
  await expect(shown(page, 'A8')).toHaveText('Grand Total');

  // A region new to the data gets a row of its own.
  await page.getByRole('tab', { name: 'Data' }).click();
  await setCell(page, 'A3', 'Central');
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'A8')).toHaveText('Central');
  await expect(shown(page, 'A9')).toHaveText('Grand Total');

  // Its cells stay read-only, and Refresh is in their menu.
  await goTo(page, 'B5');
  await page.keyboard.press('Shift+F10');
  await expect(page.getByRole('menuitem', { name: 'Refresh all' })).toBeVisible();
  await page.getByRole('menuitem', { name: /^Refresh/ }).first().click();
  await expect(shown(page, 'A9')).toHaveText('Grand Total');
});
