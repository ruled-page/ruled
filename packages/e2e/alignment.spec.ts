import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook } from './helpers';

test('aligns text to the top and indents it from the format bar, and saves both as Excel does', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  const bar = page.getByRole('group', { name: 'Format', exact: true });
  await goTo(page, 'A2');
  // Bottom is Excel's own, so it shows as on.
  await expect(bar.getByRole('button', { name: 'Bottom align' })).toHaveAttribute('aria-pressed', 'true');
  await expect(bar.getByRole('button', { name: 'Decrease indent' })).toBeDisabled();
  await bar.getByRole('button', { name: 'Top align' }).click();
  await expect(bar.getByRole('button', { name: 'Top align' })).toHaveAttribute('aria-pressed', 'true');
  await bar.getByRole('button', { name: 'Increase indent' }).click();
  await bar.getByRole('button', { name: 'Increase indent' }).click();
  await expect(bar.getByRole('button', { name: 'Decrease indent' })).toBeEnabled();
  await expect(bar.getByRole('button', { name: 'Align left' })).toHaveAttribute('aria-pressed', 'true');

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const wb = await readWorkbook(await readFile(await (await download).path()));
  expect(wb.getWorksheet('Lines')!.getCell('A2').alignment).toMatchObject({ vertical: 'top', horizontal: 'left', indent: 2 });
});
