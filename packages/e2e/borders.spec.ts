import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook } from './helpers';

test("draws borders with the line style and colour picked from Excel's Borders menu, and saves them as Excel does", async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  const borders = page.getByRole('group', { name: 'Format', exact: true }).getByRole('button', { name: 'Borders' });
  /** Opens the Borders menu and picks an item, or one of the choices in its submenu. */
  const pick = async (item: string, choice?: string) => {
    await borders.click();
    await page.getByRole('menuitem', { name: item, exact: true }).click();
    if (choice) await page.getByRole('menuitemradio', { name: choice, exact: true }).click();
  };
  await goTo(page, 'A2:B3');
  await pick('Line style', 'Thick');
  await pick('Line color', 'Red');
  // The menu shows the pen as it now is.
  await borders.click();
  await page.getByRole('menuitem', { name: 'Line style', exact: true }).click();
  await expect(page.getByRole('menuitemradio', { name: 'Thick', exact: true })).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');
  await pick('Outside borders');

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const wb = await readWorkbook(await readFile(await (await download).path()));
  const lines = wb.getWorksheet('Lines')!;
  expect(lines.getCell('A2').border).toMatchObject({ top: { style: 'thick', color: { argb: 'FFFF0000' } }, left: { style: 'thick' } });
  expect(lines.getCell('B3').border).toMatchObject({ bottom: { style: 'thick' }, right: { style: 'thick' } });
  expect(lines.getCell('A2').border?.bottom).toBeUndefined();
});
