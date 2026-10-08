import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook, shown } from './helpers';

async function save(page: Page) {
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  return readWorkbook(await readFile(await (await download).path()));
}

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test('formats cells from the format bar, and saves the formats', async ({ page }) => {
  const bar = page.getByRole('group', { name: 'Format' });
  await goTo(page, 'A1:D1');
  await bar.getByRole('button', { name: 'Bold' }).click();
  await expect(bar.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true');
  await bar.getByRole('button', { name: 'Fill colour' }).click();
  await page.getByRole('dialog', { name: 'Fill colour' }).getByRole('button', { name: 'Yellow' }).click();
  await bar.getByRole('button', { name: 'Borders' }).click();
  await page.getByRole('menuitem', { name: 'Outside borders' }).click();
  await bar.getByRole('combobox', { name: 'Font', exact: true }).selectOption('Georgia');

  await goTo(page, 'C2:C3');
  await bar.getByRole('combobox', { name: 'Number format' }).selectOption('currency');
  await expect(shown(page, 'C3')).toHaveText('$0.50');
  await bar.getByRole('button', { name: 'Increase decimal places' }).click();
  await expect(shown(page, 'C3')).toHaveText('$0.500');
  await bar.getByRole('button', { name: 'Centre' }).click();

  const lines = (await save(page)).getWorksheet('Lines')!;
  expect(lines.getCell('A1').font).toMatchObject({ bold: true, name: 'Georgia' });
  expect(lines.getCell('B1').fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } });
  expect(lines.getCell('A1').border).toMatchObject({ top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' } });
  expect(lines.getCell('B1').border.left).toBeUndefined();
  expect(lines.getCell('C2').numFmt).toBe('"$"#,##0.000');
  expect(lines.getCell('C2').alignment).toMatchObject({ horizontal: 'center' });
});

test('undoes a format change, and clears formats', async ({ page }) => {
  const bar = page.getByRole('group', { name: 'Format' });
  await goTo(page, 'A2');
  await bar.getByRole('button', { name: 'Italic' }).click();
  await expect(bar.getByRole('button', { name: 'Italic' })).toHaveAttribute('aria-pressed', 'true');
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(bar.getByRole('button', { name: 'Italic' })).toHaveAttribute('aria-pressed', 'false');
  await goTo(page, 'A1:D1');
  await bar.getByRole('button', { name: 'Bold' }).click();
  await bar.getByRole('button', { name: 'Clear formats' }).click();
  await expect(bar.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'false');
});

test('opens the colour palette and the borders menu from the keyboard', async ({ page }) => {
  const bar = page.getByRole('group', { name: 'Format' });
  await goTo(page, 'B2');
  await bar.getByRole('button', { name: 'Font colour' }).focus();
  await page.keyboard.press('Enter');
  const palette = page.getByRole('dialog', { name: 'Font colour' });
  await expect(palette.getByRole('button', { name: 'Automatic' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(palette).toBeHidden();
  await expect(bar.getByRole('button', { name: 'Font colour' })).toBeFocused();
  await page.keyboard.press('Enter');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Enter');
  await expect(palette).toBeHidden();
  await bar.getByRole('button', { name: 'Borders' }).focus();
  await page.keyboard.press('Enter');
  await expect(page.getByRole('menuitem', { name: 'All borders' })).toBeFocused();
  await page.keyboard.press('Enter');
  const lines = (await save(page)).getWorksheet('Lines')!;
  expect(lines.getCell('B2').font?.color).toMatchObject({ argb: 'FFFF0000' });
  expect(lines.getCell('B2').border).toMatchObject({ top: { style: 'thin' }, right: { style: 'thin' } });
});
