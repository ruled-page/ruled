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

test('pastes from Excel with its formats, and pastes transposed from the menu', async ({ page, browserName }) => {
  test.skip(browserName === 'firefox', 'Firefox gives a paste event made by a script its own empty clipboard data.');
  const bar = page.getByRole('group', { name: 'Format' });
  await goTo(page, 'A6');
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'Bold\tRed\r\n');
    data.setData('text/html', '<style><!--.xl65 {font-weight:700;} .xl66 {color:red; background:yellow;}--></style><table><tr><td class=xl65>Bold</td><td class=xl66>Red</td></tr></table>');
    document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await expect(shown(page, 'B6')).toHaveText('Red');
  await expect(bar.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true');

  // Copy A1:B2 and paste it on its side at D6.
  await goTo(page, 'A1:B2');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Copy' }).click();
  await goTo(page, 'D6');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Paste special' }).click();
  await page.getByRole('menuitem', { name: 'Transposed' }).click();
  await expect(shown(page, 'E6')).toHaveText('Bolts');
  await expect(shown(page, 'D7')).toHaveText('Qty');

  const lines = (await save(page)).getWorksheet('Lines')!;
  expect(lines.getCell('A6').font).toMatchObject({ bold: true });
  expect(lines.getCell('B6').fill).toMatchObject({ fgColor: { argb: 'FFFFFF00' } });
});
