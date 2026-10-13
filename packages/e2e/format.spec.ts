import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { goTo, grid as gridSurface, openFile, ordersWorkbook, readWorkbook, shown, tabKey } from './helpers';

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

test('types a custom number format, shows the active cell in it, and saves it', async ({ page }) => {
  const bar = page.getByRole('group', { name: 'Format' });
  const numberFormat = bar.getByRole('combobox', { name: 'Number format' });
  await goTo(page, 'C2:C3');
  await numberFormat.selectOption('more');
  const dialog = page.getByRole('dialog', { name: 'Number format' });
  const type = dialog.getByRole('textbox', { name: 'Type' });
  await expect(type).toBeFocused();
  await expect(type).toHaveValue('General');
  await expect(dialog.getByRole('listbox', { name: 'Formats' })).toHaveValue('General');

  await type.fill('"abc');
  await expect(dialog.getByRole('status', { name: 'Sample' })).toHaveText('');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByText('This number format cannot be used')).toBeVisible();
  await expect(type).toHaveAttribute('aria-invalid', 'true');

  await type.fill('0.0 "kg"');
  await expect(dialog.getByRole('status', { name: 'Sample' })).toHaveText('3.0 kg');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  await expect(shown(page, 'C3')).toHaveText('0.5 kg');
  await expect(numberFormat).toHaveValue('custom');

  // The workbook's own code is listed after Excel's, and picking one shows it.
  await numberFormat.selectOption('more');
  await expect(dialog.getByRole('listbox', { name: 'Formats' })).toHaveValue('0.0 "kg"');
  await dialog.getByRole('listbox', { name: 'Formats' }).selectOption('0%');
  await expect(type).toHaveValue('0%');
  await expect(dialog.getByRole('status', { name: 'Sample' })).toHaveText('300%');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(shown(page, 'C2')).toHaveText('3.0 kg');

  const lines = (await save(page)).getWorksheet('Lines')!;
  expect(lines.getCell('C3').numFmt).toBe('0.0 "kg"');
});

test('inserts Excel 365 checkboxes, ticks them by click and Space, unticks then removes them with Delete, and saves them', async ({ page }) => {
  const bar = page.getByRole('group', { name: 'Format' });
  const checkbox = bar.getByRole('button', { name: 'Checkbox' });
  await goTo(page, 'E2:E3');
  await checkbox.click();
  await expect(checkbox).toHaveAttribute('aria-pressed', 'true');
  await expect(shown(page, 'E2')).toHaveAttribute('aria-label', /^E2, Checkbox, not ticked/);
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('FALSE');

  // A click on E2's box ticks it; Space ticks the selected E3.
  const grid = (await gridSurface(page).first().boundingBox())!;
  await page.mouse.click(grid.x + 40 + 64 * 4 + 32, grid.y + 24 + 20 + 10);
  await expect(shown(page, 'E2')).toHaveAttribute('aria-label', /^E2, Checkbox, ticked/);
  await goTo(page, 'E3');
  await page.keyboard.press('Space');
  await expect(shown(page, 'E3')).toHaveAttribute('aria-label', /^E3, Checkbox, ticked/);

  // Delete unticks them first.
  await goTo(page, 'E2:E3');
  await page.keyboard.press('Delete');
  await expect(shown(page, 'E2')).toHaveAttribute('aria-label', /^E2, Checkbox, not ticked/);
  await expect(shown(page, 'E3')).toHaveAttribute('aria-label', /^E3, Checkbox, not ticked/);
  await goTo(page, 'E2');
  await page.keyboard.press('Space');

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const bytes = await readFile(await (await download).path());
  const files = unzipSync(new Uint8Array(bytes));
  expect(strFromU8(files['xl/featurePropertyBag/featurePropertyBag.xml'])).toContain('<bag type="Checkbox"/>');
  expect(strFromU8(files['xl/styles.xml'])).toContain('<xfpb:xfComplement i="0"/>');
  const lines = (await readWorkbook(bytes)).getWorksheet('Lines')!;
  expect([lines.getCell('E2').value, lines.getCell('E3').value]).toEqual([true, false]);

  // With none ticked, Delete takes them away.
  await goTo(page, 'E2:E3');
  await page.keyboard.press('Delete');
  await page.keyboard.press('Delete');
  await expect(checkbox).toHaveAttribute('aria-pressed', 'false');
  await expect(shown(page, 'E2')).not.toHaveAttribute('aria-label', /Checkbox/);
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
  await page.keyboard.press(tabKey(page));
  await page.keyboard.press(tabKey(page));
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
