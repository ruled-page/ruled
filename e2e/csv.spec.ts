import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { readWorkbook, setCell, shown } from './helpers';

/** A CSV file as Excel's "CSV UTF-8" writes it: a byte order mark, quoted fields where needed, CRLF. */
const PRICES = '﻿Item,Price,Note\r\nBolts,"1,234.50",=1+2\r\nNuts,0.25,"two\nlines"\r\n';

test.describe('CSV files', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    const input = page.locator('input[type="file"]');
    await expect(input).toHaveAttribute('accept', /\.csv/);
    await input.setInputFiles({ name: 'prices.csv', mimeType: 'text/csv', buffer: Buffer.from(PRICES) });
    await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  });

  test('open from the file picker as a sheet named after the file, with formulas kept as text', async ({ page }) => {
    await expect(page.getByRole('tab', { name: 'prices' })).toHaveAttribute('aria-selected', 'true');
    await expect(shown(page, 'A2')).toHaveText('Bolts');
    await expect(shown(page, 'B2')).toHaveText('1,234.50');
    await expect(shown(page, 'C2')).toHaveText('=1+2');
    await expect(shown(page, 'B3')).toHaveText('0.25');
  });

  test('save as .xlsx with the edits', async ({ page }) => {
    await setCell(page, 'B3', '0.5');
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: /Save as Excel/ }).click();
    const file = await download;
    expect(file.suggestedFilename()).toBe('prices.xlsx');

    const sheet = (await readWorkbook(await readFile(await file.path()))).getWorksheet('prices')!;
    expect(sheet.getCell('B2').value).toBe(1234.5);
    expect(sheet.getCell('C2').value).toBe('=1+2');
    expect(sheet.getCell('C3').value).toBe('two\nlines');
    expect(sheet.getCell('B3').value).toBe(0.5);
  });

  test('download the sheet shown as CSV, values as shown and formulas escaped', async ({ page }) => {
    await setCell(page, 'D2', '=LEN(A2)');
    await setCell(page, 'A3', '@cmd');
    await expect(shown(page, 'D2')).toHaveText('5');

    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Download this sheet as CSV' }).click();
    const file = await download;
    expect(file.suggestedFilename()).toBe('prices.csv');
    await expect(page.getByText('Downloaded prices.csv: the values on prices, as shown.', { exact: false })).toBeVisible();

    const bytes = await readFile(await file.path());
    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(bytes.subarray(3).toString('utf8')).toBe(['Item,Price,Note,', `Bolts,"1,234.50",'=1+2,5`, `'@cmd,0.25,"two\nlines",`, ''].join('\r\n'));
  });
});
