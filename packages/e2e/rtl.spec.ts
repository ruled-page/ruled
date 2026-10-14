import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { grid, gridScroller, openFile, shown } from './helpers';

test('a sheet read right to left has column A at the right, and selects, edits and moves from there', async ({ page }) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Arabic', { views: [{ rightToLeft: true }] });
  ws.addRow(['one', 'two', 'three']);
  ws.addRow([1, 2, 3]);
  await openFile(page, 'rtl.xlsx', new Uint8Array(await wb.xlsx.writeBuffer()));
  const nameBox = page.getByRole('combobox', { name: 'Go to cell' });
  const box = (await grid(page).first().boundingBox())!;
  // The row numbers are at the right, 40 pixels wide, under a 24 pixel heading; A's 64 pixels are before them, then B's.
  const right = box.x + box.width;
  await page.mouse.click(right - 40 - 32, box.y + 24 + 10);
  await expect(nameBox).toHaveValue('A1');
  await page.mouse.click(right - 40 - 64 - 32, box.y + 24 + 20 + 10);
  await expect(nameBox).toHaveValue('B2');
  await expect(shown(page, 'B2')).toHaveText('2');
  // Typing replaces what B2 holds; the left arrow then goes on to C, the next column.
  await page.keyboard.type('5', { delay: 20 });
  await page.keyboard.press('Enter');
  await expect(shown(page, 'B2')).toHaveText('5');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowLeft');
  await expect(nameBox).toHaveValue('C2');
  // The scrollbar is at the left, scrolled to the start: the right.
  expect(await gridScroller(page).first().evaluate((s) => getComputedStyle(s).direction)).toBe('rtl');
});

test('a sheet read right to left has its outline bars mirrored: the row bar at the right, beside the row numbers', async ({ page }) => {
  // tests/corpus/outline.xlsx, read right to left: columns B:D grouped beside their total in E.
  const files = unzipSync(new Uint8Array(readFileSync(new URL('../spreadsheet/tests/corpus/outline.xlsx', import.meta.url))));
  files['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(files['xl/worksheets/sheet1.xml']).replace('<sheetView ', '<sheetView rightToLeft="1" '));
  await openFile(page, 'outline-rtl.xlsx', zipSync(files));
  const box = (await grid(page).first().boundingBox())!;
  const rowBar = (await page.getByRole('group', { name: 'Row outline' }).boundingBox())!;
  expect(rowBar.x).toBeGreaterThanOrEqual(box.x + box.width - 1);
  // The column levels are over the row numbers, at the right; B:D's button over E, to the left of D.
  const level = (await page.getByRole('button', { name: 'Show column level 1' }).boundingBox())!;
  expect(level.x).toBeGreaterThan(box.x + box.width - 40);
  const toggle = page.getByRole('button', { name: 'Collapse columns B to D' });
  // Under the button is column E.
  const t = (await toggle.boundingBox())!;
  await page.mouse.click(t.x + t.width / 2, box.y + 24 + 10);
  await expect(page.getByRole('combobox', { name: 'Go to cell' })).toHaveValue('E1');
  await toggle.click();
  await expect(page.getByRole('button', { name: 'Expand columns B to D' })).toBeVisible();
});
