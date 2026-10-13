import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import { grid, gridScroller, openFile, shown } from './helpers';

test.skip(process.env.E2E_GRID !== 'ruled', "Glide's grid lays every sheet out left to right");

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
