import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
import { goTo, openFile, shown } from './helpers';

test("saves a LET formula another tool wrote without Excel's _xlpm. prefix with the prefix added", async ({ page }) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 2;
  // As some tools write it: Excel shows this as #NAME?, and Excel for the web refuses the file.
  ws.getCell('B1').value = { formula: '_xlfn.LET(x,A1*2,x+1)', result: 5 };
  await openFile(page, 'let.xlsx', new Uint8Array(await wb.xlsx.writeBuffer()));
  await expect(shown(page, 'B1')).toHaveText('5');
  await goTo(page, 'B1');
  await expect(page.getByRole('combobox', { name: 'Cell contents' })).toHaveValue('=LET(x,A1*2,x+1)');

  // Another cell edited, the formula is worked out again and saved with the prefix.
  await goTo(page, 'A1');
  await page.keyboard.type('3');
  await page.keyboard.press('Enter');
  await expect(shown(page, 'B1')).toHaveText('7');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const sheet = strFromU8(unzipSync(await readFile(await (await download).path()))['xl/worksheets/sheet1.xml']);
  expect(/<c r="B1"[^>]*>(.*?)<\/c>/.exec(sheet)?.[1]).toBe('<f>_xlfn.LET(_xlpm.x,A1*2,_xlpm.x+1)</f><v>7</v>');
});
