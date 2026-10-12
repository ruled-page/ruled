import ExcelJS from 'exceljs';
import { expect, test } from '@playwright/test';
import { openFile, shown } from './helpers';

/** 100,001 rows of a number and two formulas: past the 200,000 formulas recalculated as you type. */
async function bigWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Big');
  for (let r = 1; r <= 100_001; r++) ws.addRow([r, { formula: `A${r}*2`, result: r * 2 }, { formula: `B${r}+1`, result: r * 2 + 1 }]);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

test('says when a big workbook opens with live recalculation off, and turns it on from there', async ({ page }) => {
  test.slow();
  await openFile(page, 'big.xlsx', await bigWorkbook());
  const banner = page.getByRole('status').filter({ hasText: 'Live recalculation is off' });
  await expect(banner).toContainText('this workbook has 200,002 formulas');
  await banner.getByRole('button', { name: 'Turn on live recalculation' }).click();
  await expect(banner).toBeHidden({ timeout: 60_000 });
  await expect(page.getByText(/formulas recalculate live/)).toBeVisible();
  await expect(shown(page, 'C1')).toHaveText('3');
});
