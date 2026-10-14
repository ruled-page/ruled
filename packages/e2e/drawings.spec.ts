import { readFileSync } from 'node:fs';
import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { focusedGrid, openFile } from './helpers';

/** How many of the grid's pixels are drawn in this colour, near enough. */
async function gridPixels(page: Page, [r, g, b]: [number, number, number]): Promise<number> {
  return page.evaluate(
    ([r, g, b]) => {
      let n = 0;
      for (const canvas of document.querySelectorAll<HTMLCanvasElement>('[data-testid="wbe-grid"] canvas')) {
        const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
        for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i] - r) + Math.abs(data[i + 1] - g) + Math.abs(data[i + 2] - b) < 30) n++;
      }
      return n;
    },
    [r, g, b],
  );
}

/** Sheet S: numbers in A2:E2, and in F2 a line sparkline of them in magenta, with gaps for empty cells. */
async function sparklineBook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['a', 'b', 'c', 'd', 'e', 'Trend']);
  ws.addRow([3, 8, 2, 9, 5]);
  ws.getColumn(6).width = 20;
  ws.getRow(2).height = 30;
  const files = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const group =
    '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup displayEmptyCellsAs="gap" lineWeight="2.25">' +
    '<x14:colorSeries rgb="FFFF00FF"/><x14:sparklines><x14:sparkline><xm:f>S!A2:E2</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline></x14:sparklines>' +
    '</x14:sparklineGroup></x14:sparklineGroups></ext></extLst>';
  const path = 'xl/worksheets/sheet1.xml';
  files[path] = strToU8(strFromU8(files[path]).replace('</worksheet>', `${group}</worksheet>`));
  return zipSync(files);
}

test('draws a sparkline from its cells, and again as they change', async ({ page }) => {
  await openFile(page, 'sparkline.xlsx', await sparklineBook());
  const magenta: [number, number, number] = [255, 0, 255];
  await expect.poll(() => gridPixels(page, magenta)).toBeGreaterThan(40);
  // Emptied, its cells leave nothing to draw.
  const box = page.getByRole('combobox', { name: 'Go to cell' });
  await box.fill('A2:E2');
  await box.press('Enter');
  await expect(focusedGrid(page)).toHaveCount(1);
  await page.keyboard.press('Delete');
  await expect.poll(() => gridPixels(page, magenta)).toBe(0);
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => gridPixels(page, magenta)).toBeGreaterThan(40);
});

test('draws a picture and a text box where they sit, in place of placeholders', async ({ page }) => {
  // The demo's kitchen sink: a 48 pixel picture on its Notes sheet.
  const bytes = new Uint8Array(readFileSync(new URL('../../examples/demo/public/samples/kitchen-sink.xlsx', import.meta.url)));
  await openFile(page, 'kitchen-sink.xlsx', bytes);
  await page.getByRole('navigation', { name: 'Sheets' }).getByRole('tab', { name: 'Notes' }).click();
  const picture = page.locator('.wbe-object-picture img');
  await expect(picture).toBeVisible();
  expect(await picture.evaluate((img: HTMLImageElement) => [img.complete, img.naturalWidth])).toEqual([true, 48]);
  await expect(page.locator('.wbe-object-image')).toHaveCount(0);
  // The text box beside it, with its text, in the editor's stand-in for Calibri where the device lacks it.
  const box = page.getByRole('img', { name: 'TextBox 2: Text boxes are kept too.' });
  await expect(box).toBeVisible();
  await expect(box.locator('.wbe-shape-text')).toHaveText('Text boxes are kept too.');
  expect(await box.locator('path').getAttribute('stroke')).toBe('#BCBCBC');
  await expect(page.locator('.wbe-object-shape')).toHaveCount(0);
});
