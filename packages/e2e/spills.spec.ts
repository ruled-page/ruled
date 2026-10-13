import { readFileSync } from 'node:fs';
import { expect, test, type Page } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { goTo, openFile, ordersWorkbook, setCell, shown } from './helpers';

/** Dynamic arrays: formulas whose results spill into the cells around them, as in Excel 365. */

const formulaBar = (page: Page) => page.getByRole('combobox', { name: 'Cell contents' });
const templates = () => new Uint8Array(readFileSync(new URL('../spreadsheet/tests/corpus/templates.xlsx', import.meta.url)));

/** How many of the grid's pixels are drawn in this colour. */
async function gridPixels(page: Page, color: string): Promise<number> {
  return page.evaluate((css) => {
    const probe = document.createElement('canvas').getContext('2d')!;
    probe.fillStyle = css;
    probe.fillRect(0, 0, 1, 1);
    const [r, g, b] = probe.getImageData(0, 0, 1, 1).data;
    let n = 0;
    for (const canvas of document.querySelectorAll<HTMLCanvasElement>('[data-testid="wbe-grid"] canvas')) {
      const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
      for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i] - r) + Math.abs(data[i + 1] - g) + Math.abs(data[i + 2] - b) < 6) n++;
    }
    return n;
  }, color);
}

const spillLine = (page: Page) => page.evaluate(() => getComputedStyle(document.querySelector('.wbe-root')!).getPropertyValue('--wbe-grid-spill-line').trim());

test.describe('a formula whose result is an array', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
  });

  test('spills into the cells below and beside it, which screen readers hear it spilled into', async ({ page }) => {
    await setCell(page, 'F2', '=SEQUENCE(3,2)*B2');
    await expect(shown(page, 'F2')).toHaveAttribute('aria-label', 'F2, 2, formula, changed');
    await expect(shown(page, 'G4')).toHaveText('12');
    await expect(shown(page, 'G4')).toHaveAttribute('aria-label', 'G4, 12, spilled from F2');
    await expect(shown(page, 'F4')).toHaveAttribute('aria-label', 'F4, 10, spilled from F2');

    // A spilled cell holds nothing itself: the formula bar shows its formula greyed, and says where it comes from.
    await goTo(page, 'G3');
    await expect(formulaBar(page)).toHaveValue('');
    await expect(formulaBar(page)).toHaveAttribute('placeholder', '=SEQUENCE(3,2)*B2');
    await expect(page.locator('.wbe-cell-note')).toContainText('Spilled from the formula in F2.');

    // Its inputs change, and so does it.
    await setCell(page, 'B2', '3');
    await expect(shown(page, 'G4')).toHaveAttribute('aria-label', 'G4, 18, spilled from F2');
  });

  test('is outlined while it is selected, dashed while something is in its way', async ({ page }) => {
    await setCell(page, 'F2', '=SEQUENCE(3)');
    const color = await spillLine(page);
    await goTo(page, 'A6');
    const before = await gridPixels(page, color);
    await goTo(page, 'F3');
    await expect.poll(() => gridPixels(page, color)).toBeGreaterThan(before + 40);

    // Typing into one of its cells blocks it; undoing that spills it again.
    await setCell(page, 'F3', 'x');
    await expect(shown(page, 'F2')).toHaveText('#SPILL!');
    await expect(shown(page, 'F4')).toHaveAttribute('aria-label', 'F4, Empty');
    await goTo(page, 'F2');
    await expect(page.locator('.wbe-cell-note')).toHaveText('Its results need F2:F4, and something there is in the way.');
    await expect.poll(() => gridPixels(page, color)).toBeGreaterThan(before);
    await page.getByRole('button', { name: 'Undo' }).click();
    await expect(shown(page, 'F4')).toHaveAttribute('aria-label', 'F4, 3, spilled from F2');
  });

  test('saves as a dynamic array, with its values in the cells it spills into', async ({ page }) => {
    await setCell(page, 'F2', '=SEQUENCE(2,2)');
    await expect(shown(page, 'G3')).toHaveAttribute('aria-label', /^G3, 4, spilled from F2/);
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: /Save as Excel/ }).click();
    const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(await (await download).path())))['xl/worksheets/sheet1.xml']);
    expect(xml).toMatch(/<c r="F2" cm="1"><f t="array" ref="F2:G3">_xlfn.SEQUENCE\(2,2\)<\/f><v>1<\/v><\/c>/);
    expect(xml).toContain('<c r="G3"><v>4</v></c>');
  });
});

test("a file's own dynamic arrays spill here, and can be typed over, while older array formulas recalculate but stay locked", async ({ page }) => {
  await openFile(page, 'templates.xlsx', templates());
  await goTo(page, 'Arrays!A3');
  await expect(shown(page, 'A3')).toHaveAttribute('aria-label', 'A3, Hardware, spilled from A2');
  await expect(shown(page, 'H3')).toHaveAttribute('aria-label', 'H3, 5, spilled from G2');
  await expect(formulaBar(page)).toHaveAttribute('placeholder', '=SORT(UNIQUE(Orders[Category]))');
  await setCell(page, 'A3', 'Mixed');
  await expect(shown(page, 'A2')).toHaveText('#SPILL!');
  await goTo(page, 'G7');
  await expect(formulaBar(page)).toHaveAttribute('readonly', '');
  // Each cell of an array formula entered with Ctrl+Shift+Enter shows it in braces, as in Excel.
  await expect(formulaBar(page)).toHaveValue('{=Orders[Price]*2}');
  await expect(page.locator('.wbe-cell-note')).toContainText('array formula');
  // G6:G8 doubles the order prices; Nuts' is on Lines!D3.
  await setCell(page, 'Lines!D3', '5');
  await goTo(page, 'Arrays!G7');
  await expect(shown(page, 'G7')).toHaveText('10.00');
});
