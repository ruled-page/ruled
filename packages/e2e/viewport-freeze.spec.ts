import ExcelJS from 'exceljs';
import { expect, test, type Page } from '@playwright/test';
import { serverUrl } from './helpers';

/** handle.viewport on a drawn grid under a host freeze: the first row and column below and right of the frozen panes are the ones drawn there. */

type Editor = {
  freeze(panes: { rows: number; cols: number }): boolean;
  scrollTo(ref: string): boolean;
  viewport?: { sheet: string; row: number; col: number; cell: string };
};

/** A long sheet with rows of their own height and hidden rows, and columns of their own width, unfrozen in the file. */
async function workbook(): Promise<number[]> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Long list');
  for (let r = 1; r <= 400; r++) ws.addRow(Array.from({ length: 20 }, (_, c) => `R${r}C${c + 1}`));
  for (let r = 2; r <= 90; r += 7) ws.getRow(r).height = 44;
  for (const r of [5, 6, 30, 31, 32]) ws.getRow(r).hidden = true;
  ws.getColumn(2).width = 24;
  ws.getColumn(6).width = 4;
  ws.getColumn(8).hidden = true;
  return [...new Uint8Array(await wb.xlsx.writeBuffer())];
}

const viewport = (page: Page) => page.evaluate(() => (window as unknown as { editor: Editor }).editor.viewport);
const nameBox = (page: Page) => page.getByRole('combobox', { name: 'Go to cell' });

/** The cell drawn at a point of the scrolling grid, found by selecting it. */
async function cellAt(page: Page, x: number, y: number): Promise<string> {
  const box = (await page.getByTestId('wbe-grid-scroller').boundingBox())!;
  await page.mouse.click(box.x + x, box.y + y);
  return nameBox(page).inputValue();
}

/** Where the scrolling part of the grid starts: past the headings and frozen panes, found by clicking along each edge a pixel at a time until the cell is no longer a frozen one: that pixel is the first of the scrolled cells. */
async function paneEdge(page: Page, rows: number, cols: number): Promise<{ x: number; y: number }> {
  const parts = (ref: string) => /^([A-Z]+)(\d+)$/.exec(ref)!;
  let x = 44;
  if (cols) while (x < 1200 && parts(await cellAt(page, x, 28))[1] === 'A') x++;
  let y = 28;
  while (y < 700 && Number(parts(await cellAt(page, 44, y))[2]) <= rows) y++;
  return { x, y };
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 1300, height: 800 });
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  const bytes = await workbook();
  await page.evaluate(async (bytes) => {
    const url = '/packages/spreadsheet/dist/standalone/workbook-editor.js';
    const lib = (await import(/* @vite-ignore */ url)) as typeof import('@ruled/sheet/standalone');
    document.body.replaceChildren();
    document.body.style.cssText = 'margin:0;height:100vh';
    const host = document.createElement('div');
    host.style.cssText = 'height:100%';
    document.body.append(host);
    const editor = lib.mountWorkbookEditor(host, { warnOnUnload: false });
    await editor.open(Uint8Array.from(bytes), 'list.xlsx');
    (window as unknown as { editor: unknown }).editor = editor;
  }, bytes);
  await expect(page.getByTestId('wbe-grid-scroller')).toBeVisible();
  await expect.poll(() => viewport(page)).toMatchObject({ cell: 'A1' });
});

test('follows a host freeze: the first row and column past the panes are those drawn there', async ({ page }) => {
  expect(await page.evaluate(() => (window as unknown as { editor: Editor }).editor.freeze({ rows: 3, cols: 1 }))).toBe(true);
  // Frozen rows 1 to 3 (row 2 is tall) and column A stay in view, so the scrolled part starts at row 4 and column B.
  await expect.poll(() => viewport(page)).toMatchObject({ row: 4, col: 2, cell: 'B4' });
  const edge = await paneEdge(page, 3, 1);
  expect(await cellAt(page, edge.x, edge.y)).toBe('B4');

  // Aligned on a cell past hidden rows, tall rows and a hidden column.
  for (const [to, first] of [
    ['B40', 'B40'],
    ['A30', 'B33'],
    ['H60', 'I60'],
    ['C90', 'C90'],
  ]) {
    await page.evaluate((to) => (window as unknown as { editor: Editor }).editor.scrollTo(to), to);
    await expect.poll(async () => (await viewport(page))?.cell).toBe(first);
    expect(await cellAt(page, edge.x, edge.y)).toBe(first);
  }

  // Scrolled with the wheel, to places between cells.
  const box = (await page.getByTestId('wbe-grid-scroller').boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  let was = (await viewport(page))!.cell;
  for (const [dx, dy] of [
    [0, 97],
    [0, 410],
    [150, 0],
    [-60, -233],
    [0, 3000],
  ]) {
    await page.mouse.wheel(dx, dy);
    await expect.poll(async () => (await viewport(page))?.cell).not.toBe(was);
    // Between cells, the first may show only a sliver (and selecting it would scroll it fully into view), so check it by scrolling to the cell reported: it must then be the one drawn there.
    const reported = (await viewport(page))!.cell;
    await page.evaluate((to) => (window as unknown as { editor: Editor }).editor.scrollTo(to), reported);
    await expect.poll(async () => cellAt(page, edge.x + 1, edge.y + 1)).toBe(reported);
    expect((await viewport(page))!.cell).toBe(reported);
    const at = (await viewport(page))!;
    expect(at.row).toBeGreaterThan(3);
    expect(at.col).toBeGreaterThan(1);
    was = at.cell;
  }
});

test('goes back to the file’s panes when the freeze is lifted', async ({ page }) => {
  await page.evaluate(() => (window as unknown as { editor: Editor }).editor.freeze({ rows: 2, cols: 0 }));
  await expect.poll(() => viewport(page)).toMatchObject({ row: 3, col: 1, cell: 'A3' });
  await page.evaluate(() => (window as unknown as { editor: Editor }).editor.freeze({ rows: 0, cols: 0 }));
  await expect.poll(() => viewport(page)).toMatchObject({ row: 1, col: 1, cell: 'A1' });
});
