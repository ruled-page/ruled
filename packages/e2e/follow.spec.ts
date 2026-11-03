import ExcelJS from 'exceljs';
import { expect, test, type Page } from '@playwright/test';
import { serverUrl } from './helpers';

/**
 * Two editors side by side, each following the other's selection and scroll
 * position through nothing but the handle: selection, viewport, subscribe,
 * goTo without the focus and scrollTo.
 */

/** A long sheet with rows of their own height, hidden rows, a collapsed group and frozen panes. */
async function workbook(): Promise<number[]> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Long list');
  for (let r = 1; r <= 400; r++) ws.addRow([`Row ${r}`, r, r * 2, r * 3]);
  for (let r = 20; r <= 40; r += 5) ws.getRow(r).height = 48;
  for (const r of [9, 10, 11]) ws.getRow(r).hidden = true;
  for (let r = 50; r <= 70; r++) {
    ws.getRow(r).outlineLevel = 1;
    ws.getRow(r).hidden = true;
  }
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];
  wb.addWorksheet('Other').addRow(['other']);
  return [...new Uint8Array(await wb.xlsx.writeBuffer())];
}

interface Linked {
  left: { selection?: { sheet: string; active: string; ranges: string[] }; viewport?: { sheet: string; row: number; col: number; cell: string } };
  right: Linked['left'];
}

/** What each editor's handle says now. */
const state = (page: Page) =>
  page.evaluate(() => {
    const { left, right } = (window as unknown as { editors: Record<string, { selection: unknown; viewport: unknown }> }).editors;
    return { left: { selection: left.selection, viewport: left.viewport }, right: { selection: right.selection, viewport: right.viewport } } as Linked;
  });

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 800 });
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  const bytes = await workbook();
  await page.evaluate(async (bytes) => {
    const url = '/packages/spreadsheet/dist/standalone/workbook-editor.js';
    const lib = (await import(/* @vite-ignore */ url)) as typeof import('@ruled/sheet/standalone');
    document.body.replaceChildren();
    document.body.style.cssText = 'margin:0;display:flex;height:100vh';
    const place = (id: string) => {
      const el = document.createElement('div');
      el.id = id;
      el.style.cssText = 'flex:1;height:100%;min-width:0';
      document.body.append(el);
      return el;
    };
    const left = lib.mountWorkbookEditor(place('left'), { warnOnUnload: false });
    const right = lib.mountWorkbookEditor(place('right'), { warnOnUnload: false });
    const quote = (sheet: string) => `'${sheet.replace(/'/g, "''")}'`;
    // Each follows the other, leaving alone what already matches so they don't echo.
    const follow = (from: typeof left, to: typeof left) => {
      let selection = from.selection;
      let viewport = from.viewport;
      from.subscribe(() => {
        const s = from.selection;
        if (s && s !== selection) {
          selection = s;
          const t = to.selection;
          if (!t || t.sheet !== s.sheet || t.ranges.join() !== s.ranges.join()) to.goTo(`${quote(s.sheet)}!${s.ranges[0]}`, { focus: false });
        }
        const v = from.viewport;
        if (v && v !== viewport) {
          viewport = v;
          const w = to.viewport;
          if (!w || w.sheet !== v.sheet || w.cell !== v.cell) to.scrollTo(`${quote(v.sheet)}!${v.cell}`);
        }
      });
    };
    const data = Uint8Array.from(bytes);
    await left.open(data, 'list.xlsx');
    await right.open(data, 'list.xlsx');
    follow(left, right);
    follow(right, left);
    (window as unknown as { editors: object }).editors = { left, right };
  }, bytes);
  await expect(page.locator('#left .wbe-grid-scroller')).toBeVisible();
  await expect(page.locator('#right .wbe-grid-scroller')).toBeVisible();
  await expect.poll(async () => (await state(page)).right.viewport?.cell).toBe('B3');
});

test('scrolls as the other editor scrolls, past tall, hidden and collapsed rows', async ({ page }) => {
  // The handle's own scrollTo, reported back as the viewport and followed by the other editor.
  await page.evaluate(() => (window as unknown as { editors: { left: { scrollTo(ref: string): boolean } } }).editors.left.scrollTo('B120'));
  await expect.poll(async () => (await state(page)).left.viewport?.cell).toBe('B120');
  await expect.poll(async () => (await state(page)).right.viewport?.cell).toBe('B120');

  // Hidden rows and a collapsed group scroll to the first row shown after them.
  for (const [to, shown] of [
    ['A10', 'B12'],
    ['B60', 'B71'],
    ['C25', 'C25'],
  ]) {
    await page.evaluate((to) => (window as unknown as { editors: { right: { scrollTo(ref: string): boolean } } }).editors.right.scrollTo(to), to);
    await expect.poll(async () => (await state(page)).right.viewport?.cell).toBe(shown);
    await expect.poll(async () => (await state(page)).left.viewport?.cell).toBe(shown);
  }

  // Scrolled by the user, with the wheel, a little at a time and then a long way.
  const box = (await page.locator('#left .wbe-grid-scroller').boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  let was = 'C25';
  for (const dy of [120, 600, 4000]) {
    await page.mouse.wheel(0, dy);
    await expect.poll(async () => {
      const { left, right } = await state(page);
      return left.viewport?.cell === right.viewport?.cell && left.viewport?.cell !== was;
    }).toBe(true);
    was = (await state(page)).left.viewport!.cell;
  }
  const { left } = await state(page);
  expect(left.viewport!.row).toBeGreaterThan(71);
  // The frozen panes stay: the first row scrolled is below them, its number as in A1.
  expect(left.viewport).toMatchObject({ sheet: 'Long list', cell: `${String.fromCharCode(64 + left.viewport!.col)}${left.viewport!.row}` });
});

test('selects as the other editor selects, leaving the focus where it is', async ({ page }) => {
  // Click a cell in the left editor, then extend the selection with the keyboard.
  const box = (await page.locator('#left .wbe-grid-scroller').boundingBox())!;
  await page.mouse.click(box.x + 150, box.y + 70);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowRight');
  await expect.poll(async () => (await state(page)).left.selection?.ranges[0]).toMatch(/:/);
  const { left } = await state(page);
  await expect.poll(async () => (await state(page)).right.selection).toEqual(left.selection);
  // The left editor kept the focus: typing goes on there.
  expect(await page.evaluate(() => !!document.activeElement?.closest('#left'))).toBe(true);
  await page.keyboard.press('ArrowDown');
  await expect.poll(async () => (await state(page)).right.selection?.active).not.toBe(left.selection!.active);
  const after = await state(page);
  expect(after.right.selection).toEqual(after.left.selection);

  // The other way: the right editor goes to a cell far down, and the left one selects it and shows it.
  await page.evaluate(() => (window as unknown as { editors: { right: { goTo(ref: string, o?: object): void } } }).editors.right.goTo('D300', { focus: false }));
  await expect.poll(async () => (await state(page)).left.selection?.active).toBe('D300');
  expect(await page.evaluate(() => !!document.activeElement?.closest('#left'))).toBe(true);
  await expect.poll(async () => (await state(page)).left.viewport!.row).toBeGreaterThan(250);

  // Another sheet is followed too.
  await page.evaluate(() => (window as unknown as { editors: { left: { goTo(ref: string, o?: object): void } } }).editors.left.goTo('Other!B2', { focus: false }));
  await expect.poll(async () => (await state(page)).right.selection).toEqual({ sheet: 'Other', active: 'B2', ranges: ['B2'] });

  // goTo without options still takes the focus, as the name box does.
  await page.evaluate(() => (window as unknown as { editors: { right: { goTo(ref: string): void } } }).editors.right.goTo('C3'));
  await expect.poll(() => page.evaluate(() => !!document.activeElement?.closest('#right'))).toBe(true);
});
