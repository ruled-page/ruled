import { expect, test, type Locator } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { unzipSync } from 'fflate';
import { goTo, openFile, openSample, ordersWorkbook, serverUrl, setCell } from './helpers';

/** How many pixels a chart's canvas has drawn on. */
const inked = (chart: Locator) =>
  chart.locator('canvas').evaluate((canvas: HTMLCanvasElement) => {
    const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
    let n = 0;
    for (let i = 3; i < data.length; i += 4) if (data[i]) n++;
    return n;
  });
const drawing = (chart: Locator) => chart.locator('canvas').evaluate((canvas: HTMLCanvasElement) => canvas.toDataURL());

test('draws charts from the cells, and redraws them as the cells change', async ({ page }) => {
  const chartCode: string[] = [];
  page.on('request', (r) => {
    if (/ChartView/.test(r.url())) chartCode.push(r.url());
  });
  // A workbook without charts doesn't load the chart library.
  await openSample(page, 'Price list');
  expect(chartCode).toEqual([]);

  await openSample(page, 'Sales report');
  const chart = page.getByRole('img', { name: 'Column chart “Revenue by region”' });
  await expect(chart).toBeVisible();
  expect(chartCode).toHaveLength(1);
  await expect.poll(() => inked(chart)).toBeGreaterThan(1000);
  await expect(chart).toHaveAccessibleDescription(
    '1 series. Total: North $36,359.50, South $42,651.25, East $42,312.25, and West $36,602.25.',
  );
  const before = await drawing(chart);

  // The chart plots Summary!F5:F8 against the regions in A5:A8; F5 adds up Data's sales for the region named in A5.
  await setCell(page, 'A5', 'Nord');
  await expect(chart).toHaveAccessibleDescription('1 series. Total: Nord $0.00, South $42,651.25, East $42,312.25, and West $36,602.25.');
  await expect.poll(() => drawing(chart)).not.toBe(before);
  await expect(page.getByRole('img', { name: 'Pie chart “Share of revenue”' })).toHaveAccessibleDescription(/^1 series\. Share: Nord 0\.0%, South 35\.1%/);

  // A chart sheet's chart fills the sheet.
  await page.getByRole('tab', { name: 'Monthly Chart' }).click();
  const line = page.getByRole('img', { name: 'Line chart “Monthly revenue 2026”' });
  await expect(line).toBeVisible();
  await expect.poll(() => inked(line)).toBeGreaterThan(1000);
  await expect(line).toHaveAccessibleDescription(/^1 series\. Monthly revenue: Jan 2026 \$12,058\.50, Feb 2026 \$12,862\.25/);
});

/** The <workbook-editor> element of the HTML host page, given options and a workbook. */
test('shows placeholders, and never loads the chart library, when the host turns charts off', async ({ page }) => {
  const chartCode: string[] = [];
  page.on('request', (r) => {
    if (/ChartView/.test(r.url())) chartCode.push(r.url());
  });
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await page.evaluate(() => {
    const editor = document.createElement('workbook-editor');
    editor.options = { warnOnUnload: false, features: { charts: false } };
    editor.setAttribute('src', '/examples/demo/public/samples/kitchen-sink.xlsx');
    document.getElementById('editor')!.replaceWith(editor);
  });
  await expect(page.getByRole('tab', { name: 'Monthly Chart' })).toBeVisible();

  const placeholder = page.locator('.wbe-object-chart').filter({ hasText: 'Share of revenue' });
  await expect(placeholder).toBeVisible();
  await expect(placeholder).toContainText('Kept in file, not shown here');
  await expect(page.locator('.wbe-objects-overlay canvas')).toHaveCount(0);
  await expect(page.getByRole('img', { name: /chart/ })).toHaveCount(0);

  await page.getByRole('tab', { name: 'Monthly Chart' }).click();
  await expect(page.getByText('Line chart “Monthly revenue 2026” fills this whole sheet.')).toBeVisible();
  expect(chartCode).toEqual([]);
});

test('keeps a chart where it is anchored as the sheet scrolls past a tall row, and out of the scrollbars', async ({ page }) => {
  await openSample(page, 'Sales report');
  const pie = page.getByRole('img', { name: 'Pie chart “Share of revenue”' });
  await expect(pie).toBeVisible();
  // Row 19, where the pie chart starts, made 200 pixels tall.
  await goTo(page, 'A19');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Rows' }).focus();
  await page.keyboard.press('Enter');
  await page.getByRole('menuitem', { name: 'Row height…' }).focus();
  await page.keyboard.press('Enter');
  await page.getByRole('dialog', { name: 'Height of row 19' }).getByRole('textbox').fill('150');
  await page.keyboard.press('Enter');
  const top = async () => (await pie.boundingBox())!.y;
  await expect.poll(top).toBeGreaterThan(0);
  const before = await top();

  // Scrolled until row 19 has gone above the panes' join, the chart goes up with it, as far as the
  // sheet went: browsers scroll a turn of the wheel by different amounts.
  const scrolled = () => page.locator('.dvn-scroller').evaluateAll((all) => Math.max(...all.map((s) => s.scrollTop)));
  const start = await scrolled();
  const grid = page.locator('canvas[data-testid="data-grid-canvas"]').last();
  const box = (await grid.boundingBox())!;
  await page.mouse.move(box.x + 200, box.y + box.height / 2);
  for (let turns = 0; turns < 20 && (await scrolled()) - start < 500; turns++) await page.mouse.wheel(0, 100);
  await expect.poll(async () => Math.round(before - (await top())) - ((await scrolled()) - start)).toBe(0);
  expect((await scrolled()) - start).toBeGreaterThanOrEqual(500);

  // Charts stop short of each pane's scrollbars.
  const overlaps = await page.locator('.wbe-objects-overlay').evaluateAll((overlays) =>
    overlays.filter((o) => {
      const scroller = o.parentElement!.querySelector<HTMLElement>('.dvn-scroller')!;
      const s = scroller.getBoundingClientRect();
      const r = o.getBoundingClientRect();
      return r.bottom > s.top + scroller.clientHeight + 0.5 || r.right > s.left + scroller.clientWidth + 0.5;
    }).length,
  );
  expect(overlaps).toBe(0);
});

test('inserts a chart of the selection from the cell menu or Alt+F1, saves it as a chart part, and undoes it', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await goTo(page, 'A1:C3');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Insert chart' }).focus();
  await page.keyboard.press('Enter');
  await page.getByRole('menuitem', { name: 'Column', exact: true }).focus();
  await page.keyboard.press('Enter');
  // Two series, Qty and Price, over the items: Excel's placeholder title and a drawing of them.
  const chart = page.getByRole('img', { name: 'Column chart “Chart Title”' });
  await expect(chart).toBeVisible();
  await expect(chart).toHaveAccessibleDescription(/^2 series\. Qty: Bolts 2 and Nuts 4\./);
  await expect.poll(() => inked(chart)).toBeGreaterThan(1000);

  // Alt+F1 charts one column, named by its heading.
  await goTo(page, 'D1:D3');
  await page.keyboard.press('Alt+F1');
  await expect(page.getByRole('img', { name: 'Column chart “Total”' })).toBeVisible();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const files = unzipSync(new Uint8Array(await readFile(await (await download).path())));
  expect(Object.keys(files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).toHaveLength(2);
  expect(Object.keys(files)).toContain('xl/drawings/drawing1.xml');

  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(page.getByRole('img', { name: 'Column chart “Total”' })).toHaveCount(0);
  await expect(chart).toBeVisible();
});

test('selects a chart by clicking it, moves and resizes it by dragging, and deletes it, each a step undone on its own', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await goTo(page, 'D1:D3');
  await page.keyboard.press('Alt+F1');
  const chart = page.getByRole('img', { name: 'Column chart “Total”' });
  await expect(chart).toBeVisible();
  const box = async () => (await chart.boundingBox())!;
  const start = await box();

  // A click selects it: a frame with handles, holding the keyboard.
  await page.mouse.click(start.x + start.width / 2, start.y + start.height / 2);
  const frame = page.getByRole('group', { name: 'Chart 1, selected' });
  await expect(frame).toBeFocused();

  // Dragged 64 pixels right and 40 down.
  await page.mouse.move(start.x + 100, start.y + 100);
  await page.mouse.down();
  await page.mouse.move(start.x + 140, start.y + 120, { steps: 4 });
  await page.mouse.move(start.x + 164, start.y + 140, { steps: 4 });
  await page.mouse.up();
  // Dropped, it stays where it was let go rather than flashing back while the move is saved.
  expect(Math.round((await box()).x - start.x)).toBe(64);
  await expect.poll(async () => [Math.round((await box()).x - start.x), Math.round((await box()).y - start.y)]).toEqual([64, 40]);

  // Its bottom-right handle dragged 50 pixels further out.
  const moved = await box();
  await page.mouse.move(moved.x + moved.width, moved.y + moved.height);
  await page.mouse.down();
  await page.mouse.move(moved.x + moved.width + 50, moved.y + moved.height + 50, { steps: 5 });
  await page.mouse.up();
  expect(Math.round((await box()).width - moved.width)).toBe(50);
  await expect.poll(async () => Math.round((await box()).width - moved.width)).toBe(50);

  // Arrow keys nudge it; Delete removes it, and back to the cells.
  await page.keyboard.press('ArrowLeft');
  await expect.poll(async () => Math.round((await box()).x - moved.x)).toBe(-1);
  await page.keyboard.press('Delete');
  await expect(chart).toHaveCount(0);
  await expect(page.locator('canvas[data-testid="data-grid-canvas"]:focus-within')).toHaveCount(1);
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(chart).toBeVisible();

  // Ctrl+Alt+5 selects it from the keyboard; Esc goes back to the cells.
  await goTo(page, 'A1');
  await page.keyboard.press('Control+Alt+5');
  await expect(frame).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(frame).toHaveCount(0);
});

test('changes a chart in its dialog, opened by double-clicking it or from its menu: type, title, legend and data', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await goTo(page, 'A1:C3');
  await page.keyboard.press('Alt+F1');
  const column = page.getByRole('img', { name: 'Column chart “Chart Title”' });
  const b = (await column.boundingBox())!;
  await page.mouse.dblclick(b.x + b.width / 2, b.y + b.height / 2);
  const dialog = page.getByRole('dialog', { name: 'Edit chart' });
  const type = dialog.getByRole('combobox', { name: 'Chart type' });
  await expect(type).toBeFocused();
  await expect(dialog.getByRole('textbox', { name: 'Data range' })).toHaveValue('=Lines!$A$1:$C$3');
  await expect(dialog.getByRole('textbox', { name: 'Series 2: Values' })).toHaveValue('=Lines!$C$2:$C$3');
  await type.selectOption('line');
  await dialog.getByRole('textbox', { name: 'Title' }).fill('Orders');
  await dialog.getByRole('combobox', { name: 'Legend' }).selectOption('top');
  await dialog.getByRole('button', { name: 'OK' }).click();
  const line = page.getByRole('img', { name: 'Line chart “Orders”' });
  await expect(line).toHaveAccessibleDescription(/^2 series\. Qty: Bolts 2 and Nuts 4\./);

  // From its menu, the series along the rows instead.
  const lb = (await line.boundingBox())!;
  await page.mouse.click(lb.x + 60, lb.y + 60, { button: 'right' });
  await page.getByRole('menuitem', { name: 'Edit chart…' }).click();
  await dialog.getByRole('combobox', { name: 'Series in' }).selectOption('rows');
  await expect(dialog.getByRole('textbox', { name: 'Series 1: Values' })).toBeDisabled();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(line).toHaveAccessibleDescription(/^2 series\. Bolts: Qty 2 and Price 3\./);

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const files = unzipSync(new Uint8Array(await readFile(await (await download).path())));
  const part = new TextDecoder().decode(files['xl/charts/chart1.xml']);
  expect(part).toContain('<c:lineChart>');
  expect(part).toContain('<a:t>Orders</a:t>');
  expect(part).toContain('<c:f>Lines!$B$2:$C$2</c:f>');

  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(line).toHaveAccessibleDescription(/^2 series\. Qty:/);
});

test('draws radar, bubble and stock charts from the chart gallery, and says what they show', async ({ page }) => {
  await openSample(page, 'Chart gallery');
  for (const name of [
    'Radar chart “Skills (radar)”',
    'Radar chart “Skills (filled radar)”',
    'Bubble chart “Price, margin and units (bubble)”',
    'Stock chart “High, low, close (stock)”',
    'Stock chart “Open, high, low, close (stock)”',
  ]) {
    const chart = page.getByRole('img', { name });
    await expect(chart).toHaveCount(1);
    await expect.poll(() => inked(chart)).toBeGreaterThan(500);
  }
  await expect(page.getByRole('img', { name: 'Stock chart “Open, high, low, close (stock)”' })).toHaveAccessibleDescription(/^4 series\. Open: 1 Oct 41\.2, 2 Oct 42\.5/);
  await expect(page.getByRole('img', { name: 'Bubble chart “Price, margin and units (bubble)”' })).toHaveAccessibleDescription(/^1 series\. Products: x 12\.5, y 0\.32, size 1200/);
});
