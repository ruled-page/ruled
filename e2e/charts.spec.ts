import { expect, test, type Locator } from '@playwright/test';
import { openSample, setCell } from './helpers';

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
  await page.getByRole('navigation', { name: 'Sheets' }).getByRole('button', { name: 'Monthly Chart' }).click();
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
  await page.goto('http://localhost:4174/examples/html/');
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await page.evaluate(() => {
    const editor = document.createElement('workbook-editor');
    editor.options = { warnOnUnload: false, features: { charts: false } };
    editor.setAttribute('src', '/samples/kitchen-sink.xlsx');
    document.getElementById('editor')!.replaceWith(editor);
  });
  await expect(page.getByRole('button', { name: 'Monthly Chart' })).toBeVisible();

  const placeholder = page.locator('.wbe-object-chart').filter({ hasText: 'Share of revenue' });
  await expect(placeholder).toBeVisible();
  await expect(placeholder).toContainText('Kept in file, not shown here');
  await expect(page.locator('.wbe-objects-overlay canvas')).toHaveCount(0);
  await expect(page.getByRole('img', { name: /chart/ })).toHaveCount(0);

  await page.getByRole('button', { name: 'Monthly Chart' }).click();
  await expect(page.getByText('Line chart “Monthly revenue 2026” fills this whole sheet.')).toBeVisible();
  expect(chartCode).toEqual([]);
});
