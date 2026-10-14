import { readFileSync } from 'node:fs';
import { expect, test, type Locator } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openFile, setCell } from './helpers';

/** How many pixels a chart's canvas has drawn on. */
const inked = (chart: Locator) =>
  chart.locator('canvas').evaluate((canvas: HTMLCanvasElement) => {
    const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
    let n = 0;
    for (let i = 3; i < data.length; i += 4) if (data[i]) n++;
    return n;
  });
const drawing = (chart: Locator) => chart.locator('canvas').evaluate((canvas: HTMLCanvasElement) => canvas.toDataURL());

/**
 * The demo's kitchen sink with, on its Notes sheet, a waterfall ("Cash
 * flow", May a total) and a funnel ("Pipeline") of the months in B9:F10,
 * in Excel 2016's chart format.
 */
function stepsWorkbook(): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync(new URL('../../examples/demo/public/samples/kitchen-sink.xlsx', import.meta.url))));
  const frame = (id: number, rid: string, c1: number, c2: number) =>
    `<xdr:twoCellAnchor><xdr:from><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>26</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  const part = (layout: string, title: string, extra = '') =>
    '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0">' +
    '<cx:strDim type="cat"><cx:f>Notes!$B$9:$F$9</cx:f><cx:lvl ptCount="5"><cx:pt idx="0">Jan</cx:pt><cx:pt idx="1">Feb</cx:pt><cx:pt idx="2">Mar</cx:pt><cx:pt idx="3">Apr</cx:pt><cx:pt idx="4">May</cx:pt></cx:lvl></cx:strDim>' +
    '<cx:numDim type="val"><cx:f>Notes!$B$10:$F$10</cx:f><cx:lvl ptCount="5" formatCode="General"><cx:pt idx="0">12</cx:pt><cx:pt idx="1">18</cx:pt><cx:pt idx="2">9</cx:pt><cx:pt idx="3">24</cx:pt><cx:pt idx="4">30</cx:pt></cx:lvl></cx:numDim>' +
    `</cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>${title}</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataId val="0"/>${extra}</cx:series></cx:plotAreaRegion>` +
    '<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>';
  const d = 'xl/drawings/drawing2.xml';
  files[d] = strToU8(strFromU8(files[d]).replace('</xdr:wsDr>', frame(201, 'rIdW', 1, 5) + frame(202, 'rIdF', 5, 9) + '</xdr:wsDr>'));
  const r = 'xl/drawings/_rels/drawing2.xml.rels';
  const ex = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
  files[r] = strToU8(
    strFromU8(files[r]).replace('</Relationships>', `<Relationship Id="rIdW" Type="${ex}" Target="../charts/chartEx1.xml"/><Relationship Id="rIdF" Type="${ex}" Target="../charts/chartEx2.xml"/></Relationships>`),
  );
  files['xl/charts/chartEx1.xml'] = strToU8(part('waterfall', 'Cash flow', '<cx:layoutPr><cx:subtotals><cx:idx val="4"/></cx:subtotals></cx:layoutPr>'));
  files['xl/charts/chartEx2.xml'] = strToU8(part('funnel', 'Pipeline'));
  return zipSync(files);
}

test("draws Excel 2016's waterfall and funnel charts from the cells, redraws them as the cells change, and doesn't offer to edit them", async ({ page }) => {
  await openFile(page, 'steps.xlsx', stepsWorkbook());
  await page.getByRole('navigation', { name: 'Sheets' }).getByRole('tab', { name: 'Notes' }).click();
  const waterfall = page.getByRole('img', { name: /Cash flow/ });
  const funnel = page.getByRole('img', { name: /Pipeline/ });
  await expect(waterfall).toBeVisible();
  await expect(funnel).toBeVisible();
  await expect.poll(() => inked(waterfall)).toBeGreaterThan(1000);
  await expect.poll(() => inked(funnel)).toBeGreaterThan(1000);
  // No placeholders for them.
  await expect(page.locator('.wbe-object-chart:not(.wbe-drawn)')).toHaveCount(0);

  const [w, f] = [await drawing(waterfall), await drawing(funnel)];
  await setCell(page, 'Notes!C10', '-40');
  await expect.poll(() => drawing(waterfall)).not.toBe(w);
  await expect.poll(() => drawing(funnel)).not.toBe(f);

  // Neither opens Excel's chart dialog, from a double-click or the menu.
  const box = (await waterfall.boundingBox())!;
  await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
  await expect(page.getByRole('dialog', { name: 'Edit chart' })).toHaveCount(0);
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' });
  await expect(page.getByRole('menuitem', { name: 'Delete' })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: 'Edit chart…' })).toHaveCount(0);
});
