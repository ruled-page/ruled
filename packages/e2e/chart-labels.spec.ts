import { readFileSync } from 'node:fs';
import { expect, test, type Page } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openFile, setCell } from './helpers';

/**
 * The demo's kitchen sink with data labels: on its Summary sheet the column
 * chart's values (as #,##0) at the bars' ends and the pie's categories and
 * percentages; on its Notes sheet a waterfall of the months in B9:F10
 * labelled with each month and its value.
 */
function labelledWorkbook(): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync(new URL('../../examples/demo/public/samples/kitchen-sink.xlsx', import.meta.url))));
  const patch = (path: string, fn: (xml: string) => string) => (files[path] = strToU8(fn(strFromU8(files[path]))));
  patch('xl/charts/chart1.xml', (x) => x.replace('<c:cat>', '<c:dLbls><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls><c:cat>'));
  patch('xl/charts/chart2.xml', (x) => x.replace('<c:cat>', '<c:dLbls><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls><c:cat>'));
  patch('xl/drawings/drawing2.xml', (x) =>
    x.replace(
      '</xdr:wsDr>',
      '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>26</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="201" name="Waterfall"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdW"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
    ),
  );
  patch('xl/drawings/_rels/drawing2.xml.rels', (x) =>
    x.replace('</Relationships>', '<Relationship Id="rIdW" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/></Relationships>'),
  );
  files['xl/charts/chartEx1.xml'] = strToU8(
    '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0">' +
      '<cx:strDim type="cat"><cx:f>Notes!$B$9:$F$9</cx:f><cx:lvl ptCount="5"><cx:pt idx="0">Jan</cx:pt><cx:pt idx="1">Feb</cx:pt><cx:pt idx="2">Mar</cx:pt><cx:pt idx="3">Apr</cx:pt><cx:pt idx="4">May</cx:pt></cx:lvl></cx:strDim>' +
      '<cx:numDim type="val"><cx:f>Notes!$B$10:$F$10</cx:f><cx:lvl ptCount="5" formatCode="General"><cx:pt idx="0">12</cx:pt><cx:pt idx="1">18</cx:pt><cx:pt idx="2">9</cx:pt><cx:pt idx="3">24</cx:pt><cx:pt idx="4">30</cx:pt></cx:lvl></cx:numDim>' +
      '</cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Cash flow</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall">' +
      '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="1" value="1"/></cx:dataLabels><cx:dataId val="0"/></cx:series></cx:plotAreaRegion>' +
      '<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>',
  );
  return zipSync(files);
}

/** The texts written on canvases so far, each once: the charts draw on canvases, so their labels can't be read from the page. */
const written = (page: Page) => page.evaluate(() => [...new Set((window as unknown as { written: string[] }).written)].sort());
/** Starts the record of what is written afresh. */
const forget = (page: Page) => page.evaluate(() => void ((window as unknown as { written: string[] }).written.length = 0));

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    const w = window as unknown as { written: string[] };
    w.written = [];
    const fillText = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text, ...rest) {
      w.written.push(String(text));
      return fillText.call(this, text, ...rest);
    };
  });
});

test("draws charts' data labels as the file sets them, and redraws them as the cells change", async ({ page }) => {
  await openFile(page, 'labels.xlsx', labelledWorkbook());
  // Summary: the columns' values as #,##0 (North's total is 36,359.50), and each slice's category and share.
  await expect.poll(() => written(page)).toContain('36,360');

  await page.getByRole('navigation', { name: 'Sheets' }).getByRole('tab', { name: 'Notes' }).click();
  await expect(page.getByRole('img', { name: /Cash flow/ })).toBeVisible();
  await expect.poll(async () => (await written(page)).filter((s) => /^(Jan|Feb|Mar|Apr|May), /.test(s))).toEqual(['Apr, 24', 'Feb, 18', 'Jan, 12', 'Mar, 9', 'May, 30']);
  await forget(page);
  await setCell(page, 'C10', '-40');
  await expect.poll(() => written(page)).toContain('Feb, -40');
});

test("a pie's labels show each slice's category and share", async ({ page }) => {
  await openFile(page, 'labels.xlsx', labelledWorkbook());
  await expect(page.getByRole('img', { name: /Share of revenue/ })).toBeVisible();
  // The shares of $157,925.25 in the sample: 23%, 27%, 27% and 23%.
  await expect.poll(async () => (await written(page)).filter((s) => /^(North|South|East|West), \d+%$/.test(s))).toEqual(['East, 27%', 'North, 23%', 'South, 27%', 'West, 23%']);
});
