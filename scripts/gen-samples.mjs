/**
 * Builds the sample workbooks that need Node:
 *
 * - public/samples/pivot.xlsx: shared formulas (as Excel writes them when a
 *   formula is filled down) and a pivot table. ExcelJS cannot create pivot
 *   tables, so a minimal one is added to the package by hand. It has no cached
 *   records and is marked to refresh on open, so Excel rebuilds it.
 * - public/samples/big.xlsx: 100,000 rows x 12 columns with 300,000 formulas,
 *   for measuring load and recalculation time.
 *
 * public/samples/kitchen-sink.xlsx comes from scripts/gen_kitchen_sink.py.
 */
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'public', 'samples');
mkdirSync(outDir, { recursive: true });

const REGIONS = ['North', 'South', 'East', 'West'];
const PRODUCTS = { Widget: 12.5, Gadget: 30, Gizmo: 7.25 };
const CUSTOMERS = ['Acme Ltd', 'Globex', 'Initech', 'Umbrella', 'Hooli', 'Stark Industries', 'Wayne Enterprises', 'Tyrell Corp'];

/** Small deterministic PRNG so the samples are reproducible. */
function rng(seed) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 2 ** 32;
  };
}

async function pivotSample() {
  const rand = rng(7);
  const rows = [];
  for (let i = 0; i < 100; i++) {
    const region = REGIONS[Math.floor(rand() * REGIONS.length)];
    const product = Object.keys(PRODUCTS)[Math.floor(rand() * 3)];
    const units = 1 + Math.floor(rand() * 99);
    rows.push({ region, product, units, price: PRODUCTS[product] });
  }

  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.columns = [
    { header: 'Region', width: 12 },
    { header: 'Product', width: 12 },
    { header: 'Units', width: 9 },
    { header: 'Price', width: 9, style: { numFmt: '"$"#,##0.00' } },
    { header: 'Revenue', width: 12, style: { numFmt: '"$"#,##0.00' } },
  ];
  data.getRow(1).font = { bold: true };
  rows.forEach((r, i) => {
    const n = i + 2;
    data.getCell(`A${n}`).value = r.region;
    data.getCell(`B${n}`).value = r.product;
    data.getCell(`C${n}`).value = r.units;
    data.getCell(`D${n}`).value = r.price;
    // Excel stores a filled-down formula once on the first cell and points
    // the rest at it ("shared formula").
    data.getCell(`E${n}`).value =
      i === 0
        ? { formula: 'C2*D2', result: r.units * r.price, shareType: 'shared', ref: `E2:E${rows.length + 1}` }
        : { sharedFormula: 'E2', result: r.units * r.price };
  });
  const totalRow = rows.length + 3;
  data.getCell(`D${totalRow}`).value = 'Total';
  data.getCell(`D${totalRow}`).font = { bold: true };
  data.getCell(`E${totalRow}`).value = {
    formula: `SUM(E2:E${rows.length + 1})`,
    result: rows.reduce((a, r) => a + r.units * r.price, 0),
  };

  const pivot = wb.addWorksheet('Pivot');
  pivot.getColumn(1).width = 16;
  pivot.getColumn(2).width = 14;
  pivot.getCell('A1').value = 'Units by region (pivot table)';
  pivot.getCell('A1').font = { bold: true };
  const sorted = [...REGIONS].sort();
  pivot.getCell('A3').value = 'Row Labels';
  pivot.getCell('B3').value = 'Sum of Units';
  sorted.forEach((region, i) => {
    pivot.getCell(`A${4 + i}`).value = region;
    pivot.getCell(`B${4 + i}`).value = rows.filter((r) => r.region === region).reduce((a, r) => a + r.units, 0);
  });
  pivot.getCell('A8').value = 'Grand Total';
  pivot.getCell('B8').value = rows.reduce((a, r) => a + r.units, 0);
  pivot.getCell('D3').value = 'Edit units on the Data sheet; the pivot refreshes when the file is opened in Excel.';

  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  const files = unzipSync(bytes);
  const text = (p) => strFromU8(files[p]);
  const set = (p, s) => (files[p] = strToU8(s));

  const regionsInOrder = [...new Set(rows.map((r) => r.region))];
  const productsInOrder = [...new Set(rows.map((r) => r.product))];
  const num = (vals, integer) =>
    `<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${integer ? ' containsInteger="1"' : ''} minValue="${Math.min(...vals)}" maxValue="${Math.max(...vals)}"/>`;
  set(
    'xl/pivotCache/pivotCacheDefinition1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveData="0" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${rows.length}"><cacheSource type="worksheet"><worksheetSource ref="A1:E${rows.length + 1}" sheet="Data"/></cacheSource><cacheFields count="5"><cacheField name="Region" numFmtId="0"><sharedItems count="${regionsInOrder.length}">${regionsInOrder.map((r) => `<s v="${r}"/>`).join('')}</sharedItems></cacheField><cacheField name="Product" numFmtId="0"><sharedItems count="${productsInOrder.length}">${productsInOrder.map((p) => `<s v="${p}"/>`).join('')}</sharedItems></cacheField><cacheField name="Units" numFmtId="0">${num(rows.map((r) => r.units), true)}</cacheField><cacheField name="Price" numFmtId="0">${num(rows.map((r) => r.price), false)}</cacheField><cacheField name="Revenue" numFmtId="0">${num(rows.map((r) => r.units * r.price), false)}</cacheField></cacheFields></pivotCacheDefinition>`,
  );
  const items = sorted.map((r) => `<item x="${regionsInOrder.indexOf(r)}"/>`).join('');
  set(
    'xl/pivotTables/pivotTable1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="UnitsByRegion" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A3:B8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="5"><pivotField axis="axisRow" showAll="0"><items count="${sorted.length + 1}">${items}<item t="default"/></items></pivotField><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/><pivotField showAll="0"/><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="${sorted.length + 1}"><i><x/></i>${sorted
      .slice(1)
      .map((_, i) => `<i><x v="${i + 1}"/></i>`)
      .join('')}<i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Units" fld="2" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`,
  );
  set(
    'xl/pivotTables/_rels/pivotTable1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
  );

  const ct = text('[Content_Types].xml').replace(
    '</Types>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>',
  );
  set('[Content_Types].xml', ct);

  const wbRels = text('xl/_rels/workbook.xml.rels').replace(
    '</Relationships>',
    '<Relationship Id="rIdPivotCache1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>',
  );
  set('xl/_rels/workbook.xml.rels', wbRels);

  // <pivotCaches> follows calcPr (and the rarely used oleSize/customWorkbookViews).
  let wbXml = text('xl/workbook.xml');
  const calcPr = /<calcPr\b[^>]*\/>/.exec(wbXml);
  const at = calcPr ? calcPr.index + calcPr[0].length : wbXml.indexOf('</workbook>');
  wbXml = wbXml.slice(0, at) + '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches>' + wbXml.slice(at);
  if (!/xmlns:r=/.test(wbXml.slice(0, 400))) {
    wbXml = wbXml.replace('<workbook ', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  }
  set('xl/workbook.xml', wbXml);

  const sheetRelsPath = 'xl/worksheets/_rels/sheet2.xml.rels';
  const rel = '<Relationship Id="rIdPivot1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>';
  set(
    sheetRelsPath,
    files[sheetRelsPath]
      ? text(sheetRelsPath).replace('</Relationships>', rel + '</Relationships>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`,
  );

  const ordered = { '[Content_Types].xml': files['[Content_Types].xml'] };
  for (const [k, v] of Object.entries(files)) if (k !== '[Content_Types].xml') ordered[k] = v;
  writeFileSync(join(outDir, 'pivot.xlsx'), zipSync(ordered));
  console.log('Wrote pivot.xlsx');
}

async function bigSample(rowCount = 100_000) {
  const rand = rng(99);
  const path = join(outDir, 'big.xlsx');
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path, useSharedStrings: true, useStyles: true });
  const ws = wb.addWorksheet('Transactions', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: 'ID', key: 'id', width: 9 },
    { header: 'Date', key: 'date', width: 12, style: { numFmt: 'yyyy-mm-dd' } },
    { header: 'Region', key: 'region', width: 10 },
    { header: 'Product', key: 'product', width: 10 },
    { header: 'Customer', key: 'customer', width: 18 },
    { header: 'Units', key: 'units', width: 8 },
    { header: 'Price', key: 'price', width: 9, style: { numFmt: '#,##0.00' } },
    { header: 'Revenue', key: 'revenue', width: 12, style: { numFmt: '#,##0.00' } },
    { header: 'Cost', key: 'cost', width: 12, style: { numFmt: '#,##0.00' } },
    { header: 'Margin', key: 'margin', width: 12, style: { numFmt: '#,##0.00' } },
    { header: 'Margin %', key: 'pct', width: 10, style: { numFmt: '0.0%' } },
    { header: 'Notes', key: 'notes', width: 24 },
  ];
  const totals = { units: 0, revenue: 0, cost: 0, byRegion: Object.fromEntries(REGIONS.map((r) => [r, 0])) };
  const start = Date.UTC(2025, 0, 1);
  for (let i = 0; i < rowCount; i++) {
    const n = i + 2;
    const product = Object.keys(PRODUCTS)[Math.floor(rand() * 3)];
    const region = REGIONS[Math.floor(rand() * 4)];
    const units = 1 + Math.floor(rand() * 250);
    const price = PRODUCTS[product];
    const revenue = units * price;
    const cost = Math.round(revenue * (0.45 + rand() * 0.3) * 100) / 100;
    totals.units += units;
    totals.revenue += revenue;
    totals.cost += cost;
    totals.byRegion[region] += revenue;
    ws.addRow({
      id: 100000 + i,
      date: new Date(start + Math.floor(i / (rowCount / 365)) * 86_400_000),
      region,
      product,
      customer: CUSTOMERS[Math.floor(rand() * CUSTOMERS.length)],
      units,
      price,
      revenue: { formula: `F${n}*G${n}`, result: revenue },
      cost,
      margin: { formula: `H${n}-I${n}`, result: revenue - cost },
      pct: { formula: `IF(H${n}=0,0,J${n}/H${n})`, result: revenue ? (revenue - cost) / revenue : 0 },
      notes: rand() < 0.1 ? 'Follow up with account manager' : null,
    }).commit();
  }
  ws.commit();

  const summary = wb.addWorksheet('Totals');
  summary.columns = [{ width: 22 }, { width: 16 }];
  const last = rowCount + 1;
  const add = (label, formula, result, numFmt) => {
    const row = summary.addRow([label, { formula, result }]);
    if (numFmt) row.getCell(2).numFmt = numFmt;
    row.commit();
  };
  add('Total units', `SUM(Transactions!F2:F${last})`, totals.units);
  add('Total revenue', `SUM(Transactions!H2:H${last})`, totals.revenue, '#,##0.00');
  add('Total cost', `SUM(Transactions!I2:I${last})`, totals.cost, '#,##0.00');
  add('Margin', 'B2-B3', totals.revenue - totals.cost, '#,##0.00');
  for (const r of REGIONS) {
    add(`Revenue ${r}`, `SUMIFS(Transactions!H2:H${last},Transactions!C2:C${last},"${r}")`, totals.byRegion[r], '#,##0.00');
  }
  summary.commit();
  await wb.commit();
  console.log(`Wrote big.xlsx (${rowCount.toLocaleString()} rows)`);
}

await pivotSample();
if (!process.argv.includes('--no-big')) await bigSample();
void readFileSync;
