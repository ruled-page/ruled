/**
 * Builds the sample workbooks that need Node:
 *
 * - public/samples/pivot.xlsx: shared formulas (as Excel writes them when a
 *   formula is filled down) and a pivot table. ExcelJS cannot create pivot
 *   tables, so a minimal one is added to the package by hand. It has no cached
 *   records and is marked to refresh on open, so Excel rebuilds it.
 * - public/samples/big.xlsx: 100,000 rows x 12 columns with 300,000 formulas,
 *   for measuring load and recalculation time.
 * - public/samples/lookups.xlsx: 20,000 orders looking up 20,000 products
 *   (80,000 lookups) and totals by category, for trying lookup-heavy
 *   recalculation.
 *
 * `--no-big` leaves out the two large ones; `--only=lookups` (or any other
 * sample's name) writes just that one.
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

async function lookupSample(count = 20_000) {
  const rand = rng(7);
  const CATEGORIES = ['Audio', 'Bags', 'Batteries', 'Cables', 'Cameras', 'Chargers', 'Desks', 'Drives', 'Headsets', 'Keyboards', 'Lamps', 'Monitors', 'Mice', 'Networking', 'Paper', 'Pens', 'Printers', 'Seating', 'Storage', 'Tablets'];
  const WORDS = ['Compact', 'Classic', 'Deluxe', 'Eco', 'Pro', 'Slim', 'Smart', 'Travel', 'Ultra', 'Value'];
  const sku = (i) => `P${String(i + 1).padStart(5, '0')}`;
  const products = Array.from({ length: count }, (_, i) => {
    const category = CATEGORIES[i % CATEGORIES.length];
    return { sku: sku(i), name: `${WORDS[Math.floor(rand() * WORDS.length)]} ${category.toLowerCase()} ${i + 1}`, category, price: Math.round((2 + rand() * 400) * 100) / 100, stock: Math.floor(rand() * 60) };
  });
  const bySku = new Map(products.map((p) => [p.sku, p]));
  const path = join(outDir, 'lookups.xlsx');
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path, useSharedStrings: true, useStyles: true });

  const ps = wb.addWorksheet('Products', { views: [{ state: 'frozen', ySplit: 1 }] });
  ps.columns = [
    { header: 'SKU', key: 'sku', width: 9 },
    { header: 'Name', key: 'name', width: 26 },
    { header: 'Category', key: 'category', width: 13 },
    { header: 'Price', key: 'price', width: 10, style: { numFmt: '#,##0.00' } },
    { header: 'Stock', key: 'stock', width: 8 },
  ];
  for (const p of products) ps.addRow(p).commit();
  ps.commit();

  const last = count + 1;
  const table = `Products!$A$2:$E$${last}`;
  const os = wb.addWorksheet('Orders', { views: [{ state: 'frozen', ySplit: 1 }] });
  os.columns = [
    { header: 'Order', key: 'order', width: 10 },
    { header: 'SKU', key: 'sku', width: 9 },
    { header: 'Quantity', key: 'qty', width: 9 },
    { header: 'Product', key: 'name', width: 26 },
    { header: 'Category', key: 'category', width: 13 },
    { header: 'Price', key: 'price', width: 10, style: { numFmt: '#,##0.00' } },
    { header: 'Amount', key: 'amount', width: 12, style: { numFmt: '#,##0.00' } },
    { header: 'In stock', key: 'stocked', width: 9 },
  ];
  const totals = new Map(CATEGORIES.map((c) => [c, { orders: 0, amount: 0, short: 0 }]));
  for (let i = 0; i < count; i++) {
    const n = i + 2;
    // One order in a hundred names a product that isn't in the list.
    const code = rand() < 0.01 ? `P9${String(Math.floor(rand() * 10000)).padStart(4, '0')}` : sku(Math.floor(rand() * count));
    const p = bySku.get(code);
    const qty = 1 + Math.floor(rand() * 40);
    const price = p ? p.price : 0;
    const stocked = (p ? p.stock : 0) >= qty ? 'Yes' : 'Short';
    if (p) {
      const t = totals.get(p.category);
      t.orders++;
      t.amount += qty * price;
      if (stocked === 'Short') t.short++;
    }
    os.addRow({
      order: `SO-${100001 + i}`,
      sku: code,
      qty,
      name: { formula: `IFERROR(VLOOKUP(B${n},${table},2,FALSE),"Unknown SKU")`, result: p ? p.name : 'Unknown SKU' },
      category: { formula: `IFERROR(INDEX(Products!$C$2:$C$${last},MATCH(B${n},Products!$A$2:$A$${last},0)),"")`, result: p ? p.category : '' },
      price: { formula: `IFERROR(VLOOKUP(B${n},Products!$A:$E,4,FALSE),0)`, result: price },
      amount: { formula: `C${n}*F${n}`, result: qty * price },
      stocked: { formula: `IF(IFERROR(VLOOKUP(B${n},${table},5,FALSE),0)>=C${n},"Yes","Short")`, result: stocked },
    }).commit();
  }
  os.commit();

  const ss = wb.addWorksheet('Summary');
  ss.columns = [
    { header: 'Category', key: 'category', width: 13 },
    { header: 'Orders', key: 'orders', width: 9 },
    { header: 'Amount', key: 'amount', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Short of stock', key: 'short', width: 14 },
  ];
  const orders = (col) => `Orders!$${col}$2:$${col}$${last}`;
  CATEGORIES.forEach((c, i) => {
    const n = i + 2;
    const t = totals.get(c);
    ss.addRow({
      category: c,
      orders: { formula: `COUNTIFS(${orders('E')},A${n})`, result: t.orders },
      amount: { formula: `SUMIFS(${orders('G')},${orders('E')},A${n})`, result: t.amount },
      short: { formula: `COUNTIFS(${orders('E')},A${n},${orders('H')},"Short")`, result: t.short },
    }).commit();
  });
  const end = CATEGORIES.length + 1;
  const sum = (key) => [...totals.values()].reduce((s, t) => s + t[key], 0);
  ss.addRow({
    category: 'Total',
    orders: { formula: `SUM(B2:B${end})`, result: sum('orders') },
    amount: { formula: `SUM(C2:C${end})`, result: sum('amount') },
    short: { formula: `SUM(D2:D${end})`, result: sum('short') },
  }).commit();
  ss.commit();
  await wb.commit();
  console.log(`Wrote lookups.xlsx (${count.toLocaleString()} orders and products)`);
}

/**
 * Workbooks for the example schemas in src/demo/schemas.ts. Each has a few
 * deliberate problems so validation has something to show.
 */
async function schemaSamples() {
  const money = '"$"#,##0.00';

  const priceList = async (file, categoryHeader) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Prices');
    ws.columns = [
      { header: 'SKU', width: 12 },
      { header: 'Description', width: 30 },
      { header: categoryHeader, width: 12 },
      { header: 'Unit price', width: 12, style: { numFmt: money } },
      { header: 'Price inc. tax', width: 14, style: { numFmt: money } },
      { header: 'Effective from', width: 14, style: { numFmt: 'yyyy-mm-dd' } },
      { header: 'Active', width: 8 },
    ];
    ws.getRow(1).font = { bold: true };
    const items = [
      ['SKU-1001', 'Steel bracket, small', 'Hardware', 4.5],
      ['SKU-1002', 'Steel bracket, large', 'Hardware', 7.25],
      ['SKU-1003', 'Mounting kit', 'Hardware', 19.99],
      ['SKU-1004', 'Inventory app, annual licence', 'Software', 240],
      ['SKU-1005', 'Inventory app, monthly licence', 'Software', 24],
      ['SKU-1006', 'Installation (per hour)', 'Services', 85],
      ['SKU-1003', 'Mounting kit (duplicate entry)', 'Hardware', 19.99],
      ['SKU-1008', 'Cable ties, pack of 100', 'Hardware', 6.5],
      ['SKU-1009', 'Training session', 'Services', 450],
      ['SKU-1010', '', 'Services', 120],
      ['SKU-1011', 'Wall anchors, pack of 50', 'Hardwear', 8.75],
      ['SKU-1012', 'Support plan, gold', 'Services', 1200.125],
      ['SKU-1013', 'Legacy adaptor', 'Hardware', 3.2],
      ['SKU-1014', 'Reporting add-on', 'Software', 60],
    ];
    items.forEach(([sku, desc, cat, price], i) => {
      const r = i + 2;
      ws.getCell(`A${r}`).value = sku;
      ws.getCell(`B${r}`).value = desc || null;
      ws.getCell(`C${r}`).value = cat;
      ws.getCell(`D${r}`).value = price;
      ws.getCell(`E${r}`).value = { formula: `D${r}*1.2`, result: Math.round(price * 1.2 * 1e6) / 1e6 };
      ws.getCell(`F${r}`).value = new Date(Date.UTC(2026, i % 12, 1));
      ws.getCell(`G${r}`).value = i % 5 !== 4;
    });
    const notes = wb.addWorksheet('Notes');
    notes.getCell('A1').value = 'Anything on this sheet is free text; the schema only describes the Prices sheet.';
    await wb.xlsx.writeFile(join(outDir, file));
    console.log(`Wrote ${file}`);
  };
  await priceList('price-list.xlsx', 'Category');
  await priceList('price-list-old.xlsx', 'Group');

  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Forecast');
    ws.getColumn(1).width = 14;
    const labels = ['Period', 'Revenue', 'Cost', 'Margin', 'Margin %'];
    labels.forEach((l, i) => {
      ws.getCell(i + 1, 1).value = l;
      ws.getCell(i + 1, 1).font = { bold: true };
    });
    const quarters = [
      ['2026-Q1', 120000, 70000],
      ['2026-Q2', 135000, 76000],
      ['2026-Q4', 150000, 81000],
      ['2027-Q1', 158000, -5],
    ];
    quarters.forEach(([period, revenue, cost], i) => {
      const c = i + 2;
      const col = String.fromCharCode(64 + c);
      ws.getColumn(c).width = 13;
      ws.getCell(1, c).value = period;
      ws.getCell(1, c).font = { bold: true };
      ws.getCell(2, c).value = revenue;
      ws.getCell(3, c).value = cost;
      ws.getCell(4, c).value = { formula: `${col}2-${col}3`, result: revenue - cost };
      ws.getCell(5, c).value = { formula: `IF(${col}2=0,0,${col}4/${col}2)`, result: revenue ? (revenue - cost) / revenue : 0 };
      for (const r of [2, 3, 4]) ws.getCell(r, c).numFmt = '#,##0';
      ws.getCell(5, c).numFmt = '0.0%';
    });
    await wb.xlsx.writeFile(join(outDir, 'forecast.xlsx'));
    console.log('Wrote forecast.xlsx');
  }

  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Budget');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
    const lines = {
      Salaries: [42000, 42000, 43500, 43500, 43500, 45000],
      Rent: [6500, 6500, 6500, 6500, 6500, 6500],
      Marketing: [8000, 12000, 9500, 7000, 15000, 11000],
      Travel: [2500, 3100, 6200, 1800, 2400, 4900],
      Equipment: [1200, 0, 3400, 950, 0, 2200.5],
    };
    ws.getColumn(1).width = 14;
    ws.getCell('A1').value = 'Category';
    ws.getCell('A1').font = { bold: true };
    months.forEach((m, i) => {
      ws.getCell(1, i + 2).value = m;
      ws.getCell(1, i + 2).font = { bold: true };
      ws.getColumn(i + 2).width = 11;
    });
    Object.entries(lines).forEach(([label, values], r) => {
      ws.getCell(r + 2, 1).value = label;
      values.forEach((v, i) => {
        ws.getCell(r + 2, i + 2).value = v;
        ws.getCell(r + 2, i + 2).numFmt = '#,##0.00';
      });
    });
    const total = Object.values(lines).flat().reduce((a, b) => a + b, 0);
    ws.getCell('A8').value = 'Approved by';
    ws.getCell('A9').value = 'Total';
    ws.getCell('B9').value = { formula: 'SUM(B2:G6)', result: total };
    ws.getCell('B9').numFmt = '#,##0.00';
    await wb.xlsx.writeFile(join(outDir, 'budget.xlsx'));
    console.log('Wrote budget.xlsx');
  }
}

/**
 * public/samples/orders.xlsx: Excel's own list validations, with no schema.
 * Covers an inline list, a range on another sheet, a defined name, each
 * alert style, and a non-list rule.
 */
async function validationSample() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Orders');
  const lists = wb.addWorksheet('Lists');
  const products = ['Widget', 'Gadget', 'Gizmo', 'Doohickey', 'Thingamajig'];
  const regions = ['North', 'South', 'East', 'West'];
  lists.getCell('A1').value = 'Products';
  lists.getCell('C1').value = 'Regions';
  products.forEach((p, i) => (lists.getCell(`A${i + 2}`).value = p));
  regions.forEach((r, i) => (lists.getCell(`C${i + 2}`).value = r));
  lists.getColumn(1).width = 14;
  lists.getColumn(3).width = 12;
  wb.definedNames.add(`Lists!$C$2:$C$${regions.length + 1}`, 'Regions');
  lists.getCell('I1').value = 'Max discount';
  lists.getCell('I2').value = 0.5;
  wb.definedNames.add('Lists!$I$2', 'MaxDiscount');
  // Models per product, each a named range, for a dependent dropdown with INDIRECT.
  const models = { Widget: ['W-100', 'W-200', 'W-300'], Gadget: ['G-1', 'G-2'], Gizmo: ['Z-mini', 'Z-max'] };
  Object.entries(models).forEach(([product, list], i) => {
    const col = String.fromCharCode(69 + i); // E, F, G
    lists.getCell(`${col}1`).value = product;
    list.forEach((m, r) => (lists.getCell(`${col}${r + 2}`).value = m));
    wb.definedNames.add(`Lists!$${col}$2:$${col}$${list.length + 1}`, product);
  });

  ws.columns = [
    { header: 'Order', width: 10 },
    { header: 'Customer', width: 14 },
    { header: 'Product', width: 14 },
    { header: 'Region', width: 10 },
    { header: 'Status', width: 12 },
    { header: 'Qty', width: 8 },
    { header: 'Due', width: 13, style: { numFmt: 'yyyy-mm-dd' } },
    { header: 'Ref', width: 10 },
    { header: 'Discount', width: 10, style: { numFmt: '0%' } },
    { header: 'Model', width: 10 },
    { header: 'Note', width: 22 },
    { header: 'Alt product', width: 14 },
  ];
  ws.getRow(1).font = { bold: true };
  const rand = rng(11);
  const statuses = ['Open', 'Shipped', 'Delivered', 'Cancelled'];
  const customers = ['Acme Ltd', 'Globex', 'Initech', 'Umbrella'];
  for (let i = 0; i < 12; i++) {
    const r = i + 2;
    ws.getCell(`A${r}`).value = 5001 + i;
    ws.getCell(`B${r}`).value = customers[Math.floor(rand() * customers.length)];
    ws.getCell(`C${r}`).value = products[Math.floor(rand() * products.length)];
    ws.getCell(`D${r}`).value = regions[Math.floor(rand() * regions.length)];
    ws.getCell(`E${r}`).value = statuses[Math.floor(rand() * statuses.length)];
    ws.getCell(`F${r}`).value = 1 + Math.floor(rand() * 40);
    ws.getCell(`G${r}`).value = new Date(Date.UTC(2026, 1 + (i % 9), 1 + i));
    ws.getCell(`H${r}`).value = `PO-${100 + i}`;
    ws.getCell(`I${r}`).value = (i % 4) * 0.05;
  }
  // One rule per column range: ExcelJS's per-cell merging sorts addresses as
  // text ("B10" before "B2") and writes overlapping ranges.
  // Customer: inline list, warning alert (the value can be kept anyway).
  ws.dataValidations.add('B2:B60', {
    type: 'list', allowBlank: true, formulae: [`"${customers.join(',')}"`],
    showErrorMessage: true, errorStyle: 'warning', errorTitle: 'New customer?', error: "That customer isn't on the list.",
  });
  // Product: a range on another sheet.
  ws.dataValidations.add('C2:C60', {
    type: 'list', allowBlank: true, formulae: [`Lists!$A$2:$A$${products.length + 1}`],
    showErrorMessage: true, errorTitle: 'Unknown product', error: 'Pick a product from the list on the Lists sheet.',
  });
  // Region: a defined name, information alert.
  ws.dataValidations.add('D2:D60', {
    type: 'list', allowBlank: true, formulae: ['Regions'], showErrorMessage: true, errorStyle: 'information',
    error: "That region isn't one of the usual four.",
  });
  // Status: inline list, stop alert, with an input message.
  ws.dataValidations.add('E2:E60', {
    type: 'list', allowBlank: true, formulae: [`"${statuses.join(',')}"`],
    showErrorMessage: true, errorTitle: 'Status', error: 'Choose Open, Shipped, Delivered or Cancelled.',
    showInputMessage: true, promptTitle: 'Status', prompt: 'Where the order is up to.',
  });
  // Qty: whole numbers from 1 to 1,000.
  ws.dataValidations.add('F2:F60', {
    type: 'whole', operator: 'between', allowBlank: true, formulae: [1, 1000],
    showErrorMessage: true, errorTitle: 'Quantity', error: 'Enter a whole number from 1 to 1,000.',
  });
  // Due: a date from 2026 onwards.
  ws.dataValidations.add('G2:G60', {
    type: 'date', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [new Date(Date.UTC(2026, 0, 1))],
    showErrorMessage: true, errorTitle: 'Due date', error: 'Due dates start in 2026.',
  });
  // Ref: a custom formula, relative to the first cell of the range.
  ws.dataValidations.add('H2:H60', {
    type: 'custom', allowBlank: true, formulae: ['AND(LEN(H2)<=8,ISNUMBER(SEARCH("-",H2)))'],
    showErrorMessage: true, errorTitle: 'Reference', error: 'References have a dash and at most 8 characters, like PO-123.',
  });
  // Discount: a decimal up to the MaxDiscount cell on the Lists sheet, warning only.
  ws.dataValidations.add('I2:I60', {
    type: 'decimal', operator: 'between', allowBlank: true, formulae: [0, 'MaxDiscount'],
    showErrorMessage: true, errorStyle: 'warning', errorTitle: 'Large discount', error: 'Discounts over 50% need approval.',
  });
  // Model: a dependent list, INDIRECT of the product's named range.
  ws.dataValidations.add('J2:J60', {
    type: 'list', allowBlank: true, formulae: ['INDIRECT(C2)'], showErrorMessage: true, error: 'Pick a model of the chosen product.',
  });
  // Note: text no longer than 20 characters.
  ws.dataValidations.add('K2:K60', {
    type: 'textLength', operator: 'lessThanOrEqual', allowBlank: true, formulae: [20],
    showErrorMessage: true, error: 'Keep notes to 20 characters.',
  });
  // Alt product: the product list sized with OFFSET and COUNTA, so it grows with the Lists sheet.
  ws.dataValidations.add('L2:L60', {
    type: 'list', allowBlank: true, formulae: ['OFFSET(Lists!$A$2,0,0,COUNTA(Lists!$A:$A)-1,1)'], showErrorMessage: true,
  });
  await wb.xlsx.writeFile(join(outDir, 'orders.xlsx'));
  console.log('Wrote orders.xlsx');
}

/**
 * The supplier price list for the embedded-app demo (src/embedded): a blank
 * template a supplier fills in, and a filled-in one with a problem of each
 * kind. Excel's own validation rules ride along in both, as a supplier's
 * Excel would enforce them.
 */
async function supplierSamples() {
  const make = async (file, rows) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Price list');
    ws.columns = [
      { header: 'SKU', width: 11 },
      { header: 'Description', width: 30 },
      { header: 'Category', width: 12 },
      { header: 'Unit price', width: 11, style: { numFmt: '#,##0.00' } },
      { header: 'Currency', width: 10 },
      { header: 'Min. order qty', width: 14, style: { numFmt: '#,##0' } },
      { header: 'Valid from', width: 12, style: { numFmt: 'yyyy-mm-dd' } },
      { header: 'Order value', width: 13, style: { numFmt: '#,##0.00' } },
    ];
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    rows.forEach(([sku, description, category, price, currency, moq, from], i) => {
      const r = i + 2;
      ws.getRow(r).values = [sku, description || null, category, price, currency, moq, from];
      ws.getCell(`H${r}`).value = { formula: `D${r}*F${r}`, result: price * moq };
    });
    const last = 501;
    ws.dataValidations.add(`C2:C${last}`, {
      type: 'list', allowBlank: true, formulae: ['"Hardware,Software,Services"'],
      showErrorMessage: true, errorTitle: 'Category', error: 'Choose Hardware, Software or Services.',
    });
    ws.dataValidations.add(`D2:D${last}`, {
      type: 'decimal', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [0],
      showErrorMessage: true, errorTitle: 'Unit price', error: 'Prices can’t be negative.',
    });
    ws.dataValidations.add(`E2:E${last}`, {
      type: 'list', allowBlank: true, formulae: ['"USD,EUR,GBP"'],
      showErrorMessage: true, errorTitle: 'Currency', error: 'Use USD, EUR or GBP.',
    });
    ws.dataValidations.add(`F2:F${last}`, {
      type: 'whole', operator: 'between', allowBlank: true, formulae: [1, 10000],
      showErrorMessage: true, errorTitle: 'Minimum order', error: 'Enter a whole number from 1 to 10,000.',
      showInputMessage: true, promptTitle: 'Minimum order', prompt: 'The smallest quantity you sell this in.',
    });
    ws.dataValidations.add(`G2:G${last}`, {
      type: 'date', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [new Date(Date.UTC(2026, 0, 1))],
      showErrorMessage: true, errorTitle: 'Valid from', error: 'Price lists start from 2026.',
    });
    const help = wb.addWorksheet('How to fill this in');
    help.getColumn(1).width = 90;
    [
      'One product per row on the Price list sheet.',
      'SKU: the product code from our catalogue, like SKU-1001.',
      'Unit price: the price per unit, in the currency you give. Use one currency for the whole list.',
      'Min. order qty: the smallest quantity you sell, from 1 to 10,000.',
      'Order value is worked out for you.',
    ].forEach((text, i) => (help.getCell(`A${i + 1}`).value = text));
    await wb.xlsx.writeFile(join(outDir, file));
    console.log(`Wrote ${file}`);
  };
  const from = new Date(Date.UTC(2026, 10, 1));
  await make('supplier-template.xlsx', []);
  await make('supplier-price-list.xlsx', [
    ['SKU-1001', 'Steel bracket, small', 'Hardware', 4.75, 'USD', 100, from],
    // Up 30% on the last approved price: a warning from the portal.
    ['SKU-1002', 'Steel bracket, large', 'Hardware', 9.4, 'USD', 50, from],
    ['SKU-1003', 'Mounting kit', 'Hardware', 19.99, 'USD', 10, from],
    ['SKU-1004', 'Inventory app, annual licence', 'Software', 240, 'USD', 1, from],
    ['SKU-1005', 'Inventory app, monthly licence', 'Software', 24, 'USD', 1, from],
    // Breaks Excel's own rule for the column.
    ['SKU-1006', 'Installation (per hour)', 'Services', 85, 'USD', 0, from],
    // Not one of the categories.
    ['SKU-1007', 'Cable ties, pack of 100', 'Hardwear', 6.5, 'USD', 20, from],
    // A second currency in the list.
    ['SKU-1008', 'Training session', 'Services', 450, 'EUR', 1, from],
    // Not in the portal's catalogue.
    ['SKU-1099', 'Rack shelf', 'Hardware', 55, 'USD', 5, from],
    // Three decimal places.
    ['SKU-1010', 'Support plan, gold', 'Services', 1250.125, 'USD', 1, from],
    ['SKU-1011', 'Reporting add-on', 'Software', 60, 'USD', 1, from],
  ]);
}

const only = process.argv.find((a) => a.startsWith('--only='))?.slice('--only='.length);
const large = !process.argv.includes('--no-big');
const samples = { pivot: pivotSample, schemas: schemaSamples, suppliers: supplierSamples, validation: validationSample, ...(large && { big: bigSample, lookups: lookupSample }) };
for (const [name, write] of Object.entries(samples)) if (!only || only === name) await write();
void readFileSync;
