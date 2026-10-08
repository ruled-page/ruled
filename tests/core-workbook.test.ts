import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { Finding } from '../src/core/types';
import { english, errorText } from '../src/i18n';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

const rels = (list: [id: string, type: string, target: string, external?: boolean][]) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list
    .map(([id, type, target, external]) => `<Relationship Id="${id}" Type="${type.includes('/') ? type : `${REL}/${type}`}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`)
    .join('')}</Relationships>`;

/** A package from raw parts, with the content types and root relationships every workbook has. */
function pack(files: Record<string, string>): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    '_rels/.rels': strToU8(rels([['rId1', 'officeDocument', 'xl/workbook.xml']])),
    ...Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])),
  });
}

const anchor = (body: string) =>
  `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
const chartFrame = (name: string, rid: string) =>
  anchor(`<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${name}"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`);

/**
 * A workbook with one of each thing the editor keeps but can't show: charts
 * of several kinds, a diagram, form controls, slicers, timelines, a pivot
 * table built from an external source, links to other workbooks, data
 * connections, pictures in cells, a protected structure, an old dialog sheet,
 * and two chart sheets.
 */
function kitchen(): Uint8Array {
  return pack({
    'xl/workbook.xml':
      `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><workbookProtection lockStructure="1"/><sheets>` +
      '<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="Dialog" sheetId="2" r:id="rId2"/><sheet name="Titled" sheetId="3" r:id="rId3"/><sheet name="Bare" sheetId="4" r:id="rId4"/>' +
      '</sheets></workbook>',
    'xl/_rels/workbook.xml.rels': rels([
      ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
      ['rId2', 'dialogsheet', 'dialogsheets/sheet1.xml'],
      ['rId3', 'chartsheet', 'chartsheets/sheet1.xml'],
      ['rId4', 'chartsheet', 'chartsheets/sheet2.xml'],
      ['rId5', 'externalLink', 'externalLinks/externalLink1.xml'],
      ['rId6', 'externalLink', 'externalLinks/externalLink2.xml'],
      ['rId7', 'connections', 'connections.xml'],
      ['rId8', 'http://schemas.microsoft.com/office/2007/relationships/slicerCache', 'slicerCaches/slicerCache1.xml'],
      ['rId9', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml'],
    ]),
    'xl/richData/rdrichvalue.xml': '<rvData/>',
    'xl/pivotCache/pivotCacheDefinition1.xml': `<pivotCacheDefinition xmlns="${MAIN}"><cacheSource type="external" connectionId="1"/></pivotCacheDefinition>`,
    'xl/pivotTables/pivotTable1.xml': `<pivotTableDefinition xmlns="${MAIN}" name="Sales"><location ref="H1:I4"/></pivotTableDefinition>`,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': rels([['rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml']]),
    'xl/worksheets/sheet1.xml':
      `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetPr><tabColor rgb="FF00B050"/></sheetPr><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>` +
      '<sheetFormatPr defaultColWidth="12" defaultRowHeight="18"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
      '<dataValidations count="2"><dataValidation type="none" sqref="A1"/><dataValidation type="list" sqref="A2"/></dataValidations>' +
      '<drawing r:id="rId1"/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels': rels([
      ['rId1', 'drawing', '../drawings/drawing1.xml'],
      ['rId2', 'pivotTable', '../pivotTables/pivotTable1.xml'],
      ['rId3', 'pivotTable', '../pivotTables/missing.xml'],
      ['rId4', 'table', '../tables/missing.xml'],
      ['rId5', 'ctrlProp', '../ctrlProps/ctrlProp1.xml'],
      ['rId6', 'http://schemas.microsoft.com/office/2007/relationships/slicer', '../slicers/slicer1.xml'],
      ['rId7', 'http://schemas.microsoft.com/office/2011/relationships/timeline', '../timelines/timeline1.xml'],
    ]),
    'xl/drawings/drawing1.xml':
      `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL}">` +
      chartFrame('Column chart', 'rId1') +
      chartFrame('Line chart', 'rId2') +
      chartFrame('Funnel', 'rId3') +
      chartFrame('Lost', 'rId4') +
      anchor('<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="6" name="Diagram"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData/></a:graphic></xdr:graphicFrame>') +
      '<xdr:notAnAnchor/></xdr:wsDr>',
    'xl/drawings/_rels/drawing1.xml.rels': rels([
      ['rId1', 'chart', '../charts/chart1.xml'],
      ['rId2', 'chart', '../charts/chart2.xml'],
      ['rId3', 'http://schemas.microsoft.com/office/2014/relationships/chartEx', '../charts/chartEx1.xml'],
      ['rId4', 'chart', '../charts/missing.xml'],
    ]),
    // A title that is a reference to a cell, read from its cached text.
    'xl/charts/chart1.xml':
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>' +
      '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/></c:barChart></c:plotArea></c:chart></c:chartSpace>',
    // No title: Excel shows the only series' name.
    'xl/charts/chart2.xml':
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>',
    'xl/charts/chartEx1.xml':
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>',
    'xl/dialogsheets/sheet1.xml': `<dialogsheet xmlns="${MAIN}"/>`,
    'xl/chartsheets/sheet1.xml': `<chartsheet xmlns="${MAIN}" xmlns:r="${REL}"><drawing r:id="rId1"/></chartsheet>`,
    'xl/chartsheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing2.xml']]),
    'xl/drawings/drawing2.xml': `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL}">${chartFrame('Only a title', 'rId1')}</xdr:wsDr>`,
    'xl/drawings/_rels/drawing2.xml.rels': rels([['rId1', 'chart', '../charts/chart3.xml']]),
    'xl/charts/chart3.xml': '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:p xmlns:a="a"><a:r><a:t>Forecast</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>',
    'xl/chartsheets/sheet2.xml': `<chartsheet xmlns="${MAIN}"/>`,
  });
}

const en = (f: Finding) => ({ kind: f.kind, title: english.text(f.title), detail: english.text(f.detail), ...(f.sheet ? { sheet: f.sheet } : {}), ...(f.where ? { where: f.where } : {}) });

describe('reading a workbook', () => {
  const s = WorkbookSession.open(kitchen(), 'k.xlsx');
  const findings = s.model.findings.map(en);
  const of = (kind: string) => findings.filter((f) => f.kind === kind);

  it('names the kind and title of each chart, however the file gives them', () => {
    expect(of('chart').map((f) => [f.title, f.detail.split('.')[0], f.where])).toEqual([
      ['Chart "Sales"', 'Column chart', 'B3'],
      ['Chart "Units"', 'Line chart', 'B3'],
      ['Chart', 'Funnel chart', 'B3'],
      ['Chart', 'Chart', 'B3'],
    ]);
    const chart = s.model.sheets[0].info.objects.find((o) => o.kind === 'chart');
    expect(chart).toMatchObject({ name: 'Column chart', from: { row: 2, col: 1, colOffPx: 1, rowOffPx: 0 }, to: { row: 8, col: 4 } });
  });

  it('describes chart sheets by what is known of their chart', () => {
    expect(of('chartsheet').map((f) => f.detail.split(' that')[0])).toEqual(['A chart titled "Forecast"', 'A chart']);
  });

  it('counts diagrams as shapes and lists form controls, slicers and timelines', () => {
    expect(of('shape').map((f) => f.detail)).toEqual(['1 shape. Kept in the file; shown here as placeholders.']);
    expect(of('control').map((f) => f.detail)).toEqual(['1 button or form control. Kept in the file; not shown or usable here.']);
    expect(of('slicer').map((f) => f.title)).toEqual(['Slicers', 'Timelines']);
  });

  it('says where a pivot table gets its data, and skips parts that are missing', () => {
    expect(of('pivot')).toEqual([
      {
        kind: 'pivot',
        title: 'Pivot table "Sales"',
        detail: "Built from an external source. Kept in the file; when its data is edited here, Excel refreshes it as the file opens. Its cells can't be edited here.",
        sheet: 'S',
        where: 'H1:I4',
      },
    ]);
    expect(of('table')).toEqual([]);
    expect(s.model.sheets[0].info.blocked).toMatchObject([{ kind: 'pivot', r1: 0, c1: 7, r2: 3, c2: 8 }]);
  });

  it('reports workbook-wide features', () => {
    expect(of('externalLink').map((f) => f.detail)).toEqual([
      "2 linked workbooks. Kept in the file. Formulas that use them show Excel's last value and are not recalculated here.",
    ]);
    expect(of('connection').map((f) => f.title)).toEqual(['Data connections']);
    expect(of('cellImage').map((f) => f.title)).toEqual(['Pictures or data types in cells']);
    expect(of('protection').map((f) => f.title)).toEqual(['Workbook structure is protected']);
    expect(of('macro')).toEqual([
      { kind: 'macro', title: 'Unsupported sheet "Dialog"', detail: 'A dialog or macro sheet from older versions of Excel. Kept in the file; not shown in this editor.', sheet: 'Dialog' },
    ]);
  });

  it('counts validation rules it does not enforce', () => {
    expect(of('dataValidation').map((f) => f.detail)).toEqual(["2 other rules are kept in the file but not enforced here. Pasted values aren't checked, as in Excel."]);
  });

  it('reads sheet display settings', () => {
    const info = s.model.sheets[0].info;
    expect(info).toMatchObject({ showGridLines: false, tabColor: '#00B050', defaultColWidthPx: 84, defaultRowHeightPx: 24 });
    expect(s.model.sheets.map((m) => m.info.kind)).toEqual(['worksheet', 'other', 'chartsheet', 'chartsheet']);
  });

  it('notes timeline caches and query tables without a slicer or connection part on a sheet', () => {
    const wb = pack({
      'xl/workbook.xml': `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': rels([
        ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
        ['rId2', 'http://schemas.microsoft.com/office/2011/relationships/timelineCache', 'timelineCaches/t.xml'],
      ]),
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`,
      'xl/queryTables/queryTable1.xml': '<queryTable/>',
    });
    expect(WorkbookSession.open(wb, 'w.xlsx').model.findings.map((f) => f.kind).sort()).toEqual(['connection', 'slicer']);
  });

  it('refuses a package without a workbook part', () => {
    const empty = pack({ 'docProps/app.xml': '<Properties/>' });
    let error: unknown;
    try {
      WorkbookSession.open(empty, 'x.xlsx');
    } catch (e) {
      error = e;
    }
    expect(english.text(errorText(error))).toBe('The workbook part is missing from this file.');
  });
});
