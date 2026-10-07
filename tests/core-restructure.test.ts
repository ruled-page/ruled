import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { StructuralOp } from '../src/core/refTransform';
import { english, errorText } from '../src/i18n';
import { loadedFormula, loadedValue, parts, sample, sheetIndex } from './helpers';

/** What a call throws, in English. */
function thrown(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (e) {
    return english.text(errorText(e));
  }
  return undefined;
}

/** Rewrites parts of a package; a string adds or replaces the part. */
function patch(bytes: Uint8Array, edits: Record<string, string | ((xml: string) => string)>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [path, edit] of Object.entries(edits)) {
    files[path] = strToU8(typeof edit === 'string' ? edit : edit(strFromU8(files[path])));
  }
  return zipSync(files);
}

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

/**
 * Data!A1:A6 = 1..6, columns B and D widened, a page break after row 3 and
 * one after column B, table T at C1:E3 (X, Y, Z), a picture over F5:H9, the
 * selection on A5, a threaded comment on A5, a validation rule (A1:A6) and a
 * formatting rule (A2:A6) written the Excel 2010 way, and a protected range
 * and an ignored error on A1:A6.
 * Calc!A1:A3 is one shared formula, Data!A1*2 copied down.
 * Workbook names: Here (local to Data, unqualified) = $A$4.
 * Empty is a sheet with no cells.
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  const calc = wb.addWorksheet('Calc');
  wb.addWorksheet('Empty');
  for (let i = 1; i <= 6; i++) data.getCell(`A${i}`).value = i;
  data.getColumn(2).width = 20;
  data.getColumn(4).width = 30;
  data.getRow(3).addPageBreak();
  data.addTable({ name: 'T', ref: 'C1', headerRow: true, columns: [{ name: 'X' }, { name: 'Y' }, { name: 'Z' }], rows: [[1, 2, 3], [4, 5, 6]] });
  // ExcelJS reads these but its types want more.
  data.addImage(wb.addImage({ base64: PNG, extension: 'png' }), { tl: { col: 5, row: 4 }, br: { col: 7, row: 8 } } as unknown as ExcelJS.ImageRange);
  calc.getCell('A1').value = { formula: 'Data!A1*2', shareType: 'shared', ref: 'A1:A3', result: 2 } as ExcelJS.CellValue;
  calc.getCell('A2').value = { sharedFormula: 'A1', result: 4 };
  calc.getCell('A3').value = { sharedFormula: 'A1', result: 6 };
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  const thread =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
    '<threadedComment ref="A5" dT="2026-01-01T00:00:00.00" id="{00000000-0000-0000-0000-000000000001}"><text>Check this</text></threadedComment></ThreadedComments>';
  return patch(bytes, {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml
        .replace('<sheetFormatPr', '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A5" sqref="A5"/></sheetView></sheetViews><sheetFormatPr')
        .replace('</sheetData>', '</sheetData><protectedRanges><protectedRange name="P" sqref="A1:A6"/></protectedRanges>')
        .replace('</rowBreaks>', '</rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks>')
        .replace(
          '</worksheet>',
          '<ignoredErrors><ignoredError sqref="A1:A6" numberStoredAsText="1"/></ignoredErrors>' +
            '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
            '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1">' +
            '<x14:formula1><xm:f>Calc!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>A1:A6</xm:sqref></x14:dataValidation></x14:dataValidations></ext>' +
            '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>' +
            '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="1" id="{00000000-0000-0000-0000-000000000002}">' +
            '<xm:f>Calc!$A$1&gt;1</xm:f><x14:dxf><font><b/></font></x14:dxf></x14:cfRule><xm:sqref>A2:A6</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>',
        ),
    'xl/worksheets/_rels/sheet1.xml.rels': (xml) =>
      xml.replace(
        '</Relationships>',
        '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
      ),
    'xl/threadedComments/threadedComment1.xml': thread,
    'xl/workbook.xml': (xml) => xml.replace('</sheets>', '</sheets><definedNames><definedName name="Here" localSheetId="0">$A$4</definedName></definedNames>'),
  });
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

const DATA = 0;
const CALC = 1;
const EMPTY = 2;

function after(op: StructuralOp, from = bytes, name = 'rich.xlsx') {
  const s = WorkbookSession.open(from, name);
  const next = s.restructure(op);
  const saved = next.save();
  return { next, file: WorkbookSession.open(saved, name), xml: parts(saved) };
}

const sheetXml = (xml: Record<string, string>, n: number) => xml[`xl/worksheets/sheet${n}.xml`];
const anchors = (drawing: string, tag: 'row' | 'col') => [...drawing.matchAll(new RegExp(`<xdr:${tag}>(\\d+)</xdr:${tag}>`, 'g'))].map((m) => Number(m[1]));

describe('inserting rows', () => {
  it('moves page breaks, the selection, local names, threaded comments, pictures and protected ranges', () => {
    const { xml } = after({ type: 'insertRows', sheet: DATA, at: 1, count: 2 });
    const sheet = sheetXml(xml, 1);
    expect(sheet).toMatch(/<rowBreaks count="1" manualBreakCount="1"><brk id="5" /);
    // Column breaks are left alone by row changes.
    expect(sheet).toContain('<brk id="2" max="1048575"');
    expect(sheet).toContain('<selection activeCell="A7" sqref="A7"/>');
    expect(sheet).toContain('<protectedRange name="P" sqref="A1:A8"/>');
    expect(sheet).toContain('<ignoredError sqref="A1:A8"');
    // Each moves once, though Excel 2010 rules keep their ranges in their own <sqref> elements.
    expect(sheet).toContain('<xm:sqref>A1:A8</xm:sqref>');
    expect(sheet).toContain('<xm:sqref>A4:A8</xm:sqref>');
    expect(xml['xl/workbook.xml']).toContain('<definedName name="Here" localSheetId="0">$A$6</definedName>');
    expect(xml['xl/threadedComments/threadedComment1.xml']).toContain('ref="A7"');
    expect(anchors(xml['xl/drawings/drawing1.xml'], 'row')).toEqual([6, 10]);
    expect(anchors(xml['xl/drawings/drawing1.xml'], 'col')).toEqual([5, 7]);
  });

  it('writes out every member of a shared formula on another sheet that points at the moved cells', () => {
    const { file, xml } = after({ type: 'insertRows', sheet: DATA, at: 1, count: 2 });
    expect(['A1', 'A2', 'A3'].map((r) => loadedFormula(file, 'Calc', r))).toEqual(['Data!A1*2', 'Data!A4*2', 'Data!A5*2']);
    expect(sheetXml(xml, 2)).not.toContain('t="shared"');
  });

  it('grows a table the rows go into', () => {
    const { file, xml } = after({ type: 'insertRows', sheet: DATA, at: 1, count: 2 });
    expect(xml['xl/tables/table1.xml']).toMatch(/<table [^>]*ref="C1:E5"/);
    expect(xml['xl/tables/table1.xml']).toContain('<autoFilter ref="C1:E5">');
    expect(file.model.sheets[DATA].info.tables[0].rect).toEqual({ r1: 0, c1: 2, r2: 4, c2: 4 });
  });
});

describe('deleting rows', () => {
  it('removes threaded comments on deleted cells and resets a selection that went with them', () => {
    const { file, xml } = after({ type: 'deleteRows', sheet: DATA, at: 4, count: 2 });
    expect(xml['xl/threadedComments/threadedComment1.xml']).not.toContain('<threadedComment ');
    expect(file.model.sheets[DATA].info.notes).toEqual([]);
    expect(sheetXml(xml, 1)).toContain('<selection activeCell="A1" sqref="A1"/>');
    // The picture's top edge was on a deleted row: it now starts where they were.
    expect(anchors(xml['xl/drawings/drawing1.xml'], 'row')).toEqual([4, 6]);
  });

  it('pulls page breaks up with the rows above them', () => {
    const { xml } = after({ type: 'deleteRows', sheet: DATA, at: 1, count: 2 });
    expect(sheetXml(xml, 1)).toMatch(/<rowBreaks count="1" manualBreakCount="1"><brk id="1" /);
    const gone = after({ type: 'deleteRows', sheet: DATA, at: 2, count: 2 });
    expect(sheetXml(gone.xml, 1)).toMatch(/<rowBreaks count="0" manualBreakCount="0"><\/rowBreaks>/);
  });

  it('drops ranges that are wholly deleted, with their containers', () => {
    // With the table out of the way, lower down.
    const moved = patch(bytes, { 'xl/tables/table1.xml': (x) => x.replace(/C1:E3/g, 'C20:E22') });
    const { xml } = after({ type: 'deleteRows', sheet: DATA, at: 0, count: 6 }, moved);
    const sheet = sheetXml(xml, 1);
    expect(sheet).not.toContain('<protectedRanges');
    expect(sheet).not.toContain('<ignoredErrors');
    expect(sheet).not.toContain('<x14:dataValidation ');
    expect(sheet).not.toContain('<x14:conditionalFormatting ');
    expect(xml['xl/tables/table1.xml']).toMatch(/<table [^>]*ref="C14:E16"/);
    expect(xml['xl/workbook.xml']).toMatch(/name="Here"[^>]*>#REF!</);
  });
});

describe('columns', () => {
  it('move column widths, column page breaks and pictures', () => {
    const { file, xml } = after({ type: 'insertCols', sheet: DATA, at: 1, count: 1 });
    const sheet = sheetXml(xml, 1);
    expect(sheet).toContain('<col min="3" max="3" width="20"');
    expect(sheet).toContain('<col min="5" max="5" width="30"');
    expect(sheet).toMatch(/<colBreaks count="1" manualBreakCount="1"><brk id="3" /);
    expect(anchors(xml['xl/drawings/drawing1.xml'], 'col')).toEqual([6, 8]);
    // The table was right of the new column, so it moves without changing.
    expect(file.model.sheets[DATA].info.tables[0]).toMatchObject({ rect: { r1: 0, c1: 3, r2: 2, c2: 5 }, columns: ['X', 'Y', 'Z'] });
  });

  it('name columns inserted inside a table, in the table and in its header row', () => {
    const { file, xml } = after({ type: 'insertCols', sheet: DATA, at: 3, count: 2 });
    expect(xml['xl/tables/table1.xml']).toMatch(/<tableColumns count="5"><tableColumn id="1" name="X"[^>]*\/><tableColumn id="4" name="Column1"\/><tableColumn id="5" name="Column2"\/><tableColumn id="2" name="Y"/);
    expect(file.model.sheets[DATA].info.tables[0]).toMatchObject({ rect: { r1: 0, c1: 2, r2: 2, c2: 6 }, columns: ['X', 'Column1', 'Column2', 'Y', 'Z'] });
    expect(['C1', 'D1', 'E1', 'F1'].map((r) => loadedValue(file, 'Data', r))).toEqual(['X', 'Column1', 'Column2', 'Y']);
  });

  it('skip names already taken and leave tables without a header row unlabelled', () => {
    const named = patch(bytes, {
      'xl/tables/table1.xml': (x) => x.replace('name="Y"', 'name="Column1"').replace('headerRowCount="1"', 'headerRowCount="0"'),
    });
    const { xml } = after({ type: 'insertCols', sheet: DATA, at: 3, count: 1 }, named);
    expect(xml['xl/tables/table1.xml']).toContain('<tableColumn id="4" name="Column2"/>');
    // Without a header row, D1 is a record cell, so nothing is written into it.
    expect(sheetXml(xml, 1)).not.toContain('Column2');
  });

  it('stretch and shrink column ranges that span the change', () => {
    const wide = patch(bytes, { 'xl/worksheets/sheet1.xml': (x) => x.replace(/<cols>[\s\S]*?<\/cols>/, '<cols><col min="2" max="4" width="20" customWidth="1"/></cols>') });
    expect(sheetXml(after({ type: 'insertCols', sheet: DATA, at: 2, count: 2 }, wide).xml, 1)).toContain('<col min="2" max="6" width="20"');
    expect(sheetXml(after({ type: 'deleteCols', sheet: DATA, at: 2, count: 1 }, wide).xml, 1)).toContain('<col min="2" max="3" width="20"');
    // A width set past the last column stays on the sheet.
    const last = patch(bytes, { 'xl/worksheets/sheet1.xml': (x) => x.replace(/<cols>[\s\S]*?<\/cols>/, '<cols><col min="16380" max="16384" width="5" customWidth="1"/></cols>') });
    expect(sheetXml(after({ type: 'insertCols', sheet: DATA, at: 0, count: 10 }, last).xml, 1)).toContain('<col min="16384" max="16384" width="5"');
  });

  it('remove a deleted column from the table and its width from the sheet', () => {
    const { file, xml } = after({ type: 'deleteCols', sheet: DATA, at: 3, count: 1 });
    expect(xml['xl/tables/table1.xml']).toMatch(/<tableColumns count="2"><tableColumn id="1" name="X"[^>]*\/><tableColumn id="3" name="Z"/);
    expect(file.model.sheets[DATA].info.tables[0]).toMatchObject({ rect: { r1: 0, c1: 2, r2: 2, c2: 3 }, columns: ['X', 'Z'] });
    expect(sheetXml(xml, 1)).not.toContain('width="30"');
    // The column break after B is untouched by deleting D, and the picture moves left.
    expect(sheetXml(xml, 1)).toMatch(/<colBreaks count="1" manualBreakCount="1"><brk id="2" /);
    expect(anchors(xml['xl/drawings/drawing1.xml'], 'col')).toEqual([4, 6]);
  });
});

describe('moving cells', () => {
  it('into an empty sheet fills its cells and sets its used range', () => {
    const { file, xml } = after({ type: 'move', sheet: DATA, src: { r1: 3, c1: 0, r2: 5, c2: 0 }, toSheet: EMPTY, row: 1, col: 1 });
    expect(['B2', 'B3', 'B4'].map((r) => loadedValue(file, 'Empty', r))).toEqual([4, 5, 6]);
    expect(sheetXml(xml, 3)).toContain('<dimension ref="B2:B4"/>');
    // References to the cells follow them, qualified with the new sheet.
    expect(xml['xl/workbook.xml']).toContain('<definedName name="Here" localSheetId="0">Empty!$B$2</definedName>');
    expect(xml['xl/threadedComments/threadedComment1.xml']).not.toContain('<threadedComment ');
  });

  it('into a sheet written without a <sheetData> element adds one in its proper place', () => {
    const op: StructuralOp = { type: 'move', sheet: DATA, src: { r1: 3, c1: 0, r2: 5, c2: 0 }, toSheet: EMPTY, row: 1, col: 1 };
    const withCols = patch(bytes, { 'xl/worksheets/sheet3.xml': (x) => x.replace('<sheetData/>', '<cols><col min="1" max="1" width="9" customWidth="1"/></cols>') });
    const a = after(op, withCols);
    expect(sheetXml(a.xml, 3)).toMatch(/<\/cols><sheetData><row r="2"><c r="B2"><v>4<\/v><\/c><\/row>/);
    expect(loadedValue(a.file, 'Empty', 'B4')).toBe(6);
    const bare = patch(bytes, { 'xl/worksheets/sheet3.xml': (x) => x.replace(/<dimension[^>]*\/>|<sheetFormatPr[^>]*\/>|<sheetData\/>/g, '') });
    const b = after(op, bare);
    expect(sheetXml(b.xml, 3)).toMatch(/<worksheet [^>]*><sheetData>/);
    expect(loadedValue(b.file, 'Empty', 'B2')).toBe(4);
  });

  it('replaces cells pasted over on the destination sheet', () => {
    const { file } = after({ type: 'move', sheet: DATA, src: { r1: 0, c1: 0, r2: 1, c2: 0 }, toSheet: CALC, row: 1, col: 0 });
    expect(loadedValue(file, 'Calc', 'A2')).toBe(1);
    expect(loadedValue(file, 'Calc', 'A3')).toBe(2);
    expect(loadedFormula(file, 'Calc', 'A1')).toBe('A2*2');
  });

  it('keeps a moved single-cell array formula an array formula at its new address', () => {
    const ks = sample('kitchen-sink.xlsx');
    const s = WorkbookSession.open(ks, 'k.xlsx');
    const lookups = sheetIndex(s, 'Lookups');
    const { xml } = after({ type: 'move', sheet: lookups, src: { r1: 13, c1: 1, r2: 13, c2: 1 }, toSheet: lookups, row: 19, col: 7 }, ks, 'k.xlsx');
    expect(sheetXml(xml, 3)).toMatch(/<c r="H20"[^>]*><f t="array" ref="H20">SUM\(Data!D2:D201\*Data!E2:E201\)<\/f>/);
  });
});

describe('refusals', () => {
  const ks = () => WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');

  it('only change worksheets that are not protected', () => {
    const s = ks();
    const chart = sheetIndex(s, 'Monthly Chart');
    const data = sheetIndex(s, 'Data');
    const reference = sheetIndex(s, 'Reference');
    expect(thrown(() => s.restructure({ type: 'insertRows', sheet: chart, at: 0, count: 1 }))).toBe('Only worksheets can be changed like this.');
    expect(thrown(() => s.restructure({ type: 'insertRows', sheet: reference, at: 0, count: 1 }))).toBe('This sheet is protected in Excel, so it is read-only here.');
    const cell = { r1: 300, c1: 0, r2: 300, c2: 0 };
    expect(thrown(() => s.restructure({ type: 'move', sheet: data, src: cell, toSheet: chart, row: 0, col: 0 }))).toBe('Cells can only be moved between worksheets.');
    expect(thrown(() => s.restructure({ type: 'move', sheet: data, src: cell, toSheet: reference, row: 0, col: 0 }))).toBe('Reference is protected in Excel, so it is read-only here.');
  });

  it('keep data on the sheet and moved cells inside it', () => {
    const s = WorkbookSession.open(bytes, 'rich.xlsx');
    expect(thrown(() => s.restructure({ type: 'insertCols', sheet: DATA, at: 0, count: 16_380 }))).toMatch(/off the right of the sheet/);
    const src = { r1: 0, c1: 0, r2: 1, c2: 0 };
    expect(thrown(() => s.restructure({ type: 'move', sheet: DATA, src, toSheet: DATA, row: 1_048_575, col: 0 }))).toBe("The cells don't fit there.");
    expect(thrown(() => s.restructure({ type: 'move', sheet: DATA, src, toSheet: DATA, row: 0, col: -1 }))).toBe("The cells don't fit there.");
  });

  it('never split an array formula or a pivot table, or move onto one', () => {
    const s = ks();
    const lookups = sheetIndex(s, 'Lookups');
    expect(thrown(() => s.restructure({ type: 'deleteRows', sheet: lookups, at: 10, count: 1 }))).toBe(
      'That would change part of the array formula in B9:B12. Change the whole array, or none of it.',
    );
    expect(thrown(() => s.restructure({ type: 'insertCols', sheet: lookups, at: 2, count: 1 }))).toMatch(/array formula in B15:D15/);
    expect(thrown(() => s.restructure({ type: 'move', sheet: lookups, src: { r1: 30, c1: 0, r2: 30, c2: 0 }, toSheet: lookups, row: 9, col: 1 }))).toBe(
      'This cell is part of an array formula. Edit it in Excel.',
    );
    const p = WorkbookSession.open(sample('pivot.xlsx'), 'pivot.xlsx');
    expect(thrown(() => p.restructure({ type: 'insertRows', sheet: 1, at: 4, count: 1 }))).toBe('That would change part of the pivot table at A3:B8. Change the pivot table in Excel.');
    expect(thrown(() => p.restructure({ type: 'move', sheet: 1, src: { r1: 2, c1: 0, r2: 7, c2: 1 }, toSheet: 1, row: 20, col: 5 }))).toMatch(/pivot table "UnitsByRegion"/);
  });

  it('protect tables: their header, the whole table, and moves that would split them or take them to another sheet', () => {
    const s = ks();
    const data = sheetIndex(s, 'Data');
    const summary = sheetIndex(s, 'Summary');
    expect(thrown(() => s.restructure({ type: 'deleteRows', sheet: data, at: 0, count: 1 }))).toMatch(/delete the header row of table "SalesData"/);
    expect(thrown(() => s.restructure({ type: 'deleteRows', sheet: data, at: 0, count: 201 }))).toBe('That would delete all of table "SalesData". Delete the table in Excel.');
    expect(thrown(() => s.restructure({ type: 'deleteCols', sheet: data, at: 0, count: 6 }))).toMatch(/delete all of table "SalesData"/);
    expect(thrown(() => s.restructure({ type: 'move', sheet: data, src: { r1: 1, c1: 0, r2: 1, c2: 0 }, toSheet: data, row: 1, col: 10 }))).toBe(
      'That would move part of table "SalesData". Move the whole table, or cells outside it.',
    );
    expect(thrown(() => s.restructure({ type: 'move', sheet: data, src: { r1: 0, c1: 0, r2: 200, c2: 5 }, toSheet: summary, row: 40, col: 0 }))).toBe(
      'Tables can only be moved within their sheet.',
    );
  });

  it('keep merged cells on their own sheet', () => {
    const s = ks();
    const notes = sheetIndex(s, 'Notes');
    expect(thrown(() => s.restructure({ type: 'move', sheet: notes, src: { r1: 1, c1: 0, r2: 3, c2: 5 }, toSheet: sheetIndex(s, 'Settings'), row: 10, col: 0 }))).toBe(
      'Merged cells can only be moved within their sheet. Copy them instead.',
    );
  });
});

describe('the kitchen sink', () => {
  const ks = sample('kitchen-sink.xlsx');
  const s = WorkbookSession.open(ks, 'k.xlsx');
  const summary = sheetIndex(s, 'Summary');
  const notes = sheetIndex(s, 'Notes');

  it('moves charts, notes, links and merges with inserted rows', () => {
    const { file, xml } = after({ type: 'insertRows', sheet: summary, at: 0, count: 1 }, ks, 'k.xlsx');
    expect(xml['xl/charts/chart1.xml']).toContain('<c:f>Summary!$A$6:$A$9</c:f>');
    expect(xml['xl/charts/chart3.xml']).toContain('<c:f>Summary!$B$23:$B$34</c:f>');
    expect(xml['xl/comments1.xml']).toContain('<comment ref="A5"');
    expect(xml['xl/drawings/vmlDrawing1.vml']).toContain('<x:Row>4</x:Row>');
    expect(xml['xl/drawings/vmlDrawing1.vml']).toContain('<x:Anchor>1, 15, 3, 10, 2, 47, 7, 4</x:Anchor>');
    const sheet = sheetXml(xml, 1);
    expect(sheet).toContain('<mergeCell ref="A2:G2"/>');
    expect(sheet).toContain('<hyperlink ref="A20"');
    expect(sheet).toContain('<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>');
    // A selection that names no cell is left as it is.
    expect(sheet).toContain('<selection pane="bottomRight"/>');
    expect(file.model.sheets[summary].info.notes.map((n) => [n.row, n.col])).toEqual([[4, 0]]);
    expect(file.model.sheets[summary].info.objects.filter((o) => o.kind === 'chart').map((o) => o.from?.row)).toEqual(
      s.model.sheets[summary].info.objects.filter((o) => o.kind === 'chart').map((o) => o.from!.row + 1),
    );
  });

  it('removes the note on a deleted row, and its box', () => {
    const { file, xml } = after({ type: 'deleteRows', sheet: summary, at: 3, count: 1 }, ks, 'k.xlsx');
    expect(xml['xl/comments1.xml']).not.toContain('<comment ');
    expect(xml['xl/drawings/vmlDrawing1.vml']).not.toContain('<x:Row>');
    expect(file.model.sheets[summary].info.notes).toEqual([]);
    // Charts drawn from the rows below move up.
    expect(xml['xl/charts/chart1.xml']).toContain('<c:f>Summary!$A$4:$A$7</c:f>');
  });

  it('follows sparklines and drops a merge whose rows were deleted', () => {
    const ins = after({ type: 'insertRows', sheet: notes, at: 9, count: 1 }, ks, 'k.xlsx');
    expect(sheetXml(ins.xml, 4)).toContain('<xm:f>Notes!B11:F11</xm:f><xm:sqref>G11</xm:sqref>');
    const del = after({ type: 'deleteRows', sheet: notes, at: 1, count: 3 }, ks, 'k.xlsx');
    expect(sheetXml(del.xml, 4)).not.toContain('<mergeCells');
    expect(del.file.model.sheets[notes].info.merges).toEqual([]);
    expect(anchors(del.xml['xl/drawings/drawing2.xml'], 'row').slice(0, 2)).toEqual([1, 1]);
  });

  it('moves array formulas with their rows', () => {
    const lookups = sheetIndex(s, 'Lookups');
    const { xml } = after({ type: 'insertRows', sheet: lookups, at: 0, count: 1 }, ks, 'k.xlsx');
    const sheet = sheetXml(xml, 3);
    expect(sheet).toContain('<f t="array" ref="B10:B13">_xlfn._xlws.SORT(_xlfn.UNIQUE(Data!B2:B201))</f>');
    expect(sheet).toContain('<f t="array" ref="B16:D16">TRANSPOSE(G5:G7*2)</f>');
  });
});

describe('pivot tables', () => {
  const pv = sample('pivot.xlsx');

  it('read from a source that follows its rows', () => {
    const { xml } = after({ type: 'insertRows', sheet: 0, at: 0, count: 2 }, pv, 'pivot.xlsx');
    expect(xml['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('<worksheetSource ref="A3:E103" sheet="Data"/>');
    expect(xml['xl/pivotTables/pivotTable1.xml']).toContain('<location ref="A3:B8"');
  });

  it('move with rows inserted above them, leaving their source alone', () => {
    const { file, xml } = after({ type: 'insertRows', sheet: 1, at: 0, count: 1 }, pv, 'pivot.xlsx');
    expect(xml['xl/pivotTables/pivotTable1.xml']).toContain('<location ref="A4:B9"');
    expect(xml['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('<worksheetSource ref="A1:E101" sheet="Data"/>');
    expect(file.model.sheets[1].info.blocked.find((b) => b.kind === 'pivot')).toMatchObject({ r1: 3, r2: 8 });
  });
});
