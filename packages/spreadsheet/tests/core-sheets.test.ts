import ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { invertIndexMap, isSheetOp, nearestVisible, nextSheetName, sheetIndexMap, sheetNameProblem, sheetOpProblem, type SheetOp } from '../src/core/sheetRules';
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

const say = (m: Parameters<typeof english.text>[0]) => english.text(m);

/** Rewrites parts of a package; a string adds or replaces the part, null removes it. */
function patch(bytes: Uint8Array, edits: Record<string, string | null | ((xml: string) => string)>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [path, edit] of Object.entries(edits)) {
    if (edit === null) delete files[path];
    else files[path] = strToU8(typeof edit === 'string' ? edit : edit(strFromU8(files[path])));
  }
  return zipSync(files);
}

/**
 * Data!A1:A3 = 1, 2, 3, B1 = 0.2, and table Orders (Qty: 5, 7) at D1:D3.
 * My Sheet!A1:A2 = 10, 20. Hidden is hidden; Calc is the active tab.
 * Calc: A1:A3 = Data!A1*2 shared down, B1 = SUM('My Sheet'!A1:A2),
 * B2 = SUM(Orders[Qty]), B3 = Rate*2, B4 = SUM('Data:My Sheet'!A1) (3-D),
 * B5 = Data!B1+1, a list rule on C1 from Data!$A$1:$A$3, a formatting rule
 * Data!$A$1>1 on C1:C3, and a link from D1 to 'My Sheet'!A1.
 * Names: Rate = Data!$B$1; Local (Calc's own) = Calc!$A$1; My Sheet's print area.
 * Spare parts the editor doesn't read name sheets too: a form control, its
 * legacy drawing, a shape whose text comes from a cell, and a pivot chart.
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  const mine = wb.addWorksheet('My Sheet');
  const calc = wb.addWorksheet('Calc');
  wb.addWorksheet('Hidden', { state: 'hidden' }).getCell('A1').value = 'x';
  [1, 2, 3].forEach((v, i) => (data.getCell(`A${i + 1}`).value = v));
  data.getCell('B1').value = 0.2;
  data.addTable({ name: 'Orders', ref: 'D1', headerRow: true, columns: [{ name: 'Qty' }], rows: [[5], [7]] });
  mine.getCell('A1').value = 10;
  mine.getCell('A2').value = 20;
  calc.getCell('A1').value = { formula: 'Data!A1*2', shareType: 'shared', ref: 'A1:A3', result: 2 } as ExcelJS.CellValue;
  calc.getCell('A2').value = { sharedFormula: 'A1', result: 4 };
  calc.getCell('A3').value = { sharedFormula: 'A1', result: 6 };
  calc.getCell('B1').value = { formula: "SUM('My Sheet'!A1:A2)", result: 30 };
  calc.getCell('B2').value = { formula: 'SUM(Orders[Qty])', result: 12 };
  calc.getCell('B3').value = { formula: 'Rate*2', result: 0.4 };
  calc.getCell('B4').value = { formula: "SUM('Data:My Sheet'!A1)", result: 11 };
  calc.getCell('B5').value = { formula: 'Data!B1+1', result: 1.2 };
  wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 2, visibility: 'visible' }];
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  return patch(bytes, {
    'xl/workbook.xml': (xml) =>
      xml.replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="Rate">Data!$B$1</definedName><definedName name="Local" localSheetId="2">Calc!$A$1</definedName>' +
          '<definedName name="_xlnm.Print_Area" localSheetId="1">\'My Sheet\'!$A$1:$A$2</definedName></definedNames>',
      ),
    'xl/worksheets/sheet2.xml': (xml) => xml.replace('<sheetFormatPr', '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr'),
    'xl/worksheets/sheet3.xml': (xml) =>
      xml
        .replace('<sheetFormatPr', '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr')
        .replace(
          '</sheetData>',
          '</sheetData><conditionalFormatting sqref="C1:C3"><cfRule type="expression" dxfId="0" priority="1"><formula>Data!$A$1&gt;1</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="C1"><formula1>Data!$A$1:$A$3</formula1></dataValidation></dataValidations>' +
            '<hyperlinks><hyperlink ref="D1" location="\'My Sheet\'!A1" display="Mine"/></hyperlinks>',
        ),
    'docProps/app.xml': (xml) =>
      xml
        .replace('<vt:vector size="2" baseType="variant">', '<vt:vector size="4" baseType="variant">')
        .replace('<vt:i4>4</vt:i4></vt:variant>', '<vt:i4>4</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant>')
        .replace('<vt:vector size="4" baseType="lpstr">', '<vt:vector size="6" baseType="lpstr">')
        .replace('<vt:lpstr>Hidden</vt:lpstr>', "<vt:lpstr>Hidden</vt:lpstr><vt:lpstr>'My Sheet'!Print_Area</vt:lpstr><vt:lpstr>Rate</vt:lpstr>"),
    'xl/ctrlProps/ctrlProp1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Drop" fmlaLink="Data!$A$1" fmlaRange="Calc!$A$1:$A$3"/>',
    'xl/drawings/vmlDrawing9.vml': '<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><x:ClientData ObjectType="Drop"><x:FmlaLink>Data!$A$1</x:FmlaLink></x:ClientData></xml>',
    'xl/drawings/drawing9.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:sp macro="" textlink="Data!$A$1"><xdr:txBody/></xdr:sp></xdr:wsDr>',
    'xl/charts/chart9.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:pivotSource><c:name>[Book.xlsx]Data!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart><c:f>Data!$A$1:$A$3</c:f></c:chart></c:chartSpace>',
  });
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

const DATA = 0;
const MINE = 1;
const CALC = 2;
const HIDDEN = 3;

/** The workbook after the change: the new session, the saved file opened again, and its parts. */
function after(op: SheetOp, from = bytes, name = 'sheets.xlsx') {
  const s = WorkbookSession.open(from, name, 'en-US', { recalc: 'on' });
  s.ensureEngine();
  const next = s.restructure(op);
  const saved = next.save();
  return { before: s, next, file: WorkbookSession.open(saved, name), xml: parts(saved) };
}

const names = (s: WorkbookSession) => s.model.sheets.map((m) => m.info.name);
const definedNames = (xml: Record<string, string>) => [...xml['xl/workbook.xml'].matchAll(/<definedName ([^>]*)>([^<]*)</g)].map((m) => `${m[1]} = ${m[2]}`);
const titles = (xml: Record<string, string>) => [...xml['docProps/app.xml'].matchAll(/<vt:(lpstr|i4)>([^<]*)</g)].map((m) => m[2]).join(' | ');
const view = (xml: Record<string, string>) => /<workbookView [^>]*\/>/.exec(xml['xl/workbook.xml'])?.[0] ?? '';
/** What a cell shows once calculated: its recalculated value, or Excel's saved one. */
const shown = (s: WorkbookSession, sheet: string, ref: string) => {
  s.ensureEngine();
  const at = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const update = s.stateUpdates().find((u) => u.sheet === sheetIndex(s, sheet) && u.row === Number(at[2]) - 1 && u.col === at[1].charCodeAt(0) - 65);
  return update ? (update.s ?? update.n) : loadedValue(s, sheet, ref);
};

describe('sheet names', () => {
  it("follows Excel's rules", () => {
    const problem = (name: string, others: string[] = []) => {
      const m = sheetNameProblem(name, others);
      return m && say(m);
    };
    expect(problem('Q1 2026')).toBeUndefined();
    expect(problem('x'.repeat(31))).toBeUndefined();
    expect(problem('')).toBe('Type a name for the sheet.');
    expect(problem('   ')).toBe('Type a name for the sheet.');
    expect(problem('x'.repeat(32))).toBe('A sheet name can be at most 31 characters long.');
    for (const bad of ['a:b', 'a\\b', 'a/b', 'a?b', 'a*b', 'a[b', 'a]b', 'a\nb']) expect(problem(bad)).toMatch(/can't contain/);
    expect(problem("'Quoted")).toMatch(/apostrophe/);
    expect(problem("Quoted'")).toMatch(/apostrophe/);
    expect(problem("It's")).toBeUndefined();
    expect(problem('history')).toBe('Excel keeps the name “History” for itself.');
    expect(problem('data', ['Data'])).toMatch(/already a sheet called “data”/);
  });

  it('numbers new sheets one past the highest, in the language given', () => {
    expect(nextSheetName('Sheet', ['Sheet1'])).toBe('Sheet2');
    expect(nextSheetName('Sheet', ['Data', 'Summary'])).toBe('Sheet1');
    expect(nextSheetName('Sheet', ['Sheet1', 'Sheet3'])).toBe('Sheet4');
    expect(nextSheetName('Sheet', ['sheet2', 'SHEET9', 'Sheet 10', 'Sheetx'])).toBe('Sheet10');
    expect(nextSheetName('Tabelle', ['Tabelle1', 'Sheet5'])).toBe('Tabelle2');
  });

  it('maps where each sheet goes, and back', () => {
    expect(sheetIndexMap({ type: 'addSheet', at: 1, name: 'N' }, 3)).toEqual([0, 2, 3]);
    expect(sheetIndexMap({ type: 'deleteSheet', sheet: 1 }, 3)).toEqual([0, null, 1]);
    expect(sheetIndexMap({ type: 'moveSheet', sheet: 0, to: 2 }, 3)).toEqual([2, 0, 1]);
    expect(sheetIndexMap({ type: 'moveSheet', sheet: 2, to: 0 }, 3)).toEqual([1, 2, 0]);
    expect(sheetIndexMap({ type: 'renameSheet', sheet: 1, name: 'N' }, 3)).toEqual([0, 1, 2]);
    expect(invertIndexMap([0, 2, 3], 4)).toEqual([0, null, 1, 2]);
    expect(invertIndexMap([0, null, 1], 2)).toEqual([0, 2]);
    expect(isSheetOp({ type: 'hideSheet', sheet: 0, hidden: true })).toBe(true);
    expect(isSheetOp({ type: 'insertRows', sheet: 0, at: 0, count: 1 })).toBe(false);
  });

  it('says why Excel would refuse a change', () => {
    const sheets = [
      { name: 'A', state: 'visible' as const },
      { name: 'B', state: 'hidden' as const },
      { name: 'C', state: 'veryHidden' as const },
    ];
    const problem = (op: SheetOp, locked = false) => {
      const m = sheetOpProblem(op, sheets, locked);
      return m && say(m);
    };
    expect(problem({ type: 'addSheet', at: 1, name: 'D' })).toBeUndefined();
    expect(problem({ type: 'addSheet', at: 1, name: 'D' }, true)).toMatch(/structure is protected/);
    expect(problem({ type: 'addSheet', at: 4, name: 'D' })).toMatch(/isn't in the workbook/);
    expect(problem({ type: 'addSheet', at: 0, name: 'b' })).toMatch(/already a sheet/);
    expect(problem({ type: 'renameSheet', sheet: 0, name: 'a' })).toBeUndefined();
    expect(problem({ type: 'renameSheet', sheet: 0, name: 'B' })).toMatch(/already a sheet/);
    expect(problem({ type: 'renameSheet', sheet: 5, name: 'X' })).toMatch(/isn't in the workbook/);
    expect(problem({ type: 'deleteSheet', sheet: 0 })).toMatch(/at least one visible sheet, so this one can't be deleted/);
    expect(problem({ type: 'deleteSheet', sheet: 1 })).toBeUndefined();
    expect(problem({ type: 'moveSheet', sheet: 0, to: 2 })).toBeUndefined();
    expect(problem({ type: 'moveSheet', sheet: 0, to: 3 })).toMatch(/isn't in the workbook/);
    expect(problem({ type: 'hideSheet', sheet: 0, hidden: true })).toMatch(/can't be hidden/);
    expect(problem({ type: 'hideSheet', sheet: 1, hidden: true })).toBeUndefined();
    expect(problem({ type: 'hideSheet', sheet: 1, hidden: false })).toBeUndefined();
    expect(problem({ type: 'hideSheet', sheet: 2, hidden: false })).toMatch(/Visual Basic editor/);
  });

  it('finds the visible sheet nearest a position, after it first', () => {
    const states = (s: string) => [...s].map((c) => ({ state: c === 'v' ? ('visible' as const) : ('hidden' as const) }));
    expect(nearestVisible(states('vhv'), 1)).toBe(2);
    expect(nearestVisible(states('vhh'), 1)).toBe(0);
    expect(nearestVisible(states('vv'), 2)).toBe(1);
    expect(nearestVisible(states('hh'), 0)).toBe(-1);
  });
});

describe('renaming a sheet', () => {
  it('follows the new name in formulas, names, rules, 3-D references and the list of parts, and keeps calculating', () => {
    const { next, file, xml } = after({ type: 'renameSheet', sheet: DATA, name: 'Sales 2026' });
    expect(names(file)).toEqual(['Sales 2026', 'My Sheet', 'Calc', 'Hidden']);
    expect(['A1', 'A2', 'A3', 'B4', 'B5'].map((r) => loadedFormula(file, 'Calc', r))).toEqual([
      "'Sales 2026'!A1*2",
      "'Sales 2026'!A2*2",
      "'Sales 2026'!A3*2",
      "SUM('Sales 2026:My Sheet'!A1)",
      "'Sales 2026'!B1+1",
    ]);
    // The shared formula stays shared.
    expect(xml['xl/worksheets/sheet3.xml']).toContain('<f t="shared" ref="A1:A3" si="0">\'Sales 2026\'!A1*2</f>');
    expect(definedNames(xml)).toContain("name=\"Rate\" = 'Sales 2026'!$B$1");
    expect(xml['xl/worksheets/sheet3.xml']).toContain("<formula>'Sales 2026'!$A$1&gt;1</formula>");
    expect(xml['xl/worksheets/sheet3.xml']).toContain("<formula1>'Sales 2026'!$A$1:$A$3</formula1>");
    expect(titles(xml)).toBe("Worksheets | 4 | Named Ranges | 2 | Sales 2026 | My Sheet | Calc | Hidden | 'My Sheet'!Print_Area | Rate");
    expect(shown(next, 'Calc', 'A2')).toBe(4);
    expect(shown(next, 'Calc', 'B5')).toBe(1.2);
    expect(next.changes().map((c) => [c.sheet, say(c.note!)])).toEqual([['Sales 2026', 'Sheet renamed from “Data”']]);
  });

  it('follows it in links, print areas and parts the editor only keeps', () => {
    const { xml } = after({ type: 'renameSheet', sheet: MINE, name: 'Notes' });
    expect(xml['xl/worksheets/sheet3.xml']).toMatch(/<hyperlink ref="D1" location="Notes!A1" display="Mine"\/>/);
    expect(definedNames(xml)).toContain('name="_xlnm.Print_Area" localSheetId="1" = Notes!$A$1:$A$2');
    expect(titles(xml)).toContain('Notes!Print_Area');
    const { xml: data } = after({ type: 'renameSheet', sheet: DATA, name: 'Raw' });
    expect(data['xl/ctrlProps/ctrlProp1.xml']).toContain('fmlaLink="Raw!$A$1" fmlaRange="Calc!$A$1:$A$3"');
    expect(data['xl/drawings/vmlDrawing9.vml']).toContain('<x:FmlaLink>Raw!$A$1</x:FmlaLink>');
    expect(data['xl/drawings/drawing9.xml']).toContain('textlink="Raw!$A$1"');
    expect(data['xl/charts/chart9.xml']).toContain('<c:name>[Book.xlsx]Raw!PivotTable1</c:name>');
    expect(data['xl/charts/chart9.xml']).toContain('<c:f>Raw!$A$1:$A$3</c:f>');
  });

  it('writes a name that needs quoting with its quotes doubled, and a change of capitals alone', () => {
    const { file } = after({ type: 'renameSheet', sheet: DATA, name: "Bob's" });
    expect(loadedFormula(file, 'Calc', 'B5')).toBe("'Bob''s'!B1+1");
    const again = after({ type: 'renameSheet', sheet: 0, name: 'Data' }, file.save());
    expect(loadedFormula(again.file, 'Calc', 'B5')).toBe('Data!B1+1');
    const caps = after({ type: 'renameSheet', sheet: DATA, name: 'DATA' });
    expect(loadedFormula(caps.file, 'Calc', 'B5')).toBe('DATA!B1+1');
  });

  it("follows charts' series and pivot sources", () => {
    const k = after({ type: 'renameSheet', sheet: 0, name: 'Overview' }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    for (const chart of ['chart1', 'chart2', 'chart3']) {
      expect(k.xml[`xl/charts/${chart}.xml`]).toContain('<c:f>Overview!$A$');
      expect(k.xml[`xl/charts/${chart}.xml`]).not.toContain('Summary!');
    }
    expect(k.file.model.sheets[sheetIndex(k.file, 'Monthly Chart')].info.objects[0].chart?.groups[0].series[0].values.ref).toEqual({ sheet: 0, rect: { r1: 21, c1: 1, r2: 32, c2: 1 } });
    const p = after({ type: 'renameSheet', sheet: 0, name: 'Raw data' }, sample('pivot.xlsx'), 'pivot.xlsx');
    expect(p.xml['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('<worksheetSource ref="A1:E101" sheet="Raw data"/>');
  });

  it('changes nothing for the same name, and refuses a name Excel would', () => {
    const s = WorkbookSession.open(bytes, 'sheets.xlsx');
    expect(parts(s.restructure({ type: 'renameSheet', sheet: DATA, name: 'Data' }).save())['xl/workbook.xml']).toContain('name="Data"');
    expect(thrown(() => s.restructure({ type: 'renameSheet', sheet: DATA, name: 'calc' }))).toMatch(/already a sheet called “calc”/);
    expect(thrown(() => s.restructure({ type: 'renameSheet', sheet: DATA, name: 'a/b' }))).toMatch(/can't contain/);
  });
});

describe('deleting a sheet', () => {
  it('turns references to it into #REF!, keeps 3-D references to what is left, and says how many formulas broke', () => {
    const { next, file, xml } = after({ type: 'deleteSheet', sheet: DATA });
    expect(names(file)).toEqual(['My Sheet', 'Calc', 'Hidden']);
    expect(['A1', 'A2', 'B1', 'B2', 'B3', 'B4', 'B5'].map((r) => loadedFormula(file, 'Calc', r))).toEqual([
      '#REF!*2',
      '#REF!*2',
      "SUM('My Sheet'!A1:A2)",
      'SUM(#REF!)',
      'Rate*2',
      "SUM('My Sheet'!A1)",
      '#REF!+1',
    ]);
    expect(say(next.structureWarnings[0])).toBe('5 formulas now show #REF! because they used cells that were deleted or pasted over.');
    expect(shown(next, 'Calc', 'A1')).toBe('#REF!');
    expect(definedNames(xml)).toEqual(['name="Rate" = #REF!', 'name="Local" localSheetId="1" = Calc!$A$1', 'name="_xlnm.Print_Area" localSheetId="0" = \'My Sheet\'!$A$1:$A$2']);
    expect(xml['xl/worksheets/sheet3.xml']).toContain('<formula1>#REF!</formula1>');
    expect(xml['xl/worksheets/sheet3.xml']).toContain('<formula>#REF!&gt;1</formula>');
    expect(xml['xl/ctrlProps/ctrlProp1.xml']).toContain('fmlaLink="#REF!"');
    // Charts keep a reference that shows #REF!; a pivot chart's source is left for Excel.
    expect(xml['xl/charts/chart9.xml']).toContain('<c:f>#REF!</c:f>');
    expect(view(xml)).toContain('activeTab="1"');
    expect(titles(xml)).toBe("Worksheets | 3 | Named Ranges | 2 | My Sheet | Calc | Hidden | 'My Sheet'!Print_Area | Rate");
    expect(next.changes().map((c) => [c.sheet, say(c.note!)])).toEqual([['Data', 'Sheet deleted']]);
  });

  it('removes its part and everything only it used, tables included', () => {
    const { xml, file } = after({ type: 'deleteSheet', sheet: DATA });
    expect(xml['xl/worksheets/sheet1.xml']).toBeUndefined();
    expect(xml['xl/tables/table1.xml']).toBeUndefined();
    expect(xml['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined();
    expect(xml['[Content_Types].xml']).not.toContain('/xl/worksheets/sheet1.xml');
    expect(xml['[Content_Types].xml']).not.toContain('/xl/tables/table1.xml');
    expect(xml['xl/_rels/workbook.xml.rels']).not.toContain('worksheets/sheet1.xml');
    // The table's name is free again.
    expect(file.model.sheets.flatMap((m) => m.info.tables)).toEqual([]);
  });

  it("drops names local to it and moves the active tab to the next visible sheet, which Excel then shows", () => {
    const { xml, file } = after({ type: 'deleteSheet', sheet: CALC });
    expect(definedNames(xml)).toEqual(['name="Rate" = Data!$B$1', 'name="_xlnm.Print_Area" localSheetId="1" = \'My Sheet\'!$A$1:$A$2']);
    // Hidden comes next but is hidden, so the sheet before it.
    expect(view(xml)).toContain('activeTab="1"');
    expect(xml['xl/worksheets/sheet2.xml']).toMatch(/<sheetView [^>]*tabSelected="1"/);
    expect(names(file)).toEqual(['Data', 'My Sheet', 'Hidden']);
    // My Sheet's print area goes with it, and the list of parts loses both.
    const mine = after({ type: 'deleteSheet', sheet: MINE });
    expect(definedNames(mine.xml)).toEqual(['name="Rate" = Data!$B$1', 'name="Local" localSheetId="1" = Calc!$A$1']);
    expect(titles(mine.xml)).toBe('Worksheets | 3 | Named Ranges | 1 | Data | Calc | Hidden | Rate');
    // A link to it is left as it was, as Excel leaves it.
    expect(mine.xml['xl/worksheets/sheet3.xml']).toContain('location="\'My Sheet\'!A1"');
  });

  it('removes a definedNames element left empty', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('One');
    wb.addWorksheet('Two');
    const two = patch(new Uint8Array(await wb.xlsx.writeBuffer()), {
      'xl/workbook.xml': (xml) => xml.replace('</sheets>', '</sheets><definedNames><definedName name="X" localSheetId="1">Two!$A$1</definedName></definedNames>'),
    });
    expect(after({ type: 'deleteSheet', sheet: 1 }, two).xml['xl/workbook.xml']).not.toContain('definedNames');
  });

  it('shortens a 3-D reference that loses an end, and breaks one left with nothing', async () => {
    const wb = new ExcelJS.Workbook();
    for (const n of ['Jan', 'Feb', 'Mar']) wb.addWorksheet(n).getCell('A1').value = 1;
    const sum = wb.addWorksheet('Sum');
    sum.getCell('A1').value = { formula: 'SUM(Jan:Mar!A1)', result: 3 };
    sum.getCell('A2').value = { formula: 'SUM(Mar:Jan!A1)', result: 3 };
    sum.getCell('A3').value = { formula: 'SUM(Jan:Feb!A1)', result: 2 };
    sum.getCell('A4').value = { formula: 'SUM(Jan:Jan!A1)', result: 1 };
    sum.getCell('A5').value = { formula: 'SUM(Jan:Nowhere!A1)', result: 1 };
    sum.getCell('A6').value = { formula: '[1]Jan!A1', result: 1 };
    const three = new Uint8Array(await wb.xlsx.writeBuffer());
    const jan = after({ type: 'deleteSheet', sheet: 0 }, three);
    expect(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map((r) => loadedFormula(jan.file, 'Sum', r))).toEqual([
      'SUM(Feb:Mar!A1)',
      'SUM(Mar:Feb!A1)',
      'SUM(Feb!A1)',
      'SUM(#REF!)',
      'SUM(#REF!)',
      '[1]Jan!A1',
    ]);
    const mar = after({ type: 'deleteSheet', sheet: 2 }, three);
    expect(loadedFormula(mar.file, 'Sum', 'A1')).toBe('SUM(Jan:Feb!A1)');
    expect(loadedFormula(mar.file, 'Sum', 'A2')).toBe('SUM(Feb:Jan!A1)');
    const feb = after({ type: 'deleteSheet', sheet: 1 }, three);
    expect(loadedFormula(feb.file, 'Sum', 'A1')).toBe('SUM(Jan:Mar!A1)');
  });

  it('takes its drawings, charts, notes and pictures, keeping what other sheets still use', () => {
    const k = after({ type: 'deleteSheet', sheet: 0 }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    for (const gone of ['xl/drawings/drawing1.xml', 'xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/comments1.xml']) {
      expect(k.xml[gone], gone).toBeUndefined();
    }
    expect(k.xml['xl/media/image1.png']).toBeDefined();
    expect(k.xml['xl/charts/chart3.xml']).toContain('<c:f>#REF!</c:f>');
    expect(k.xml['docProps/app.xml']).not.toContain('<vt:lpstr>Summary</vt:lpstr>');
    const chartSheet = after({ type: 'deleteSheet', sheet: 6 }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    expect(chartSheet.xml['xl/charts/chart3.xml']).toBeUndefined();
    expect(chartSheet.xml['xl/chartsheets/sheet1.xml']).toBeUndefined();
    expect(titles(chartSheet.xml)).toBe('Worksheets | 6 | Named Ranges | 1 | Summary | Data | Lookups | Notes | Reference | Settings | TaxRate');
    const settings = after({ type: 'deleteSheet', sheet: 5 }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    expect(definedNames(settings.xml)).toEqual(['name="TaxRate" = #REF!']);
  });

  it("leaves a pivot table whose source went with nothing to refresh from, and keeps a cache when its pivot table goes", () => {
    const data = after({ type: 'deleteSheet', sheet: 0 }, sample('pivot.xlsx'), 'pivot.xlsx');
    const cache = data.xml['xl/pivotCache/pivotCacheDefinition1.xml'];
    expect(cache).toContain('<worksheetSource ref="#REF!"/>');
    expect(cache).not.toContain('refreshOnLoad');
    const pivot = after({ type: 'deleteSheet', sheet: 1 }, sample('pivot.xlsx'), 'pivot.xlsx');
    expect(pivot.xml['xl/pivotTables/pivotTable1.xml']).toBeUndefined();
    expect(pivot.xml['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('sheet="Data"');
  });

  it("breaks references to its tables, and a pivot cache's whose source is one", () => {
    // Orders is a table on Lines; a pivot cache reads it by name.
    const templates = patch(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))), {
      'xl/pivotCache/pivotCacheDefinition9.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1"><cacheSource type="worksheet"><worksheetSource name="Orders"/></cacheSource></pivotCacheDefinition>',
    });
    const { xml, file } = after({ type: 'deleteSheet', sheet: 1 }, templates, 'templates.xlsx');
    expect(loadedFormula(file, 'Request', 'B6')).toBe('SUM(#REF!)');
    expect(xml['xl/pivotCache/pivotCacheDefinition9.xml']).toContain('<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="#REF!"/>');
  });

  it('refuses to delete the last visible sheet, or one whose pivot tables or tables slicers use', () => {
    const one = WorkbookSession.open(sample('pivot.xlsx'), 'pivot.xlsx');
    const hidden = one.restructure({ type: 'hideSheet', sheet: 0, hidden: true });
    expect(thrown(() => hidden.restructure({ type: 'deleteSheet', sheet: 1 }))).toBe("A workbook must keep at least one visible sheet, so this one can't be deleted.");
    const sliced = patch(sample('pivot.xlsx'), {
      'xl/_rels/workbook.xml.rels': (xml) => xml.replace('</Relationships>', '<Relationship Id="rIdS" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/></Relationships>'),
    });
    expect(thrown(() => WorkbookSession.open(sliced, 'pivot.xlsx').restructure({ type: 'deleteSheet', sheet: 1 }))).toMatch(/slicers or timelines use/);
  });
});

describe('moving a sheet', () => {
  it('reorders the sheets, renumbering local names, the active tab and the list of parts', () => {
    const { xml, file } = after({ type: 'moveSheet', sheet: DATA, to: 2 });
    expect(names(file)).toEqual(['My Sheet', 'Calc', 'Data', 'Hidden']);
    expect(definedNames(xml)).toEqual(['name="Rate" = Data!$B$1', 'name="Local" localSheetId="1" = Calc!$A$1', 'name="_xlnm.Print_Area" localSheetId="0" = \'My Sheet\'!$A$1:$A$2']);
    expect(view(xml)).toContain('activeTab="1"');
    expect(titles(xml)).toBe("Worksheets | 4 | Named Ranges | 2 | My Sheet | Calc | Data | Hidden | 'My Sheet'!Print_Area | Rate");
    expect(loadedFormula(file, 'Calc', 'A2')).toBe('Data!A2*2');
    const k = after({ type: 'moveSheet', sheet: 6, to: 0 }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    expect(names(k.file)[0]).toBe('Monthly Chart');
    expect(titles(k.xml)).toMatch(/^Worksheets \| 6 \| Charts \| 1 \| Named Ranges \| 1 \| Summary/);
  });
});

describe('hiding and unhiding a sheet', () => {
  it("hides a sheet, moving the active tab off it, and unhides one", () => {
    const { xml, file, next } = after({ type: 'hideSheet', sheet: CALC, hidden: true });
    expect(xml['xl/workbook.xml']).toMatch(/<sheet [^>]*name="Calc"[^>]*state="hidden"/);
    expect(view(xml)).toContain('activeTab="1"');
    expect(xml['xl/worksheets/sheet3.xml']).not.toContain('tabSelected');
    expect(xml['xl/worksheets/sheet2.xml']).toMatch(/<sheetView [^>]*tabSelected="1"/);
    expect(file.model.sheets[CALC].info.state).toBe('hidden');
    expect(say(next.changes()[0].note!)).toBe('Sheet hidden');
    const shown = after({ type: 'hideSheet', sheet: HIDDEN, hidden: false });
    expect(shown.file.model.sheets[HIDDEN].info.state).toBe('visible');
    expect(shown.xml['xl/workbook.xml']).not.toMatch(/name="Hidden"[^>]*state=/);
    expect(say(shown.next.changes()[0].note!)).toBe('Sheet unhidden');
    // Hiding a sheet already hidden changes nothing.
    expect(after({ type: 'hideSheet', sheet: HIDDEN, hidden: true }).xml['xl/workbook.xml']).toContain('name="Hidden" state="hidden"');
  });

  it("refuses to hide the last visible sheet or show a very hidden one", async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Only');
    wb.addWorksheet('Secret', { state: 'veryHidden' });
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 'secret.xlsx');
    expect(thrown(() => s.restructure({ type: 'hideSheet', sheet: 0, hidden: true }))).toMatch(/can't be hidden/);
    expect(thrown(() => s.restructure({ type: 'hideSheet', sheet: 1, hidden: false }))).toMatch(/only Excel's Visual Basic editor/);
  });
});

describe('adding a sheet', () => {
  it('adds an empty worksheet with its part, relationship, content type and a new id, renumbering what counts sheets', () => {
    const { xml, file, next } = after({ type: 'addSheet', at: 1, name: 'Sheet1' });
    expect(names(file)).toEqual(['Data', 'Sheet1', 'My Sheet', 'Calc', 'Hidden']);
    expect(xml['xl/workbook.xml']).toContain('<sheet name="Sheet1" sheetId="5" r:id="rId');
    expect(xml['xl/worksheets/sheet5.xml']).toContain('<sheetData/>');
    expect(xml['xl/worksheets/sheet5.xml']).toContain('<sheetFormatPr defaultRowHeight="15"/>');
    expect(xml['[Content_Types].xml']).toContain('<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    expect(xml['xl/_rels/workbook.xml.rels']).toContain('Target="worksheets/sheet5.xml"');
    expect(definedNames(xml)).toEqual(['name="Rate" = Data!$B$1', 'name="Local" localSheetId="3" = Calc!$A$1', 'name="_xlnm.Print_Area" localSheetId="2" = \'My Sheet\'!$A$1:$A$2']);
    expect(view(xml)).toContain('activeTab="3"');
    expect(titles(xml)).toBe("Worksheets | 5 | Named Ranges | 2 | Data | Sheet1 | My Sheet | Calc | Hidden | 'My Sheet'!Print_Area | Rate");
    expect(file.model.sheets[1].info.kind).toBe('worksheet');
    expect(say(next.changes()[0].note!)).toBe('Sheet added');
    // It can be typed into and saved like any other.
    next.apply({ type: 'input', sheet: 1, row: 0, col: 0, text: '=Data!A3*10' });
    const typed = WorkbookSession.open(next.save(), 'sheets.xlsx');
    expect(loadedFormula(typed, 'Sheet1', 'A1')).toBe('Data!A3*10');
    expect(shown(typed, 'Sheet1', 'A1')).toBe(30);
  });

  it('adds one at the end, and to a workbook with charts, macro-free or not', () => {
    const k = after({ type: 'addSheet', at: 7, name: 'Sheet1' }, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    expect(names(k.file)[7]).toBe('Sheet1');
    expect(titles(k.xml)).toMatch(/^Worksheets \| 7 \| Charts \| 1 \| Named Ranges \| 1 \| Summary \| Data \| Lookups \| Notes \| Reference \| Settings \| Sheet1 \| Monthly Chart/);
    expect(k.xml['xl/worksheets/sheet7.xml']).toContain('<sheetData/>');
  });

  it('leaves a list of parts it cannot read as it was', () => {
    const odd = patch(bytes, { 'docProps/app.xml': (xml) => xml.replace('<vt:i4>4</vt:i4>', '<vt:i8>4</vt:i8>') });
    expect(after({ type: 'addSheet', at: 0, name: 'New' }, odd).xml['docProps/app.xml']).not.toContain('New');
    const none = patch(bytes, { 'docProps/app.xml': null });
    expect(names(after({ type: 'addSheet', at: 0, name: 'New' }, none).file)[0]).toBe('New');
  });
});

describe('a workbook whose structure is protected', () => {
  it('refuses every change to its sheets', () => {
    const locked = patch(bytes, { 'xl/workbook.xml': (xml) => xml.replace('<bookViews>', '<workbookProtection lockStructure="1"/><bookViews>') });
    const s = WorkbookSession.open(locked, 'locked.xlsx');
    expect(s.loadResult().structureLocked).toBe(true);
    const ops: SheetOp[] = [
      { type: 'addSheet', at: 0, name: 'New' },
      { type: 'renameSheet', sheet: 0, name: 'New' },
      { type: 'deleteSheet', sheet: 0 },
      { type: 'moveSheet', sheet: 0, to: 1 },
      { type: 'hideSheet', sheet: 0, hidden: true },
    ];
    for (const op of ops) expect(thrown(() => s.restructure(op))).toMatch(/structure is protected/);
  });
});

describe('changes before and after', () => {
  it("lists earlier cell changes, and rows inserted, under the sheet's new name, and drops those on a sheet deleted", () => {
    const s = WorkbookSession.open(bytes, 'sheets.xlsx');
    s.apply({ type: 'input', sheet: DATA, row: 0, col: 0, text: '5' });
    s.apply({ type: 'input', sheet: MINE, row: 0, col: 0, text: '11' });
    const inserted = s.restructure({ type: 'insertRows', sheet: DATA, at: 5, count: 1 });
    const renamed = inserted.restructure({ type: 'renameSheet', sheet: DATA, name: 'Raw' });
    expect(renamed.changes().map((c) => `${c.sheet}!${c.ref}`)).toEqual(['Raw!6:6', 'Raw!', 'Raw!A1', 'My Sheet!A1']);
    const deleted = renamed.restructure({ type: 'deleteSheet', sheet: DATA });
    expect(deleted.changes().map((c) => `${c.sheet}!${c.ref}`)).toEqual(['Raw!6:6', 'Raw!', 'Raw!', 'My Sheet!A1']);
    const moved = deleted.restructure({ type: 'moveSheet', sheet: 0, to: 1 });
    expect(moved.changes().find((c) => c.ref === 'A1')?.sheet).toBe('My Sheet');
    expect(moved.editedCount).toBe(5);
  });
});
