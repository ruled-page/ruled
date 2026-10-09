import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { CalcEngine } from '../src/core/calc';
import { renameInFormula, writeDefinedNames } from '../src/core/definedNames';
import {
  absoluteRef,
  findName,
  isListedName,
  nameOfSelection,
  nameProblem,
  nameTarget,
  refersToFormula,
  refersToProblem,
  selectionFormula,
} from '../src/core/names';
import { WorkbookSession } from '../src/core/session';
import type { DefinedName, EditResult, NameChange } from '../src/core/types';
import { english } from '../src/i18n';
import { loadedValue, parts, sample, sheetIndex, valueOf } from './helpers';

/**
 * Data!A1:A3 = 1, 2, 3 and B1 = 0.2; C1 = Here (the workbook's Here, A1).
 * Calc!A1 = SUM(Values), A2 = Rate*100, A3 = Here (Calc's own Here, C1 = 5),
 * A4 = Twice, A5 = Calc!Here*2. Names: a print area, Values (with a comment
 * and an attribute the editor doesn't know), Rate, Here twice (the
 * workbook's and Calc's), Secret (hidden) and Twice (=Rate*2).
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = 1;
  data.getCell('A2').value = 2;
  data.getCell('A3').value = 3;
  data.getCell('B1').value = 0.2;
  data.getCell('C1').value = { formula: 'Here', result: 1 };
  const calc = wb.addWorksheet('Calc');
  calc.getCell('A1').value = { formula: 'SUM(Values)', result: 6 };
  calc.getCell('A2').value = { formula: 'Rate*100', result: 20 };
  calc.getCell('A3').value = { formula: 'Here', result: 5 };
  calc.getCell('A4').value = { formula: 'Twice', result: 0.4 };
  calc.getCell('A5').value = { formula: 'Calc!Here*2', result: 10 };
  calc.getCell('C1').value = 5;
  const files = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const xml = strFromU8(files['xl/workbook.xml']).replace(/<definedNames>.*<\/definedNames>|<definedNames\/>/, '');
  files['xl/workbook.xml'] = strToU8(
    xml.replace(
      '</sheets>',
      '</sheets><definedNames>' +
        '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$3</definedName>' +
        '<definedName name="Values" comment="The values &amp; more" description="kept">Data!$A$1:$A$3</definedName>' +
        '<definedName name="Rate">Data!$B$1</definedName>' +
        '<definedName name="Here">Data!$A$1</definedName>' +
        '<definedName name="Here" localSheetId="1">Calc!$C$1</definedName>' +
        '<definedName name="Secret" hidden="1">Data!$A$2</definedName>' +
        '<definedName name="Twice">Rate*2</definedName>' +
        '</definedNames>',
    ),
  );
  return zipSync(files);
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

const DATA = 0;
const CALC = 1;
const open = (b = bytes, options = {}) => WorkbookSession.open(b, 'names.xlsx', 'en-US', options);
const reopen = (s: WorkbookSession) => open(s.save());
const name = (s: WorkbookSession, change: NameChange) => s.apply({ type: 'name', change });
const define = (s: WorkbookSession, n: DefinedName) => name(s, { to: n });
const undo = (s: WorkbookSession, r: EditResult) => s.apply({ type: 'restore', cells: r.undo, names: r.names?.undo });
const redo = (s: WorkbookSession, r: EditResult) => s.apply({ type: 'restore', cells: r.redo, names: r.names?.applied });
const rejected = (r: EditResult) => r.rejected && english.text(r.rejected);
const notes = (s: WorkbookSession) => s.changes().map((c) => [c.sheet, c.ref, c.note && english.text(c.note)]);
const sheetNames = ['Data', 'Calc', 'Other Sheet'];

/** What a cell shows now: the engine's value where it calculated one, else the file's. */
function shown(s: WorkbookSession, sheet: string, ref: string): unknown {
  const i = sheetIndex(s, sheet);
  const at = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const col = at[1].charCodeAt(0) - 65;
  const u = s.stateUpdates().find((x) => x.sheet === i && x.row === Number(at[2]) - 1 && x.col === col);
  return u ? valueOf(u) : loadedValue(s, sheet, ref);
}

function formulaAt(s: WorkbookSession, sheet: number, row: number, col: number): string | undefined {
  const c = s.current(sheet, row, col).content;
  return c.kind === 'formula' ? c.formula : undefined;
}

describe('names as Excel defines them', () => {
  it('must start with a letter, underscore or backslash and hold only letters, digits, periods and underscores', () => {
    for (const ok of ['Rate', '_x', '\\back', 'Été', 'a.b_c9', 'Year2026', 'XFE1', 'Q1Sales', 'x'.repeat(255)]) expect(nameProblem(ok), ok).toBeUndefined();
    const problems = ['', '1abc', '.x', 'has space', 'a-b', 'x'.repeat(256), 'A1', 'xfd1048576', 'R1C1', 'r', 'C', 'RC', 'R2', 'c33', 'TRUE', 'false', '_xlnm.Print_Area', '_xlfn.X'].map((n) =>
      english.text(nameProblem(n)!),
    );
    expect(problems).toEqual([
      'Type a name.',
      'A name must start with a letter, an underscore (_) or a backslash (\\).',
      'A name must start with a letter, an underscore (_) or a backslash (\\).',
      "A name can't contain spaces; use an underscore (_) or a period (.) instead.",
      'A name can only contain letters, digits, periods (.) and underscores (_), not "-".',
      'A name can be at most 255 characters long.',
      '"A1" is a cell reference, so it can\'t be a name.',
      '"xfd1048576" is a cell reference, so it can\'t be a name.',
      '"R1C1" is a cell reference, so it can\'t be a name.',
      '"r" is a cell reference, so it can\'t be a name.',
      '"C" is a cell reference, so it can\'t be a name.',
      '"RC" is a cell reference, so it can\'t be a name.',
      '"R2" is a cell reference, so it can\'t be a name.',
      '"c33" is a cell reference, so it can\'t be a name.',
      '"TRUE" is reserved by Excel, so it can\'t be a name.',
      '"false" is reserved by Excel, so it can\'t be a name.',
      '"_xlnm.Print_Area" is reserved by Excel, so it can\'t be a name.',
      '"_xlfn.X" is reserved by Excel, so it can\'t be a name.',
    ]);
  });

  it('are found as a formula on a sheet reads them: the sheet’s own first, then the workbook’s, ignoring case', () => {
    const names: DefinedName[] = [
      { name: 'Here', formula: 'Data!$A$1' },
      { name: 'Here', formula: 'Calc!$C$1', localSheet: 1 },
      { name: 'Only', formula: 'Calc!$D$1', localSheet: 1 },
    ];
    expect(findName(names, 'here', 0, sheetNames)?.formula).toBe('Data!$A$1');
    expect(findName(names, 'HERE', 1, sheetNames)?.formula).toBe('Calc!$C$1');
    expect(findName(names, 'Calc!Here', 0, sheetNames)?.formula).toBe('Calc!$C$1');
    expect(findName(names, "'Other Sheet'!Here", 0, sheetNames)?.formula).toBe('Data!$A$1');
    expect(findName(names, 'Only', 0, sheetNames)).toBeUndefined();
    expect(findName(names, 'Nowhere!Here', 0, sheetNames)).toBeUndefined();
    expect(findName(names, 'A1', 0, sheetNames)).toBeUndefined();
    expect(findName(names, '[', 0, sheetNames)).toBeUndefined();
  });

  it('refer to cells when they are plain ranges, and to nothing that can be selected otherwise', () => {
    const at = (formula: string, localSheet?: number) => nameTarget({ formula, localSheet }, sheetNames, 2);
    expect(at('Data!$A$1:$B$3')).toEqual({ sheet: 0, rects: [{ r1: 0, c1: 0, r2: 2, c2: 1 }] });
    expect(at("'Other Sheet'!$B$3:$A$1")).toEqual({ sheet: 2, rects: [{ r1: 0, c1: 0, r2: 2, c2: 1 }] });
    expect(at('calc!$C$1,Calc!$E$2:$E$4')).toEqual({ sheet: 1, rects: [{ r1: 0, c1: 2, r2: 0, c2: 2 }, { r1: 1, c1: 4, r2: 3, c2: 4 }] });
    expect(at('Data!$B:$C')).toEqual({ sheet: 0, rects: [{ r1: 0, c1: 1, r2: 1_048_575, c2: 2 }] });
    expect(at('Data!$2:$2')).toEqual({ sheet: 0, rects: [{ r1: 1, c1: 0, r2: 1, c2: 16_383 }] });
    expect(at('$A$1', 1)).toEqual({ sheet: 1, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] });
    expect(at('$A$1')).toEqual({ sheet: 2, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] });
    for (const f of ['0.2', '"text"', 'Rate*2', 'OFFSET(Data!$A$1,0,0,3,1)', 'Data!$A$1,Calc!$A$1', 'Data!$A$1 Data!$A$2', 'Nowhere!$A$1', '[1]Data!$A$1', 'Data!$A$1,', '', '#REF!', 'Data!Rate', 'Data:Calc!$A$1', '{1,2}']) {
      expect(at(f), f).toBeUndefined();
    }
  });

  it('name the selection that is exactly what one refers to, the sheet’s own first, then alphabetically', () => {
    const t = (name: string, formula: string, localSheet?: number) => ({ name: { name, formula, localSheet }, target: nameTarget({ formula, localSheet }, sheetNames, 0)! });
    const targets = [t('Zed', 'Data!$A$1:$A$3'), t('Alpha', 'Data!$A$1:$A$3'), t('Mine', 'Data!$A$1:$A$3', 0), t('Elsewhere', 'Data!$A$1:$A$3', 1), t('Two', 'Data!$C$1,Data!$E$1')];
    expect(nameOfSelection(targets, 0, [{ r1: 0, c1: 0, r2: 2, c2: 0 }])).toBe('Mine');
    expect(nameOfSelection(targets.slice(0, 2), 0, [{ r1: 0, c1: 0, r2: 2, c2: 0 }])).toBe('Alpha');
    expect(nameOfSelection(targets, 0, [{ r1: 0, c1: 4, r2: 0, c2: 4 }, { r1: 0, c1: 2, r2: 0, c2: 2 }])).toBe('Two');
    expect(nameOfSelection(targets, 0, [{ r1: 0, c1: 0, r2: 1, c2: 0 }])).toBeUndefined();
    expect(nameOfSelection(targets, 1, [{ r1: 0, c1: 0, r2: 2, c2: 0 }])).toBeUndefined();
  });

  it('are filled in from the selection with absolute references, and refers-to formulas are checked and qualified', () => {
    expect(absoluteRef({ r1: 0, c1: 0, r2: 0, c2: 0 })).toBe('$A$1');
    expect(absoluteRef({ r1: 1, c1: 1, r2: 3, c2: 2 })).toBe('$B$2:$C$4');
    expect(absoluteRef({ r1: 0, c1: 1, r2: 1_048_575, c2: 1 })).toBe('$B:$B');
    expect(absoluteRef({ r1: 4, c1: 0, r2: 5, c2: 16_383 })).toBe('$5:$6');
    expect(selectionFormula('Other Sheet', [{ r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 3, c1: 3, r2: 3, c2: 3 }])).toBe("='Other Sheet'!$A$1:$A$2,'Other Sheet'!$D$4");

    const problem = (text: string) => english.text(refersToProblem(text, sheetNames) ?? '');
    expect(problem('=Data!$A$1:$A$3')).toBe('');
    expect(problem('  =  ')).toBe('Type what the name refers to, such as =Sheet1!$A$1:$A$10.');
    expect(problem('=SUM(')).toBe("This isn't a formula Excel can read.");
    expect(problem('=Missing!$A$1+1')).toBe('This workbook has no sheet called "Missing".');
    expect(problem('=[1]Book!$A$1')).toBe('');
    expect(problem('=0.2')).toBe('');

    expect(refersToFormula('=$A$1:$B$2', 'Data')).toBe('Data!$A$1:$B$2');
    expect(refersToFormula(' = XLOOKUP(1,$A:$A,Calc!$B:$B)+Rate', 'Other Sheet')).toBe("_xlfn.XLOOKUP(1,'Other Sheet'!$A:$A,Calc!$B:$B)+Rate");
    expect(refersToFormula('=SUM(', 'Data')).toBe('SUM(');
    expect(refersToFormula('="a', 'Data')).toBe('"a');
  });

  it('list the ones people define, not hidden names or Excel’s own', () => {
    expect([{ name: 'A', formula: '1' }, { name: 'B', formula: '1', hidden: true }, { name: '_xlnm.Print_Area', formula: '1' }, { name: '_xlfn.SINGLE', formula: '1' }].filter(isListedName).map((n) => n.name)).toEqual(['A']);
  });
});

describe('reading names', () => {
  it('keeps their scope, visibility, comment and place in the file', () => {
    const s = open();
    expect(s.loadResult().names).toEqual([
      { name: '_xlnm.Print_Area', formula: 'Data!$A$1:$B$3', localSheet: 0, fileIndex: 0 },
      { name: 'Values', formula: 'Data!$A$1:$A$3', comment: 'The values & more', fileIndex: 1 },
      { name: 'Rate', formula: 'Data!$B$1', fileIndex: 2 },
      { name: 'Here', formula: 'Data!$A$1', fileIndex: 3 },
      { name: 'Here', formula: 'Calc!$C$1', localSheet: 1, fileIndex: 4 },
      { name: 'Secret', formula: 'Data!$A$2', hidden: true, fileIndex: 5 },
      { name: 'Twice', formula: 'Rate*2', fileIndex: 6 },
    ]);
    // With nothing changed, saving gives the file back as it was.
    expect(s.save()).toBe(bytes);
  });
});

describe('defining a name', () => {
  it('calculates the formulas that were waiting for it, saves it, and undoes', () => {
    const s = open();
    s.apply({ type: 'input', sheet: CALC, row: 0, col: 1, text: '=Extra*2' });
    expect(shown(s, 'Calc', 'B1')).toEqual({ error: '#NAME?' });
    const r = define(s, { name: 'Extra', formula: 'Data!$A$3' });
    expect(r.rejected).toBeUndefined();
    expect(shown(s, 'Calc', 'B1')).toBe(6);
    expect(r.names?.list.map((n) => n.name)).toContain('Extra');
    expect(notes(s)).toContainEqual(['', 'Extra', 'Defined the name Extra as =Data!$A$3']);

    const xml = parts(s.save())['xl/workbook.xml'];
    expect(xml).toContain('<definedName name="Twice">Rate*2</definedName><definedName name="Extra">Data!$A$3</definedName></definedNames>');
    // The names the file had are kept as they were written.
    expect(xml).toContain('<definedName name="Values" comment="The values &amp; more" description="kept">Data!$A$1:$A$3</definedName>');
    const again = reopen(s);
    expect(again.model.names.find((n) => n.name === 'Extra')).toMatchObject({ formula: 'Data!$A$3' });
    expect(shown(again, 'Calc', 'B1')).toBe(6);

    undo(s, r);
    expect(shown(s, 'Calc', 'B1')).toEqual({ error: '#NAME?' });
    expect(s.model.names.some((n) => n.name === 'Extra')).toBe(false);
    redo(s, r);
    expect(shown(s, 'Calc', 'B1')).toBe(6);
  });

  it('local to a sheet, where it comes before the workbook’s name of the same name', () => {
    const s = open();
    s.apply({ type: 'input', sheet: DATA, row: 0, col: 3, text: '=Rate' });
    s.apply({ type: 'input', sheet: CALC, row: 0, col: 3, text: '=Rate' });
    expect(define(s, { name: 'rate', formula: 'Data!$A$3', localSheet: DATA }).rejected).toBeUndefined();
    expect(shown(s, 'Data', 'D1')).toBe(3);
    expect(shown(s, 'Calc', 'D1')).toBe(0.2);
    const saved = parts(s.save())['xl/workbook.xml'];
    expect(saved).toContain('<definedName name="rate" localSheetId="0">Data!$A$3</definedName>');
    expect(reopen(s).model.names.filter((n) => n.name.toLowerCase() === 'rate')).toHaveLength(2);
  });

  it('is refused when the name is invalid, taken in its scope, or its sheet or formula is wrong', () => {
    const s = open();
    expect(rejected(define(s, { name: 'A1', formula: '1' }))).toBe('"A1" is a cell reference, so it can\'t be a name.');
    expect(rejected(define(s, { name: 'VALUES', formula: '1' }))).toBe('The workbook already has a name "Values".');
    expect(rejected(define(s, { name: 'here', formula: '1', localSheet: CALC }))).toBe('Sheet Calc already has a name "Here".');
    expect(rejected(define(s, { name: 'Fine', formula: '1', localSheet: 9 }))).toBe("The sheet that name belongs to isn't in this workbook.");
    expect(rejected(define(s, { name: 'Fine', formula: 'SUM(' }))).toBe("This isn't a formula Excel can read.");
    expect(rejected(name(s, { from: { name: 'Nope' }, to: null }))).toBe('The name "Nope" no longer exists.');
    expect(s.editedCount).toBe(0);
  });

  // HyperFormula can't take a name like Year2026, which Excel and the native engine can.
  it.skipIf(process.env.ENGINE === 'native')('warns when the engine can’t take it, and keeps it for Excel', () => {
    const s = open();
    const r = define(s, { name: 'Year2026', formula: 'Data!$A$1' });
    expect(r.warnings.map((w) => english.text(w))).toEqual(["Formulas that use Year2026 aren't calculated here; Excel calculates them when it opens the file."]);
    expect(parts(s.save())['xl/workbook.xml']).toContain('<definedName name="Year2026">Data!$A$1</definedName>');
  });

  it('adds the definedNames element where the schema puts it when the file has none', () => {
    const s = WorkbookSession.open(sample('budget.xlsx'), 'budget.xlsx');
    const sheet = s.model.sheets[0].info.name;
    define(s, { name: 'Top', formula: `'${sheet}'!$A$1`, comment: 'Line one\nline "two" & <three>' });
    const xml = parts(s.save())['xl/workbook.xml'];
    expect(xml.indexOf('<definedNames>')).toBeGreaterThan(xml.indexOf('</sheets>'));
    expect(xml.indexOf('<definedNames>')).toBeLessThan(xml.indexOf('<calcPr'));
    expect(xml).toContain('comment="Line one&#10;line &quot;two&quot; &amp; &lt;three&gt;"');
    const again = WorkbookSession.open(s.save(), 'budget.xlsx');
    expect(again.model.names).toEqual([{ name: 'Top', formula: `'${sheet}'!$A$1`, comment: 'Line one\nline "two" & <three>', fileIndex: 0 }]);
  });
});

describe('changing a name', () => {
  it('recalculates what uses it, directly or through other names', () => {
    const s = open();
    const r = name(s, { from: { name: 'rate' }, to: { name: 'Rate', formula: 'Data!$A$2' } });
    expect([shown(s, 'Calc', 'A2'), shown(s, 'Calc', 'A4')]).toEqual([200, 4]);
    expect(notes(s)).toEqual([['', 'Rate', 'Changed the name Rate to =Data!$A$2']]);
    expect(parts(s.save())['xl/workbook.xml']).toContain('<definedName name="Rate">Data!$A$2</definedName>');
    undo(s, r);
    expect([shown(s, 'Calc', 'A2'), shown(s, 'Calc', 'A4')]).toEqual([20, 0.4]);
    expect(s.editedCount).toBe(0);
  });

  it('keeps the attributes it doesn’t change, and notes a new comment', () => {
    const s = open();
    name(s, { from: { name: 'Values' }, to: { name: 'Values', formula: 'Data!$A$1:$A$3', comment: 'New' } });
    expect(notes(s)).toEqual([['', 'Values', 'Changed the comment on the name Values']]);
    expect(parts(s.save())['xl/workbook.xml']).toContain('<definedName name="Values" comment="New" description="kept">Data!$A$1:$A$3</definedName>');
    // A hidden name stays hidden, and its comment can go.
    name(s, { from: { name: 'Secret' }, to: { name: 'Secret', formula: 'Data!$A$3' } });
    name(s, { from: { name: 'Values' }, to: { name: 'Values', formula: 'Data!$A$1:$A$3' } });
    const xml = parts(s.save())['xl/workbook.xml'];
    expect(xml).toContain('<definedName name="Secret" hidden="1">Data!$A$3</definedName>');
    expect(xml).toContain('<definedName name="Values" description="kept">Data!$A$1:$A$3</definedName>');
  });

  it('renamed, renames it in the formulas and names that use it, leaving a sheet’s own name of the same name', () => {
    const s = open();
    const r = name(s, { from: { name: 'Here' }, to: { name: 'There', formula: 'Data!$A$1' } });
    // Data!C1 used the workbook's Here; Calc's formulas use Calc's own.
    expect(formulaAt(s, DATA, 0, 2)).toBe('There');
    expect([formulaAt(s, CALC, 2, 0), formulaAt(s, CALC, 4, 0)]).toEqual(['Here', 'Calc!Here*2']);
    expect([shown(s, 'Data', 'C1'), shown(s, 'Calc', 'A3')]).toEqual([1, 5]);
    expect(notes(s)).toContainEqual(['', 'There', 'Renamed Here to There, which refers to =Data!$A$1']);

    // Renaming Calc's own Here reaches the formula that names its sheet.
    const local = name(s, { from: { name: 'Here', localSheet: CALC }, to: { name: 'Mine', formula: 'Calc!$C$1', localSheet: CALC } });
    expect([formulaAt(s, CALC, 2, 0), formulaAt(s, CALC, 4, 0)]).toEqual(['Mine', 'Calc!Mine*2']);
    expect(shown(s, 'Calc', 'A3')).toBe(5);
    undo(s, local);
    expect([formulaAt(s, CALC, 2, 0), formulaAt(s, CALC, 4, 0)]).toEqual(['Here', 'Calc!Here*2']);
    expect(shown(s, 'Calc', 'A3')).toBe(5);

    const rate = name(s, { from: { name: 'Rate' }, to: { name: 'Tax', formula: 'Data!$B$1' } });
    expect(s.model.names.find((n) => n.name === 'Twice')?.formula).toBe('Tax*2');
    expect([formulaAt(s, CALC, 1, 0), shown(s, 'Calc', 'A2'), shown(s, 'Calc', 'A4')]).toEqual(['Tax*100', 20, 0.4]);
    undo(s, rate);
    undo(s, r);
    expect(s.model.names.find((n) => n.name === 'Twice')?.formula).toBe('Rate*2');
    expect(formulaAt(s, DATA, 0, 2)).toBe('Here');
    expect(s.editedCount).toBe(0);
    expect(s.save()).toBe(bytes);
  });
});

describe('deleting a name', () => {
  it('leaves #NAME? in the formulas that use it, and comes back in its place on undo', () => {
    const s = open();
    const r = name(s, { from: { name: 'Values' }, to: null });
    expect(shown(s, 'Calc', 'A1')).toEqual({ error: '#NAME?' });
    expect(notes(s)).toEqual([['', 'Values', 'Deleted the name Values']]);
    const xml = parts(s.save())['xl/workbook.xml'];
    expect(xml).not.toContain('Values');
    expect(xml).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$3</definedName><definedName name="Rate">');
    expect(formulaAt(reopen(s), CALC, 0, 0)).toBe('SUM(Values)');

    undo(s, r);
    expect(shown(s, 'Calc', 'A1')).toBe(6);
    expect(s.editedCount).toBe(0);
    expect(s.save()).toBe(bytes);
  });

  it('removes the definedNames element with the last name', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    name(s, { from: { name: 'TaxRate' }, to: null });
    expect(parts(s.save())['xl/workbook.xml']).not.toContain('definedName');
  });
});

describe('names and the rest of the workbook', () => {
  it('follow rows inserted above what they refer to, and keep their changes in the change list', () => {
    const s = open();
    define(s, { name: 'Pair', formula: 'Data!$A$2:$A$3' });
    const next = s.restructure({ type: 'insertRows', sheet: DATA, at: 0, count: 1 });
    expect(next.model.names.find((n) => n.name === 'Pair')?.formula).toBe('Data!$A$3:$A$4');
    expect(next.changes().map((c) => c.note && english.text(c.note))).toContain('Defined the name Pair as =Data!$A$2:$A$3');
    expect(next.editedCount).toBe(2);
  });

  it('change without live recalculation, and are saved', () => {
    const s = open(bytes, { recalc: 'off' });
    const r = name(s, { from: { name: 'Rate' }, to: { name: 'Rate', formula: 'Data!$A$3' } });
    expect(r.updates).toEqual([]);
    expect(reopen(s).model.names.find((n) => n.name === 'Rate')?.formula).toBe('Data!$A$3');
  });

  it('are renamed only where a formula means them', () => {
    const names: DefinedName[] = [{ name: 'N', formula: '1' }, { name: 'N', formula: '2', localSheet: 1 }];
    const global = { name: 'N' };
    expect(renameInFormula('N+Data!N+Calc!N+n2+"N"', global, 'M', 0, names, sheetNames)).toBe('M+Data!M+Calc!N+n2+"N"');
    expect(renameInFormula('N+Data!N+Calc!N', { name: 'N', localSheet: 1 }, 'M', 1, names, sheetNames)).toBe('M+Data!N+Calc!M');
    expect(renameInFormula('N+1', global, 'M', undefined, names, sheetNames)).toBe('M+1');
    expect(renameInFormula('SUM(A1)', global, 'M', 0, names, sheetNames)).toBeUndefined();
    expect(renameInFormula('N+[1]Book!N+Nowhere!N', global, 'M', 0, names, sheetNames)).toBe('M+[1]Book!N+Nowhere!N');
    expect(renameInFormula('"N', global, 'M', 0, names, sheetNames)).toBeUndefined();
  });

  it('are written afresh when the file’s elements can’t be matched to them', () => {
    const xml = '<workbook><sheets><sheet name="S"/></sheets><definedNames><definedName name="A">1</definedName></definedNames></workbook>';
    expect(writeDefinedNames(xml, [], [{ name: 'B', formula: '2' }])).toBe('<workbook><sheets><sheet name="S"/></sheets><definedNames><definedName name="B">2</definedName></definedNames></workbook>');
    const prefixed = '<x:workbook><x:sheets><x:sheet name="S"/></x:sheets><x:definedNames/></x:workbook>';
    expect(writeDefinedNames(prefixed, [], [{ name: 'B', formula: '2', hidden: true }])).toBe('<x:workbook><x:sheets><x:sheet name="S"/></x:sheets><x:definedNames><x:definedName name="B" hidden="1">2</x:definedName></x:definedNames></x:workbook>');
    expect(writeDefinedNames('<workbook/>', [], [{ name: 'B', formula: '2' }])).toBe('<workbook/>');
    expect(writeDefinedNames('<workbook><sheets/></workbook>', [], [])).toBe('<workbook><sheets/></workbook>');
    expect(writeDefinedNames('<workbook><sheets/><externalReferences><externalReference/></externalReferences></workbook>', [], [{ name: 'B', formula: '2' }])).toBe(
      '<workbook><sheets/><externalReferences><externalReference/></externalReferences><definedNames><definedName name="B">2</definedName></definedNames></workbook>',
    );
  });
});

describe('the engine', () => {
  it('takes names, changes and removes them, and leaves out what it can’t hold', () => {
    const e = new CalcEngine([{ name: 'S', rows: [[2, '=N*3']] }, null], [], false);
    expect(e.setName('N', undefined, 'S!$A$1')).toEqual({ changes: [{ sheet: 0, row: 0, col: 1, value: 6 }], held: true });
    expect(e.setName('n', undefined, '5').changes).toEqual([{ sheet: 0, row: 0, col: 1, value: 15 }]);
    expect(e.setName('N', undefined, '#REF!')).toMatchObject({ held: false, changes: [{ value: { error: '#NAME?' } }] });
    expect(e.setName('N', undefined, null)).toEqual({ changes: [], held: false });
    expect(e.setName('Year2026', undefined, '1').held).toBe(false);
    expect(e.setName('Local', 1, '1')).toEqual({ changes: [], held: false });
    expect(e.setName('Local', 0, '7')).toMatchObject({ held: true });
    expect(e.calculate('=Local', 0)).toBe(7);
    e.destroy();
  });
});
