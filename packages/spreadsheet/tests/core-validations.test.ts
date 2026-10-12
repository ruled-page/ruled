import ExcelJS from 'exceljs';
import { strToU8, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { LayoutOp, ListValidation, RuleValidation } from '../src/core/types';
import { subtractRect, withoutRects } from '../src/core/validations';
import { english } from '../src/i18n';
import { parts } from './helpers';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

describe('ranges with others taken out', () => {
  it('leave up to four pieces around the cut, or the range itself where they miss', () => {
    expect(subtractRect(rect(0, 0, 4, 4), rect(2, 2))).toEqual([rect(0, 0, 1, 4), rect(3, 0, 4, 4), rect(2, 0, 2, 1), rect(2, 3, 2, 4)]);
    expect(subtractRect(rect(0, 0, 2, 0), rect(5, 5))).toEqual([rect(0, 0, 2, 0)]);
    expect(subtractRect(rect(0, 0, 2, 0), rect(0, 0, 9, 9))).toEqual([]);
    expect(withoutRects([rect(0, 0, 3, 0), rect(0, 2, 3, 2)], [rect(1, 0, 2, 2)])).toEqual([rect(0, 0), rect(3, 0), rect(0, 2), rect(3, 2)]);
  });
});

/** Rules!A1:A5 must be whole numbers from 1 to 10, with a prompt; B1:B5 picks from a list; C1 takes any value with a message. */
let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Rules');
  wb.addWorksheet('Lists').getCell('A1').value = 'x';
  for (let r = 1; r <= 5; r++) {
    ws.getCell(`A${r}`).dataValidation = { type: 'whole', operator: 'between', formulae: [1, 10], showInputMessage: true, promptTitle: 'Qty', prompt: 'From 1 to 10' };
    ws.getCell(`B${r}`).dataValidation = { type: 'list', formulae: ['"Yes,No"'], allowBlank: true, showErrorMessage: true, errorStyle: 'warning', error: 'Pick one' };
  }
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

const open = (b = bytes) => WorkbookSession.open(b, 'rules.xlsx', 'en-US');
const set = (s: WorkbookSession, op: Extract<LayoutOp, { kind: 'validations' }>) => s.apply({ type: 'layout', sheet: 0, op });
const info = (s: WorkbookSession) => s.model.sheets[0].info;

describe('changing validation rules', () => {
  it('writes them as Excel does, and reads them back the same', () => {
    const s = open();
    const lists: ListValidation[] = [{ ...info(s).listValidations[0], rects: [rect(0, 1, 2, 1)] }];
    const rules: RuleValidation[] = [
      { ...info(s).ruleValidations[0], operator: 'greaterThan', formula1: '5', formula2: undefined, errorStyle: 'information', showError: true, errorTitle: 'Too few', error: 'Say "more" & <go>' },
      { rects: [rect(0, 2)], kind: 'any', operator: 'between', formula1: '', allowBlank: true, showPrompt: true, showError: false, errorStyle: 'stop', prompt: 'Line one\nLine two' },
    ];
    const r = set(s, { kind: 'validations', lists, rules });
    expect(r.layout?.undo[0].op).toMatchObject({ kind: 'validations', lists: [{ rects: [rect(0, 1, 4, 1)] }] });
    const xml = parts(s.save())['xl/worksheets/sheet1.xml'];
    expect(xml).toContain('<dataValidations count="3">');
    expect(xml).toContain('<dataValidation type="list" errorStyle="warning" allowBlank="1" showErrorMessage="1" error="Pick one" sqref="B1:B3"><formula1>"Yes,No"</formula1></dataValidation>');
    expect(xml).toContain('operator="greaterThan" errorStyle="information"');
    expect(xml).toContain('error="Say &quot;more&quot; &amp; &lt;go&gt;"');
    expect(xml).toContain('<dataValidation type="none" allowBlank="1" showInputMessage="1" prompt="Line one&#10;Line two" sqref="C1"/>');
    const back = info(open(s.save()));
    expect(back.listValidations).toEqual(lists);
    expect(back.ruleValidations).toEqual(rules);
    expect(s.changes().map((c) => english.text(c.note!))).toEqual(['Data validation rules changed']);
  });

  it('can take every rule off a sheet, and put them back on undo, leaving nothing changed', () => {
    const s = open();
    const r = set(s, { kind: 'validations', lists: [], rules: [] });
    expect(parts(s.save())['xl/worksheets/sheet1.xml']).not.toContain('dataValidation');
    s.apply({ type: 'layout', sheet: 0, op: r.layout!.undo[0].op });
    expect(info(s).ruleValidations).toHaveLength(1);
    expect(s.changes()).toEqual([]);
  });

  it("are refused on a sheet with rules the editor can't read, and on a protected sheet", () => {
    const files = parts(bytes);
    // A list with no source, which Excel wouldn't write but a file can hold.
    const kept = files['xl/worksheets/sheet1.xml'].replace(/<dataValidations count="\d+">/, '$&<dataValidation type="list" sqref="D1"/>');
    const s = open(repack(files, { 'xl/worksheets/sheet1.xml': kept }));
    expect(info(s).validationsKept).toBe(true);
    expect(english.text(set(s, { kind: 'validations', lists: [], rules: [] }).rejected!)).toMatch(/can't read/);
    const p = open(repack(files, { 'xl/worksheets/sheet1.xml': files['xl/worksheets/sheet1.xml'].replace('</sheetData>', '</sheetData><sheetProtection sheet="1"/>') }));
    expect(english.text(set(p, { kind: 'validations', lists: [], rules: [] }).rejected!)).toBe('This sheet is protected in Excel, so it is read-only here.');
  });

  it('adds them in their place to a sheet that had none, writing a list from its source where its formula is not known', () => {
    const s = open();
    const alert = { allowBlank: false, showError: false, errorStyle: 'stop' as const, inCellDropdown: true };
    const lists: ListValidation[] = [
      { ...alert, rects: [rect(0, 0)], source: { kind: 'values', values: ['a', 'say "b"'] } },
      { ...alert, rects: [rect(1, 0)], source: { kind: 'formula', formula: 'OFFSET(A1,0,0,3)' } },
      { ...alert, rects: [rect(2, 0)], source: { kind: 'range', sheet: 1, rect: rect(0, 2, 4, 2) } },
    ];
    s.apply({ type: 'layout', sheet: 1, op: { kind: 'validations', lists, rules: [] } });
    const xml = parts(s.save())['xl/worksheets/sheet2.xml'];
    expect(xml).toMatch(/<\/sheetData><dataValidations count="3">/);
    expect(xml).toContain('<formula1>"a,say ""b"""</formula1>');
    expect(xml).toContain('<formula1>OFFSET(A1,0,0,3)</formula1>');
    expect(xml).toContain('<formula1>C1:C5</formula1>');
  });

  it('takes Excel 2010 copies out of the extensions, as <dataValidations> stands in for them', () => {
    const files = parts(bytes);
    const ext =
      '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
      '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1">' +
      '<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>E1:E2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>';
    const s = open(repack(files, { 'xl/worksheets/sheet1.xml': files['xl/worksheets/sheet1.xml'].replace('</worksheet>', `${ext}</worksheet>`) }));
    expect(info(s).listValidations.map((l) => l.formula)).toContain('Lists!$A$1:$A$3');
    set(s, { kind: 'validations', lists: info(s).listValidations, rules: [] });
    const xml = parts(s.save())['xl/worksheets/sheet1.xml'];
    expect(xml).not.toContain('extLst');
    expect(xml).toContain('<formula1>Lists!$A$1:$A$3</formula1>');
  });
});

function repack(files: Record<string, string>, changes: Record<string, string>): Uint8Array {
  const all = { ...files, ...changes };
  return zipFiles(Object.fromEntries(Object.entries(all).filter(([k]) => !k.endsWith('/'))));
}

const zipFiles = (files: Record<string, string>) => zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
