import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parseA1 } from '../src/core/refs';
import { english } from '../src/i18n';
import {
  CellFlag,
  has,
  loadedFormula,
  loadedValue,
  open,
  parts,
  sample,
  sheetIndex,
  updateAt,
  valueOf,
} from './helpers';

const at = (ref: string) => parseA1(ref)!;

async function readWithExcelJS(bytes: Uint8Array) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes) as never);
  return wb;
}

describe('kitchen-sink.xlsx', () => {
  it('lists every sheet, including the chart sheet and hidden sheet', () => {
    const s = open('kitchen-sink.xlsx');
    expect(s.model.sheets.map((m) => [m.info.name, m.info.kind, m.info.state])).toEqual([
      ['Summary', 'worksheet', 'visible'],
      ['Data', 'worksheet', 'visible'],
      ['Lookups', 'worksheet', 'visible'],
      ['Notes', 'worksheet', 'visible'],
      ['Reference', 'worksheet', 'visible'],
      ['Settings', 'worksheet', 'hidden'],
      ['Monthly Chart', 'chartsheet', 'visible'],
    ]);
  });

  it('reports what the editor keeps but cannot show', () => {
    const s = open('kitchen-sink.xlsx');
    const kinds = s.model.findings.map((f) => f.kind);
    // Frozen rows are shown frozen, so there is nothing to report about them.
    expect(kinds).not.toContain('frozenRows');
    for (const k of [
      'chart', 'chartsheet', 'table', 'conditionalFormat', 'dataValidation', 'comment', 'sparkline',
      'image', 'shape', 'hiddenSheet', 'protection', 'arrayFormula', 'richText', 'hyperlink', 'mergedCells',
    ]) {
      expect(kinds, k).toContain(k);
    }
    const charts = s.model.findings.filter((f) => f.kind === 'chart').map((f) => english.text(f.title));
    expect(charts).toEqual(['Chart "Revenue by region"', 'Chart "Share of revenue"']);
    expect(english.text(s.model.findings.find((f) => f.kind === 'chartsheet')?.detail)).toContain('Monthly revenue 2026');
  });

  it('reads values, formats and formulas', () => {
    const s = open('kitchen-sink.xlsx');
    expect(loadedValue(s, 'Summary', 'A1')).toBe('Regional Sales Summary 2026');
    expect(loadedFormula(s, 'Summary', 'F9')).toBe('SUM(F5:F8)');
    expect(loadedFormula(s, 'Lookups', 'B4')).toBe('_xlfn.XLOOKUP(A4,$F$4:$F$6,$G$4:$G$6)');
    const summary = s.model.sheets[0];
    const style = s.model.styles.get(summary.scan!.grid.styles[1]);
    expect(style.bold).toBe(true);
    expect(summary.info.merges).toContainEqual({ r1: 0, c1: 0, r2: 0, c2: 6 });
    expect(summary.info.frozenCols).toBe(1);
  });

  it("builds the engine and agrees with Excel's saved values", () => {
    const s = open('kitchen-sink.xlsx');
    const report = s.ensureEngine();
    console.log(JSON.stringify({ ...report, updates: report.updates.length }, null, 1));
    expect(report.mismatched).toBe(0);
    const unsupported = report.unsupportedFunctions.map((x) => english.text(x)).join(' ');
    expect(unsupported).toMatch(/LET/);
    expect(unsupported).toMatch(/CONCAT/);
    expect(unsupported).toMatch(/Table references/);
    expect(report.liveCells).toBeGreaterThan(200);
  });

  it('recalculates dependents and marks formulas it cannot vouch for as stale', () => {
    const s = open('kitchen-sink.xlsx');
    s.ensureEngine();
    const data = sheetIndex(s, 'Data');
    const before = loadedValue(s, 'Summary', 'F9') as number;
    const units = loadedValue(s, 'Data', 'D2') as number;
    const price = loadedValue(s, 'Data', 'E2') as number;
    const r = s.apply({ type: 'input', sheet: data, ...at('D2'), text: String(units + 10) });
    expect(r.rejected).toBeUndefined();
    expect(valueOf(updateAt(r.updates, s, 'Data', 'F2'))).toBeCloseTo((units + 10) * price);
    expect(valueOf(updateAt(r.updates, s, 'Summary', 'F9'))).toBeCloseTo(before + 10 * price);
    // LET is not supported by the engine: it keeps Excel's value and turns stale,
    // and so does the formula that reads it.
    expect(has(updateAt(r.updates, s, 'Summary', 'B16')?.flags, CellFlag.Stale)).toBe(true);
    expect(has(updateAt(r.updates, s, 'Summary', 'B17')?.flags, CellFlag.Stale)).toBe(true);
    expect(has(updateAt(r.updates, s, 'Lookups', 'B8')?.flags, CellFlag.Stale)).toBe(true);
    expect(r.undo).toHaveLength(1);
  });

  it('marks a whole chain stale, through formulas it can’t calculate and the ones it can', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Chain');
    ws.getCell('A1').value = 1;
    // B1 and C1 use functions the engine doesn't have; D1 and F1 it calculates.
    ws.getCell('B1').value = { formula: '_xludf.FOO(A1)', result: 2 };
    ws.getCell('C1').value = { formula: '_xludf.BAR(B1)', result: 20 };
    ws.getCell('D1').value = { formula: 'C1*10', result: 200 };
    ws.getCell('E1').value = { formula: '_xludf.BAZ(D1)', result: 7 };
    ws.getCell('F1').value = { formula: 'E1+1', result: 8 };
    ws.getCell('G1').value = { formula: 'A1+1', result: 2 };
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 'chain.xlsx');
    s.ensureEngine();
    const r = s.apply({ type: 'input', sheet: 0, ...at('A1'), text: '5' });
    const stale = (ref: string) => has(updateAt(r.updates, s, 'Chain', ref)?.flags, CellFlag.Stale);
    expect(['B1', 'C1', 'D1', 'E1', 'F1'].map(stale)).toEqual([true, true, true, true, true]);
    // A formula the engine calculates from the changed cell is up to date.
    expect(stale('G1')).toBe(false);
    expect(valueOf(updateAt(r.updates, s, 'Chain', 'G1'))).toBe(6);
  });

  it("reads typed and pasted dates into a workbook's 1904 date system", () => {
    const s = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'dates-1904.xlsx'))), 'dates.xlsx', 'en-US', { recalc: 'on' });
    const may1 = (Date.UTC(2026, 4, 1) - Date.UTC(1904, 0, 1)) / 86_400_000;
    const typed = s.apply({ type: 'input', sheet: 0, ...at('A2'), text: '2026-05-01' });
    expect(valueOf(updateAt(typed.updates, s, 'Schedule', 'A2'))).toBe(may1);
    // The end date that adds days to it moves with it.
    expect(valueOf(updateAt(typed.updates, s, 'Schedule', 'C2'))).toBe(may1 + 10);
    // A value copied from a 1900-system workbook keeps the date its text shows.
    const pasted = s.apply({ type: 'pasteText', sheet: 0, ...at('A3'), values: [['5/1/2026']], exact: [[46143]] });
    expect(valueOf(updateAt(pasted.updates, s, 'Schedule', 'A3'))).toBe(may1);
  });

  it('saves only what changed and the result reopens with the edit', async () => {
    const s = open('kitchen-sink.xlsx');
    s.ensureEngine();
    const data = sheetIndex(s, 'Data');
    s.apply({ type: 'input', sheet: data, ...at('D2'), text: '500' });
    s.apply({ type: 'input', sheet: data, ...at('H2'), text: 'new note' });
    const bytes = s.save();

    const before = parts(sample('kitchen-sink.xlsx'));
    const after = parts(bytes);
    const changed = Object.keys(after).filter((k) => after[k] !== before[k]).sort();
    // Summary and Lookups change too: their formulas' saved results are updated.
    expect(changed).toEqual(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml']);
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    expect(after['xl/workbook.xml']).toContain('fullCalcOnLoad="1"');

    const reopened = WorkbookSession.open(bytes, 'x.xlsx');
    expect(loadedValue(reopened, 'Data', 'D2')).toBe(500);
    expect(loadedValue(reopened, 'Data', 'H2')).toBe('new note');
    expect(loadedFormula(reopened, 'Data', 'F2')).toBe('D2*E2');
    expect(loadedValue(reopened, 'Data', 'F2')).toBeCloseTo(500 * (loadedValue(s, 'Data', 'E2') as number));
    expect(reopened.model.findings.filter((f) => f.kind === 'chart')).toHaveLength(2);

    const wb = await readWithExcelJS(bytes);
    expect(wb.getWorksheet('Data')!.getCell('D2').value).toBe(500);
    expect((wb.getWorksheet('Data')!.getCell('F2').value as { formula: string }).formula).toBe('D2*E2');
  });

  it('returns the original bytes when nothing changed', () => {
    const s = open('kitchen-sink.xlsx');
    s.ensureEngine();
    const data = sheetIndex(s, 'Data');
    const r = s.apply({ type: 'input', sheet: data, ...at('D2'), text: '1' });
    s.apply({ type: 'restore', cells: r.undo });
    expect(s.editedCount).toBe(0);
    expect(Buffer.from(s.save()).equals(Buffer.from(sample('kitchen-sink.xlsx')))).toBe(true);
  });

  it('blocks edits to pivot, table headers, arrays and protected sheets', () => {
    const s = open('kitchen-sink.xlsx');
    const header = s.apply({ type: 'input', sheet: sheetIndex(s, 'Data'), ...at('B1'), text: 'Area' });
    expect(header.skipped).toBe(1);
    expect(english.text(header.skippedReasons[0])).toMatch(/table "SalesData"/);
    const spill = s.apply({ type: 'input', sheet: sheetIndex(s, 'Lookups'), ...at('B10'), text: 'x' });
    expect(english.text(spill.skippedReasons[0])).toMatch(/array formula/);
    const locked = s.apply({ type: 'input', sheet: sheetIndex(s, 'Reference'), ...at('A2'), text: 'x' });
    expect(english.text(locked.skippedReasons[0])).toMatch(/protected/);
  });

  it('rejects malformed formulas and adds _xlfn prefixes to new functions', () => {
    const s = open('kitchen-sink.xlsx');
    const lookups = sheetIndex(s, 'Lookups');
    const bad = s.apply({ type: 'input', sheet: lookups, ...at('E20'), text: '=SUM(A1:A3' });
    expect(english.text(bad.rejected)).toMatch(/problem with this formula/);
    const ok = s.apply({ type: 'input', sheet: lookups, ...at('E20'), text: '=xlookup("Gadget",F4:F6,G4:G6)' });
    const u = updateAt(ok.updates, s, 'Lookups', 'E20');
    expect(u?.f).toBe('_xlfn.xlookup("Gadget",F4:F6,G4:G6)');
    expect(valueOf(u)).toBe(30);
  });
});

describe('pivot.xlsx', () => {
  it('finds the pivot table, blocks its cells and expands shared formulas when the master changes', async () => {
    const s = open('pivot.xlsx');
    expect(english.text(s.model.findings.find((f) => f.kind === 'pivot')?.title)).toBe('Pivot table "UnitsByRegion"');
    const pivot = sheetIndex(s, 'Pivot');
    expect(s.apply({ type: 'input', sheet: pivot, ...at('B4'), text: '1' }).skipped).toBe(1);

    expect(loadedFormula(s, 'Data', 'E3')).toBe('C3*D3');
    const data = sheetIndex(s, 'Data');
    const r = s.apply({ type: 'input', sheet: data, ...at('E2'), text: '=C2*D2*2' });
    expect(valueOf(updateAt(r.updates, s, 'Data', 'E2'))).toBeCloseTo((loadedValue(s, 'Data', 'C2') as number) * (loadedValue(s, 'Data', 'D2') as number) * 2);

    const bytes = s.save();
    const p = parts(bytes);
    // The pivot's data changed, so Excel is asked to refresh it on opening.
    expect(p['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('refreshOnLoad="1"');
    const sheet1 = p['xl/worksheets/sheet1.xml'];
    expect(sheet1).not.toContain('t="shared"');
    const wb = await readWithExcelJS(bytes);
    expect((wb.getWorksheet('Data')!.getCell('E50').value as { formula: string }).formula).toBe('C50*D50');
  });
});

describe('editing operations', () => {
  it('parses pasted text like Excel and creates number formats for it', () => {
    const s = open('kitchen-sink.xlsx');
    const notes = sheetIndex(s, 'Notes');
    const r = s.apply({
      type: 'pasteText',
      sheet: notes,
      ...at('A20'),
      values: [['12%', '2026-03-15', '$1,234.50', '=B10*2', 'hello', 'TRUE']],
    });
    const cells = ['A20', 'B20', 'C20', 'D20', 'E20', 'F20'].map((ref) => updateAt(r.updates, s, 'Notes', ref)!);
    expect(cells.map(valueOf)).toEqual([0.12, 46096, 1234.5, 24, 'hello', true]);
    expect(r.newStyles.map((n) => n.style.numFmt)).toEqual(['0%', 'yyyy-mm-dd', '$#,##0.00']);
    const bytes = s.save();
    const styles = parts(bytes)['xl/styles.xml'];
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
  });

  it('fills series and shifts formulas when dragging', () => {
    const s = open('kitchen-sink.xlsx');
    const notes = sheetIndex(s, 'Notes');
    s.apply({ type: 'pasteText', sheet: notes, ...at('A30'), values: [['1', 'Jan', 'Item 9', '=A30*10']] });
    s.apply({ type: 'pasteText', sheet: notes, ...at('F30'), values: [['1'], ['3']] });
    s.apply({ type: 'fill', sheet: notes, src: { r1: 29, c1: 0, r2: 29, c2: 3 }, dst: { r1: 29, c1: 0, r2: 32, c2: 3 } });
    const r = s.apply({ type: 'fill', sheet: notes, src: { r1: 29, c1: 5, r2: 30, c2: 5 }, dst: { r1: 29, c1: 5, r2: 33, c2: 5 } });
    expect(['F32', 'F33', 'F34'].map((ref) => valueOf(updateAt(r.updates, s, 'Notes', ref)))).toEqual([5, 7, 9]);
    const snap = (ref: string) => s.current(notes, at(ref).row, at(ref).col).content;
    expect(snap('A33')).toEqual({ kind: 'number', value: 1 });
    expect(snap('B33')).toEqual({ kind: 'string', value: 'Apr' });
    expect(snap('C33')).toEqual({ kind: 'string', value: 'Item 12' });
    expect(snap('D33')).toEqual({ kind: 'formula', formula: 'A33*10' });
  });

  it('copies formulas with adjusted references and pastes values over formulas', () => {
    const s = open('kitchen-sink.xlsx');
    const summary = sheetIndex(s, 'Summary');
    const r = s.apply({ type: 'copy', srcSheet: summary, src: { r1: 8, c1: 1, r2: 8, c2: 1 }, sheet: summary, dst: { r1: 9, c1: 1, r2: 9, c2: 3 } });
    expect(updateAt(r.updates, s, 'Summary', 'C10')?.f).toBe('SUM(C6:C9)');
  });
});

describe('engine verdicts reach the UI', () => {
  it('sends flags for formulas the engine cannot calculate', () => {
    const s = open('kitchen-sink.xlsx');
    const report = s.ensureEngine();
    const b8 = updateAt(report.updates, s, 'Lookups', 'B8');
    expect(has(b8?.flags, CellFlag.Unsupported)).toBe(true);
    expect(english.text(report.cellReasons.get(`${sheetIndex(s, 'Lookups')}:7:1`))).toMatch(/table references/);
  });
});

describe('a protected sheet', () => {
  const form = () => WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))), 'templates.xlsx', 'en-US', { recalc: 'on' });

  it('takes input in the cells its format leaves unlocked, and nowhere else', () => {
    const s = form();
    s.ensureEngine();
    const request = sheetIndex(s, 'Request');
    expect(s.apply({ type: 'input', sheet: request, ...at('B3'), text: 'Ana' }).rejected).toBeUndefined();
    const locked = s.apply({ type: 'input', sheet: request, ...at('A3'), text: 'Who' });
    expect([locked.skipped, locked.skippedReasons.map((x) => english.text(x))]).toEqual([1, ['This cell is locked: its sheet is protected in Excel.']]);
    // The budget is an input; the check that compares with it recalculates.
    const r = s.apply({ type: 'input', sheet: request, ...at('B5'), text: '100' });
    expect(valueOf(updateAt(r.updates, s, 'Request', 'B7'))).toBe('No');
    // Rows can't go in or out of it.
    expect(() => s.restructure({ type: 'insertRows', sheet: request, at: 2, count: 1 })).toThrow();
    expect(english.text(s.model.findings.find((f) => f.kind === 'protection')?.detail)).toMatch(/only the cells it leaves unlocked can be edited/);
  });
});
