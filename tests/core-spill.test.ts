import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { english, errorText } from '../src/i18n';
import { WorkbookSession } from '../src/core/session';
import { parseA1, parseRange, rangeToA1 } from '../src/core/refs';
import { CellFlag, VT, type CellUpdate, type EditResult } from '../src/core/types';
import { parts, sheetIndex, valueOf } from './helpers';

const at = (ref: string) => parseA1(ref)!;
const rect = (ref: string) => parseRange(ref)!;

/** Sheet S: A1:A6 = 3, 1, 2, 3, 1, 2; C5 = "below". */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  [3, 1, 2, 3, 1, 2].forEach((v, i) => (ws.getCell(`A${i + 1}`).value = v));
  ws.getCell('C5').value = 'below';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const fresh = (from = bytes) => {
  const s = WorkbookSession.open(from, 'w.xlsx', 'en-US', { recalc: 'on' });
  s.ensureEngine();
  return s;
};
const corpus = () => fresh(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))));

/** What the grid shows for a cell after everything so far: its value, whether it is spilled, and the spill it is part of. */
function shown(s: WorkbookSession, ref: string, sheet = 0): { value: unknown; spilled: boolean; spill?: string; flags: number } | null {
  const { row, col } = at(ref);
  const u: CellUpdate | undefined = s.stateUpdates().find((x) => x.sheet === sheet && x.row === row && x.col === col);
  if (!u) return null;
  return { value: valueOf(u), spilled: !!(u.flags & CellFlag.Spilled), ...(u.spill && { spill: rangeToA1(u.spill) }), flags: u.flags };
}
const value = (s: WorkbookSession, ref: string, sheet = 0) => shown(s, ref, sheet)?.value ?? null;
/** The update an edit sent for a cell. */
const sent = (r: EditResult, ref: string) => r.updates.find((u) => u.row === at(ref).row && u.col === at(ref).col);
const column = (s: WorkbookSession, refs: string[], sheet = 0) => refs.map((r) => value(s, r, sheet));
const type = (s: WorkbookSession, ref: string, text: string, sheet = 0) => s.apply({ type: 'input', sheet, ...at(ref), text });

/** A saved workbook with one part changed. */
function patched(saved: Uint8Array, path: string, fn: (xml: string) => string): Uint8Array {
  const files = unzipSync(saved);
  files[path] = strToU8(fn(strFromU8(files[path])));
  return zipSync(files);
}

const cellXml = (xml: string, ref: string) => new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(xml)?.[0];

describe('a formula typed here whose result is an array', () => {
  it('spills its values into the cells below and to the right, which hold nothing of their own', () => {
    const s = fresh();
    const r = type(s, 'C1', '=SEQUENCE(3,2)');
    expect(r.updates.filter((u) => u.flags & CellFlag.Spilled).length).toBe(5);
    expect(shown(s, 'C1')).toMatchObject({ value: 1, spilled: false, spill: 'C1:D3' });
    expect(shown(s, 'D3')).toMatchObject({ value: 6, spilled: true, spill: 'C1:D3' });
    expect(s.current(0, 2, 3).content).toEqual({ kind: 'empty' });
    // Other formulas read the spilled values.
    type(s, 'F1', '=SUM(C2:D3)');
    expect(value(s, 'F1')).toBe(18);
  });

  it('shows #SPILL! while something is in its way, and spills again once that is cleared or undone', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(3)');
    const blocker = type(s, 'C2', 'x');
    expect(shown(s, 'C1')).toMatchObject({ value: { error: '#SPILL!' }, spill: 'C1:C3' });
    expect(sent(blocker, 'C3')).toMatchObject({ t: VT.None, flags: 0 });
    s.apply({ type: 'restore', cells: blocker.undo });
    expect(column(s, ['C1', 'C2', 'C3'])).toEqual([1, 2, 3]);
    s.apply({ type: 'restore', cells: blocker.redo });
    expect(value(s, 'C1')).toEqual({ error: '#SPILL!' });
    s.apply({ type: 'clear', sheet: 0, rects: [rect('C2')] });
    expect(column(s, ['C1', 'C2', 'C3'])).toEqual([1, 2, 3]);
  });

  it('takes only the room its result needs, and grows and shrinks with its inputs', () => {
    const s = fresh();
    // UNIQUE could need six rows; it needs three, and C5 is below them.
    type(s, 'C1', '=UNIQUE(A1:A6)');
    expect(column(s, ['C1', 'C2', 'C3', 'C4'])).toEqual([3, 1, 2, null]);
    type(s, 'A6', '9');
    expect(column(s, ['C1', 'C2', 'C3', 'C4'])).toEqual([3, 1, 2, 9]);
    expect(sent(type(s, 'A6', '1'), 'C4')).toMatchObject({ t: VT.None, flags: 0 });
    // Growing into C5 blocks it, until C5 is cleared.
    type(s, 'A5', '7');
    type(s, 'A6', '8');
    expect(value(s, 'C1')).toEqual({ error: '#SPILL!' });
    s.apply({ type: 'clear', sheet: 0, rects: [rect('C5')] });
    expect(column(s, ['C1', 'C2', 'C3', 'C4', 'C5'])).toEqual([3, 1, 2, 7, 8]);
    // A size worked out from a cell.
    type(s, 'B1', '2');
    type(s, 'E1', '=SEQUENCE(B1)');
    expect(column(s, ['E1', 'E2', 'E3'])).toEqual([1, 2, null]);
    type(s, 'B1', '4');
    expect(column(s, ['E1', 'E2', 'E3', 'E4', 'E5'])).toEqual([1, 2, 3, 4, null]);
  });

  it('goes when the formula is cleared, and comes back when that is undone', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(2)');
    const cleared = s.apply({ type: 'clear', sheet: 0, rects: [rect('C1')] });
    expect(sent(cleared, 'C2')).toMatchObject({ t: VT.None, flags: 0 });
    s.apply({ type: 'restore', cells: cleared.undo });
    expect(shown(s, 'C2')).toMatchObject({ value: 2, spilled: true, spill: 'C1:C2' });
  });

  it('saves with its spill range as the ref and its values in the cells it covers, or #SPILL! in its own cell when blocked', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(2,2)');
    type(s, 'F1', '=SEQUENCE(3)');
    type(s, 'F3', 'in the way');
    const xml = parts(s.save())['xl/worksheets/sheet1.xml'];
    expect(cellXml(xml, 'C1')).toBe('<c r="C1" cm="1"><f t="array" ref="C1:D2">_xlfn.SEQUENCE(2,2)</f><v>1</v></c>');
    expect(cellXml(xml, 'D1')).toBe('<c r="D1"><v>2</v></c>');
    expect(cellXml(xml, 'D2')).toBe('<c r="D2"><v>4</v></c>');
    expect(cellXml(xml, 'F1')).toBe('<c r="F1" cm="1" t="e"><f t="array" ref="F1">_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c>');
    expect(cellXml(xml, 'F2')).toBeUndefined();
    // Opened again, the spill is the file's own, and spills here.
    const again = fresh(s.save());
    expect(again.calcReport?.liveArrays).toEqual([{ sheet: 0, rect: rect('C1:D2') }]);
    expect(again.calcReport?.mismatched).toBe(0);
    expect(shown(again, 'D2')).toMatchObject({ value: 4, spilled: true });
  });
});

describe("a file's own dynamic array", () => {
  it('spills here, unlocked, while one entered with Ctrl+Shift+Enter stays locked', () => {
    const s = corpus();
    const arrays = sheetIndex(s, 'Arrays');
    const info = s.model.sheets[arrays].info;
    // B2:B4 reads the spill as A2#, which HyperFormula can't calculate (the native engine can); G6:G8 is the older kind, calculated but changed only as a whole.
    const native = process.env.ENGINE !== 'hyperformula';
    expect(s.calcReport?.liveArrays?.map((a) => rangeToA1(a.rect))).toEqual(native ? ['A2:A4', 'B2:B4', 'G2:I3', 'E5:E6'] : ['A2:A4', 'G2:I3', 'E5:E6']);
    expect(info.blocked.map((b) => rangeToA1(b))).toEqual(native ? ['G6:G8'] : ['B2:B4', 'G6:G8']);
    expect(column(s, ['A2', 'A3', 'A4'], arrays)).toEqual(['Electrical', 'Hardware', 'Safety']);
    expect(shown(s, 'H3', arrays)).toMatchObject({ value: 5, spilled: true, spill: 'G2:I3' });
    // UNIQUE could need five rows, where A6 has text; it needs three.
    expect(shown(s, 'A2', arrays)).toMatchObject({ spill: 'A2:A4' });
    const r = type(s, 'G7', '1', arrays);
    expect(english.text(r.skippedReasons[0])).toMatch(/array formula/);
    // Typing into a spilled cell blocks the spill, as in Excel.
    type(s, 'H3', 'x', arrays);
    expect(value(s, 'G2', arrays)).toEqual({ error: '#SPILL!' });
  });

  it('follows its inputs, and when saved clears the values it no longer reaches', () => {
    const s = corpus();
    const arrays = sheetIndex(s, 'Arrays');
    const lines = sheetIndex(s, 'Lines');
    // Nuts is now Safety: one Hardware line left.
    type(s, 'B3', 'Safety', lines);
    expect(column(s, ['E5', 'E6'], arrays)).toEqual(['Bolts', null]);
    expect(shown(s, 'E6', arrays)).toMatchObject({ spilled: false });
    const xml = parts(s.save())['xl/worksheets/sheet3.xml'];
    expect(cellXml(xml, 'E5')).toMatch(/<f t="array" ref="E5">_xlfn._xlws.FILTER/);
    expect(cellXml(xml, 'E6')).toBeUndefined();
    // Values that are as saved stay as Excel wrote them.
    expect(cellXml(xml, 'A3')).toBe('<c r="A3" t="s"><v>12</v></c>');
  });

  it('keeps its saved values, locked, when the engine spills it differently', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(3)');
    const saved = patched(s.save(), 'xl/worksheets/sheet1.xml', (xml) => xml.replace('<c r="C3"><v>3</v></c>', '<c r="C3"><v>30</v></c>'));
    const again = fresh(saved);
    expect(again.calcReport?.mismatched).toBe(1);
    expect(again.calcReport?.liveArrays).toBeUndefined();
    expect(again.model.sheets[0].info.blocked.map((b) => rangeToA1(b))).toEqual(['C1:C3']);
    expect(shown(again, 'C1')?.flags ?? 0).toBe(CellFlag.Formula | CellFlag.Mismatch | CellFlag.Unsupported | CellFlag.ArrayMember);
    // Formulas reading it get Excel's values.
    type(again, 'E1', '=C3+1');
    expect(value(again, 'E1')).toBe(31);
  });
});

describe('an array formula entered with Ctrl+Shift+Enter', () => {
  /** S!A1:A3 = 1, 2, 3, with Excel's values for C1:C3 {=A1:A3*2}, D1:E4 {=A1:A3*10} and F1:H1 {=TRANSPOSE(A1:A2)}. */
  async function arrays(second = '=A1:A3*10'): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const cells: Record<string, ExcelJS.CellValue> = { A1: 1, A2: 2, A3: 3, C1: 2, C2: 4, C3: 6, D1: 10, D2: 20, D3: 30, E1: 10, E2: 20, E3: 30, F1: 1, G1: 2 };
    for (const [ref, v] of Object.entries(cells)) ws.getCell(ref).value = v;
    for (const ref of ['D4', 'E4', 'H1']) ws.getCell(ref).value = { error: '#N/A' };
    const formulas: [string, string, string][] = [
      ['C1', 'C1:C3', '=A1:A3*2'],
      ['D1', 'D1:E4', second],
      ['F1', 'F1:H1', '=TRANSPOSE(A1:A2)'],
    ];
    return patched(new Uint8Array(await wb.xlsx.writeBuffer()), 'xl/worksheets/sheet1.xml', (xml) =>
      formulas.reduce((x, [cell, ref, f]) => x.replace(new RegExp(`(<c r="${cell}"[^>]*>)`), `$1<f t="array" ref="${ref}">${f.slice(1)}</f>`), xml),
    );
  }
  const array = (s: WorkbookSession, ref: string) => {
    const u = s.stateUpdates().find((x) => x.sheet === 0 && x.row === at(ref).row && x.col === at(ref).col);
    return u?.array && rangeToA1(u.array);
  };

  it('fills the range it was entered in, repeating a single row or column and showing #N/A past its result, and follows its inputs', async () => {
    const s = fresh(await arrays());
    expect([s.calcReport?.unsupported, s.calcReport?.mismatched]).toEqual([0, 0]);
    expect(column(s, ['C1', 'C2', 'C3'])).toEqual([2, 4, 6]);
    expect(column(s, ['D1', 'E1', 'D3', 'E3', 'D4', 'E4'])).toEqual([10, 10, 30, 30, { error: '#N/A' }, { error: '#N/A' }]);
    expect(column(s, ['F1', 'G1', 'H1'])).toEqual([1, 2, { error: '#N/A' }]);
    expect([array(s, 'C1'), array(s, 'E4'), shown(s, 'E4')?.spill]).toEqual(['C1:C3', 'D1:E4', undefined]);
    type(s, 'A2', '5');
    expect(column(s, ['C2', 'D2', 'E2', 'G1'])).toEqual([10, 50, 50, 5]);
  });

  it('is changed only as a whole: its cells refuse typing and rows inserted through it, and copy as their values', async () => {
    const s = fresh(await arrays());
    expect(english.text(type(s, 'C2', '9').skippedReasons[0])).toMatch(/array formula/);
    let refused: unknown;
    try {
      s.restructure({ type: 'insertRows', sheet: 0, at: 1, count: 1 });
    } catch (e) {
      refused = e;
    }
    expect(english.text(errorText(refused))).toMatch(/array formula in C1:C3/);
    const r = s.apply({ type: 'copy', srcSheet: 0, src: rect('C1:C2'), sheet: 0, dst: rect('J1:J2') });
    expect(english.text(r.warnings[0])).toMatch(/^1 formula .* pasted as values/);
    expect(s.current(0, 0, 9).content).toEqual({ kind: 'number', value: 2 });
    expect(s.current(0, 1, 9).content).toEqual({ kind: 'number', value: 4 });
  });

  it('saves its new values in its cells, its formula and range as they were', async () => {
    const s = fresh(await arrays());
    type(s, 'A2', '5');
    const xml = parts(s.save())['xl/worksheets/sheet1.xml'];
    expect(cellXml(xml, 'C1')).toBe('<c r="C1"><f t="array" ref="C1:C3">A1:A3*2</f><v>2</v></c>');
    expect(cellXml(xml, 'C2')).toBe('<c r="C2"><v>10</v></c>');
    expect(cellXml(xml, 'E4')).toBe('<c r="E4" t="e"><v>#N/A</v></c>');
    const again = fresh(s.save());
    expect([again.calcReport?.mismatched, value(again, 'E2')]).toEqual([0, 50]);
  });

  it("keeps Excel's values, out of date once its inputs change, when its formula isn't calculated here", async () => {
    // An add-in's function, which neither engine has.
    const s = fresh(await arrays('=_xludf.TIMESTEN(A1:A3)'));
    expect(s.calcReport?.unsupported).toBe(1);
    const r = type(s, 'A2', '5');
    expect([s.current(0, 1, 3).content, s.current(0, 1, 4).content]).toEqual([
      { kind: 'number', value: 20 },
      { kind: 'number', value: 20 },
    ]);
    expect(sent(r, 'E2')!.flags & CellFlag.Stale).toBe(CellFlag.Stale);
    expect(sent(r, 'E2')!.array).toBeUndefined();
  });
});

describe('edits around a spill', () => {
  it('refuse to sort part of it', () => {
    const s = fresh();
    type(s, 'B1', '=SEQUENCE(3)');
    const r = s.apply({ type: 'sort', sheet: 0, rect: rect('A1:B6'), keys: [{ col: 0, descending: false }], header: false });
    expect(english.text(r.rejected)).toBe('Sorting would move part of the results that spill into B1:B3. Sort around them, or sort with the SORT function.');
    expect(s.apply({ type: 'sort', sheet: 0, rect: rect('A2:B6'), keys: [{ col: 0, descending: false }], header: true }).rejected).toBeDefined();
    // Its header row, which stays where it is, may be one of them.
    expect(s.apply({ type: 'sort', sheet: 0, rect: rect('A3:B6'), keys: [{ col: 0, descending: false }], header: true }).rejected).toBeUndefined();
  });

  it('copy spilled cells as their values, and a whole spill as its formula, which spills again', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(3)*10');
    s.apply({ type: 'copy', srcSheet: 0, src: rect('C2:C3'), sheet: 0, dst: rect('E1:E2') });
    expect(s.current(0, 0, 4).content).toEqual({ kind: 'number', value: 20 });
    s.apply({ type: 'copy', srcSheet: 0, src: rect('C1:C3'), sheet: 0, dst: rect('G1:G3') });
    expect(s.current(0, 0, 6).content).toEqual({ kind: 'formula', formula: '_xlfn.SEQUENCE(3)*10' });
    expect(shown(s, 'G3')).toMatchObject({ value: 30, spilled: true, spill: 'G1:G3' });
    s.apply({ type: 'copy', srcSheet: 0, src: rect('C2:C3'), sheet: 0, dst: rect('I1:I2'), what: 'values' });
    expect(s.current(0, 1, 8).content).toEqual({ kind: 'number', value: 30 });
  });

  it('fill from spilled cells with their values', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(2)');
    s.apply({ type: 'fill', sheet: 0, src: rect('C2'), dst: rect('C2:E2') });
    expect([s.current(0, 1, 3).content, s.current(0, 1, 4).content]).toEqual([
      { kind: 'number', value: 2 },
      { kind: 'number', value: 2 },
    ]);
  });

  it('keep it when only its cells’ formats change', () => {
    const s = fresh();
    type(s, 'C1', '=SEQUENCE(3)');
    s.apply({ type: 'format', sheet: 0, rects: [rect('C1:C3')], change: { bold: true } });
    expect(column(s, ['C1', 'C2', 'C3'])).toEqual([1, 2, 3]);
    const xml = parts(s.save())['xl/worksheets/sheet1.xml'];
    expect(cellXml(xml, 'C2')).toMatch(/^<c r="C2" s="\d+"><v>2<\/v><\/c>$/);
  });

  it('let rows be inserted and deleted through it, after which it spills again; deleting its formula takes its values', () => {
    const s = fresh();
    type(s, 'E2', '=SEQUENCE(3)');
    const inserted = s.restructure({ type: 'insertRows', sheet: 0, at: 2, count: 1 });
    inserted.ensureEngine();
    expect(column(inserted, ['E2', 'E3', 'E4', 'E5'])).toEqual([1, 2, 3, null]);
    const through = s.restructure({ type: 'deleteRows', sheet: 0, at: 2, count: 1 });
    through.ensureEngine();
    expect(column(through, ['E2', 'E3', 'E4'])).toEqual([1, 2, 3]);
    const deleted = s.restructure({ type: 'deleteRows', sheet: 0, at: 1, count: 1 });
    deleted.ensureEngine();
    expect(deleted.current(0, 1, 4).content).toEqual({ kind: 'empty' });
    expect([value(deleted, 'E2'), value(deleted, 'E3')]).toEqual([null, null]);
    expect(parts(deleted.save())['xl/worksheets/sheet1.xml']).not.toMatch(/<c r="E\d+"/);
  });

  // E4's LAMBDA is the one formula there HyperFormula can't calculate; the native engine calculates it.
  it.skipIf(process.env.ENGINE !== 'hyperformula')('mark its values out of date with the formula, when it reads a value that is', () => {
    const s = corpus();
    const arrays = sheetIndex(s, 'Arrays');
    // E4 holds a LAMBDA formula, which isn't calculated here; what it reads can't be told, so once anything changes it may be out of date.
    type(s, 'K1', '=SEQUENCE(2)*E4', arrays);
    expect(shown(s, 'K1', arrays)!.flags & CellFlag.Stale).toBe(CellFlag.Stale);
    expect(shown(s, 'K2', arrays)!.flags).toBe(CellFlag.Spilled | CellFlag.Stale);
    type(s, 'K5', '=SEQUENCE(2)*D9', arrays);
    expect(shown(s, 'K6', arrays)?.flags).toBe(CellFlag.Spilled);
  });

  it('list a typed value that blocks it with the value it covered', () => {
    const s = corpus();
    const arrays = sheetIndex(s, 'Arrays');
    type(s, 'A3', 'Mixed', arrays);
    expect(s.changes().find((c) => c.ref === 'A3')).toMatchObject({ before: 'Hardware', after: 'Mixed' });
  });
});

