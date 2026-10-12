import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { compareForSort, contentText, WorkbookSession } from '../src/core/session';
import { CellFlag, type EditResult } from '../src/core/types';
import { english } from '../src/i18n';
import { sample, sheetIndex } from './helpers';

/**
 * S: A1:A4 = 3, 1, 2, 1 and B1:B4 = x, b, y, a; D1:E2 merged, holding "M";
 * F1 a date, F2 a date and time; G1 has a whole-number rule (1 to 10).
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  [3, 1, 2, 1].forEach((v, i) => (ws.getCell(`A${i + 1}`).value = v));
  ['x', 'b', 'y', 'a'].forEach((v, i) => (ws.getCell(`B${i + 1}`).value = v));
  ws.getCell('D1').value = 'M';
  ws.mergeCells('D1:E2');
  ws.getCell('F1').value = new Date('2026-01-05T00:00:00Z');
  ws.getCell('F1').numFmt = 'yyyy-mm-dd';
  ws.getCell('F2').value = new Date('2026-01-05T06:30:00Z');
  ws.getCell('F2').numFmt = 'yyyy-mm-dd hh:mm';
  (ws as unknown as { dataValidations: { add(ref: string, v: object): void } }).dataValidations.add('G1', { type: 'whole', operator: 'between', formulae: [1, 10] });
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

const open = (options = {}) => WorkbookSession.open(bytes, 's.xlsx', 'en-US', options);
const value = (s: WorkbookSession, row: number, col: number) => {
  const c = s.current(0, row, col).content;
  return c.kind === 'empty' ? null : c.kind === 'formula' ? `=${c.formula}` : c.value;
};
const said = (r: EditResult) => ({ warnings: r.warnings.map((w) => english.text(w)), skipped: r.skippedReasons.map((w) => english.text(w)), rejected: r.rejected && english.text(r.rejected) });

describe('cells that cannot be edited', () => {
  it('ignore typing into the covered part of a merged cell, silently', () => {
    const s = open();
    const r = s.apply({ type: 'pasteText', sheet: 0, row: 0, col: 3, values: [['top', 'covered'], ['covered', 'covered']] });
    expect([value(s, 0, 3), value(s, 0, 4), value(s, 1, 3)]).toEqual(['top', null, null]);
    expect(r.skipped).toBe(0);
    expect(said(r).skipped).toEqual([]);
  });

  it('turn down edits to a chart sheet, saying why', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const r = s.apply({ type: 'input', sheet: sheetIndex(s, 'Monthly Chart'), row: 0, col: 0, text: 'x' });
    expect([r.skipped, said(r).skipped]).toEqual([1, ['This sheet cannot be edited here.']]);
  });

  it('copy array formula cells as their values, and say so', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const lookups = sheetIndex(s, 'Lookups');
    // B15:D15 holds an array formula entered with Ctrl+Shift+Enter, which changes only as a whole.
    const r = s.apply({ type: 'copy', srcSheet: lookups, src: { r1: 14, c1: 1, r2: 14, c2: 1 }, sheet: lookups, dst: { r1: 29, c1: 7, r2: 29, c2: 7 } });
    expect(said(r).warnings).toEqual(["1 formula couldn't be adjusted for the new position and was pasted as values."]);
    expect(s.current(lookups, 29, 7).content).toEqual({ kind: 'number', value: 25 });
  });
});

describe('typed and pasted text', () => {
  it('keeps a pasted formula that is not valid as text, and drops what falls off the sheet', () => {
    const s = open();
    const r = s.apply({ type: 'pasteText', sheet: 0, row: 1_048_575, col: 8, values: [['=SUM(', 'ok'], ['gone', 'gone']] });
    expect(said(r).warnings).toEqual(["1 pasted value looked like a formula but wasn't valid, so it was pasted as text."]);
    expect(s.current(0, 1_048_575, 8).content).toEqual({ kind: 'string', value: '=SUM(' });
    expect(r.redo.length).toBe(2);
  });

  it('writes values set by schema logic as typed, keeping invalid formulas as text with a warning', () => {
    const s = open();
    const r = s.apply({ type: 'inputs', sheet: 0, cells: [{ row: 9, col: 0, text: '=SUM(' }, { row: 9, col: 1, text: '12%' }] });
    expect(said(r).warnings[0]).toMatch(/^There's a problem with this formula: /);
    expect(value(s, 9, 0)).toBe('=SUM(');
    expect(value(s, 9, 1)).toBe(0.12);
    expect(s.model.styles.get(s.current(0, 9, 1).style).numFmt).toBe('0%');
  });

  it('checks typed values against a validation rule, refusing bad formulas and passing cells without one', () => {
    const s = open();
    expect(s.checkValidation(0, 0, 6, '5', 0)).toBe(true);
    expect(s.checkValidation(0, 0, 6, '11', 0)).toBe(false);
    expect(s.checkValidation(0, 0, 6, '=SUM(', 0)).toBe(false);
    expect(s.checkValidation(0, 0, 6, '11', 5)).toBe(true);
  });

  it('clearing empty cells changes nothing', () => {
    const s = open();
    const r = s.apply({ type: 'clear', sheet: 0, rects: [{ r1: 20, c1: 0, r2: 30, c2: 5 }] });
    expect(r.updates).toEqual([]);
    expect(s.editedCount).toBe(0);
  });
});

describe('fill', () => {
  it('extends a series upwards and leftwards as well as down and right', () => {
    const s = open();
    s.apply({ type: 'pasteText', sheet: 0, row: 10, col: 0, values: [['1'], ['2']] });
    s.apply({ type: 'fill', sheet: 0, src: { r1: 10, c1: 0, r2: 11, c2: 0 }, dst: { r1: 8, c1: 0, r2: 11, c2: 0 } });
    expect([value(s, 8, 0), value(s, 9, 0)]).toEqual([-1, 0]);
    s.apply({ type: 'pasteText', sheet: 0, row: 20, col: 2, values: [['Mon', 'Tue']] });
    s.apply({ type: 'fill', sheet: 0, src: { r1: 20, c1: 2, r2: 20, c2: 3 }, dst: { r1: 20, c1: 0, r2: 20, c2: 3 } });
    expect([value(s, 20, 0), value(s, 20, 1)]).toEqual(['Sat', 'Sun']);
    s.apply({ type: 'input', sheet: 0, row: 25, col: 0, text: '=A1*2' });
    s.apply({ type: 'fill', sheet: 0, src: { r1: 25, c1: 0, r2: 25, c2: 0 }, dst: { r1: 25, c1: 0, r2: 25, c2: 2 } });
    expect([value(s, 25, 1), value(s, 25, 2)]).toEqual(['=B1*2', '=C1*2']);
  });
});

describe('sorting', () => {
  it('sorts by several keys, keeping tied rows in their order', () => {
    const s = open();
    s.apply({ type: 'sort', sheet: 0, rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, keys: [{ col: 0, descending: false }, { col: 1, descending: true }], header: false });
    expect([0, 1, 2, 3].map((r) => [value(s, r, 0), value(s, r, 1)])).toEqual([
      [1, 'b'],
      [1, 'a'],
      [2, 'y'],
      [3, 'x'],
    ]);
  });

  it('refuses to sort through a merged cell', () => {
    const r = open().apply({ type: 'sort', sheet: 0, rect: { r1: 0, c1: 0, r2: 3, c2: 4 }, keys: [{ col: 0, descending: false }], header: false });
    expect(said(r).rejected).toBe('Sorting would move part of the merged cell D1:E2. Unmerge it, or sort around it.');
  });

  it("orders values Excel's way: numbers, text, logicals, errors, with blanks last either way", () => {
    const values = [true, null, 'b', { error: '#N/A' }, 2, false, 'A', '', 1];
    expect([...values].sort((a, b) => compareForSort(a, b, false))).toEqual([1, 2, 'A', 'b', false, true, { error: '#N/A' }, null, '']);
    expect([...values].sort((a, b) => compareForSort(a, b, true))).toEqual([{ error: '#N/A' }, true, false, 'b', 'A', 2, 1, null, '']);
  });
});

describe('recalculation', () => {
  it('works out formulas typed while it was off once it is turned on', () => {
    const s = open({ recalc: 'off' });
    s.apply({ type: 'input', sheet: 0, row: 0, col: 9, text: '=A1*2' });
    s.apply({ type: 'input', sheet: 0, row: 1, col: 9, text: '=_xludf.MYFN(1)' });
    const report = s.enableRecalc();
    expect(report.updates.find((u) => u.row === 0 && u.col === 9)?.n).toBe(6);
    // A function the engine doesn't have stays marked as not calculated here.
    expect(report.updates.find((u) => u.row === 1 && u.col === 9)!.flags & CellFlag.Unsupported).toBe(CellFlag.Unsupported);
    expect(s.enableRecalc()).toBe(report);
  });

  it("follows a sheet whose name has an apostrophe, which formulas write doubled ('Bob''s'!A1)", async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet("Bob's").getCell('A1').value = 5;
    const main = wb.addWorksheet('Main');
    main.getCell('A1').value = { formula: "'Bob''s'!A1*2", result: 10 };
    main.getCell('A2').value = { formula: "SUM('Bob''s'!A1:A2)", result: 5 };
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 'bob.xlsx');
    const r = s.apply({ type: 'input', sheet: 0, row: 0, col: 0, text: '7' });
    const at = (row: number) => r.updates.find((u) => u.sheet === 1 && u.row === row && u.col === 0)!;
    expect([at(0).n, at(1).n]).toEqual([14, 7]);
    expect(at(0).flags & (CellFlag.Stale | CellFlag.Mismatch)).toBe(0);
  });
});

describe('the change list', () => {
  it('drops a cell set back to what the file held across a structural change, and keeps its first value otherwise', () => {
    const s = open();
    s.apply({ type: 'input', sheet: 0, row: 0, col: 0, text: '30' });
    const moved = s.restructure({ type: 'insertRows', sheet: 0, at: 0, count: 1 });
    moved.apply({ type: 'input', sheet: 0, row: 1, col: 1, text: 'changed' });
    moved.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: '31' });
    expect(moved.changes().filter((c) => !c.note)).toEqual([
      { sheet: 'S', ref: 'A2', before: '3', after: '31' },
      { sheet: 'S', ref: 'B2', before: 'x', after: 'changed' },
    ]);
    moved.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: '3' });
    expect(moved.changes().filter((c) => !c.note).map((c) => c.ref)).toEqual(['B2']);
  });
});

describe('cell text', () => {
  it('shows dates by their format, from either date system, and logicals in capitals', () => {
    expect(contentText({ kind: 'number', value: 46027 }, 'yyyy-mm-dd')).toBe('2026-01-05');
    expect(contentText({ kind: 'number', value: 46027.25 }, 'dd/mm/yyyy hh:mm')).toBe('2026-01-05 06:00');
    expect(contentText({ kind: 'number', value: 44565 }, 'yyyy-mm-dd', true)).toBe('2026-01-05');
    expect(contentText({ kind: 'number', value: 0.1 + 0.2 })).toBe('0.3');
    expect(contentText({ kind: 'bool', value: false })).toBe('FALSE');
    expect(contentText({ kind: 'error', value: '#N/A' })).toBe('#N/A');
  });

  it('reads dates from the file as their format shows them', () => {
    const s = open();
    const text = (row: number) => contentText(s.current(0, row, 5).content, s.model.styles.get(s.current(0, row, 5).style).numFmt);
    expect([text(0), text(1)]).toEqual(['2026-01-05', '2026-01-05 06:30']);
  });
});
