import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { EditOp, FormatChange, Rect } from '../src/core/types';
import { english } from '../src/i18n';
import { parts } from './helpers';

/** Data!A1:C4 with numbers, B1 bold in Arial and A2:B2 merged; Locked is protected. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  for (let r = 1; r <= 4; r++) data.addRow([r, r * 2.5, r * 1000]);
  data.getCell('B1').font = { bold: true, name: 'Arial', size: 12 };
  data.mergeCells('A2:B2');
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const rect = (r1: number, c1: number, r2 = r1, c2 = c1): Rect => ({ r1, c1, r2, c2 });
const open = (b = bytes) => WorkbookSession.open(b, 'format.xlsx', 'en-US');
const format = (s: WorkbookSession, rects: Rect[], change: FormatChange, extra: Partial<Extract<EditOp, { type: 'format' }>> = {}) =>
  s.apply({ type: 'format', sheet: 0, rects, change, ...extra });
const style = (s: WorkbookSession, row: number, col: number, sheet = 0) => s.model.styles.get(s.current(sheet, row, col).style);

describe('formatting cells', () => {
  it('changes fonts, fills, alignment and number formats, keeping the rest of each format', () => {
    const s = open();
    format(s, [rect(0, 0, 0, 1)], { italic: true, fill: '#FFEE00', hAlign: 'center', wrap: true });
    format(s, [rect(0, 2)], { numFmt: '#,##0.00', color: '#C00000', fontSizePt: 14 });
    const t = open(s.save());
    // B1 keeps its bold and its font.
    expect(style(t, 0, 1)).toMatchObject({ bold: true, italic: true, fontFamily: 'Arial', fill: '#FFEE00', hAlign: 'center', wrap: true });
    expect(style(t, 0, 0)).toMatchObject({ bold: false, italic: true, fill: '#FFEE00' });
    expect(style(t, 0, 2)).toMatchObject({ numFmt: '#,##0.00', color: '#C00000', fontSizePx: 18.7 });
    expect(style(t, 1, 2)).toMatchObject({ numFmt: 'General' });
  });

  it("writes Excel's line styles as themselves", () => {
    const s = open();
    const lines = [
      { width: 2, pattern: 'dashDot', color: '#FF0000' },
      { width: 1, pattern: 'hair' },
      { width: 2, pattern: 'slantDashDot' },
      { width: 1, pattern: 'dashDotDot' },
    ] as const;
    lines.forEach((edge, i) => format(s, [rect(5, i)], {}, { borders: { where: 'all', edge } }));
    const t = open(s.save());
    // Excel's automatic colour reads back as the black it stands for.
    expect(lines.map((_e, i) => style(t, 5, i).border?.left)).toEqual(lines.map((e) => ({ color: '#000000', ...e })));
  });

  it('sets vertical alignment and indent, bottom and no indent being the defaults it leaves out', () => {
    const s = open();
    format(s, [rect(0, 0)], { vAlign: 'top', indent: 2, hAlign: 'left' });
    format(s, [rect(0, 1)], { vAlign: 'center' });
    const t = open(s.save());
    expect(style(t, 0, 0)).toMatchObject({ vAlign: 'top', indent: 2, hAlign: 'left' });
    expect(style(t, 0, 1)).toMatchObject({ vAlign: 'center', indent: 0 });
    format(t, [rect(0, 0)], { vAlign: 'bottom', indent: 0 });
    expect(style(open(t.save()), 0, 0)).toMatchObject({ vAlign: undefined, indent: 0 });
  });

  it('formats empty cells too, and every cell of a merge, so its borders show', () => {
    const s = open();
    format(s, [rect(1, 0, 1, 1), rect(6, 0)], { fill: '#00B050' }, { borders: { where: 'outside' } });
    const t = open(s.save());
    expect(style(t, 1, 1).border).toMatchObject({ top: { width: 1 }, bottom: { width: 1 }, right: { width: 1 } });
    expect(style(t, 1, 1).border?.left).toBeUndefined();
    expect(style(t, 6, 0)).toMatchObject({ fill: '#00B050', border: { left: {}, right: {}, top: {}, bottom: {} } });
  });

  it('puts borders where asked: all, inside, one side, or none', () => {
    const s = open();
    const thick = { width: 3 as const, pattern: 'solid' as const, color: '#0070C0' };
    format(s, [rect(0, 0, 2, 2)], {}, { borders: { where: 'inside' } });
    expect(style(s, 0, 0).border).toEqual({ right: { width: 1, pattern: 'solid' }, bottom: { width: 1, pattern: 'solid' } });
    expect(style(s, 1, 1).border).toMatchObject({ left: {}, right: {}, top: {}, bottom: {} });
    format(s, [rect(0, 0, 2, 2)], {}, { borders: { where: 'bottom', edge: thick } });
    expect(style(s, 2, 1).border?.bottom).toEqual(thick);
    expect(style(s, 1, 1).border?.bottom).toEqual({ width: 1, pattern: 'solid' });
    format(s, [rect(0, 0, 2, 2)], {}, { borders: { where: 'none' } });
    expect(style(s, 1, 1).border).toBeUndefined();
  });

  it('clears formats back to the default, keeping the values, and undoes', () => {
    const s = open();
    const r = format(s, [rect(0, 0, 0, 1)], {}, { clear: true });
    expect(style(s, 0, 1)).toMatchObject({ bold: false, fontFamily: 'Calibri' });
    expect(s.current(0, 0, 1).content).toEqual({ kind: 'number', value: 2.5 });
    s.apply({ type: 'restore', cells: r.undo });
    expect(style(s, 0, 1)).toMatchObject({ bold: true, fontFamily: 'Arial' });
  });

  it('changes a font family without the theme overriding it, and reuses formats already made', () => {
    const s = open();
    const first = format(s, [rect(3, 0)], { fontFamily: 'Georgia', bold: true });
    const again = format(s, [rect(3, 1)], { fontFamily: 'Georgia', bold: true });
    expect(first.newStyles).toHaveLength(1);
    expect(again.newStyles).toHaveLength(0);
    const xml = parts(s.save())['xl/styles.xml'];
    const georgia = /<font><b\/>(?:(?!<\/font>).)*<name val="Georgia"\/><\/font>/.exec(xml)?.[0];
    expect(georgia).toBeDefined();
    expect(georgia).not.toMatch(/scheme/);
    // Counts follow what was added.
    const fonts = /<fonts\b[^>]*count="(\d+)"/.exec(xml)![1];
    expect(Number(fonts)).toBe((xml.match(/<font>/g) ?? []).length);
  });

  it('is refused on a protected sheet, and on too many cells at once', () => {
    const s = open();
    const r = s.apply({ type: 'format', sheet: 1, rects: [rect(0, 0)], change: { bold: true } });
    expect(r.undo).toEqual([]);
    expect(r.skipped).toBe(1);
    const huge = format(s, [rect(0, 0, 2_000, 999)], { bold: true });
    expect(english.text(huge.rejected!)).toBe('Formats can be changed on up to 1,000,000 cells at once.');
    // A whole column goes as far as the sheet's cells.
    const column = format(s, [rect(0, 2, 1_048_575, 2)], { bold: true });
    expect(column.redo.map((c) => c.row)).toEqual([0, 1, 2, 3]);
  });
});

describe('formats made one after another', () => {
  it('keep what each earlier change made', () => {
    const s = open();
    const r = rect(3, 0);
    format(s, [r], { bold: true });
    format(s, [r], { fill: '#FFFF00' }, { borders: { where: 'all' } });
    format(s, [r], { fontFamily: 'Georgia' });
    format(s, [r], { italic: true }, { borders: { where: 'bottom', edge: { width: 2, pattern: 'solid' } } });
    const t = open(s.save());
    expect(style(t, 3, 0)).toMatchObject({ bold: true, italic: true, fontFamily: 'Georgia', fill: '#FFFF00', border: { left: {}, top: {}, bottom: { width: 2 } } });
  });
});
