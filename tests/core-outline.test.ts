import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { EditResult, LayoutOp } from '../src/core/types';
import { english } from '../src/i18n';
import { parts } from './helpers';

/** tests/corpus/outline.xlsx, made by scripts/gen_outline.py with XlsxWriter. */
const bytes = new Uint8Array(readFileSync(join(__dirname, 'corpus', 'outline.xlsx')));
const REPORT = 0;
const ABOVE = 1;
const LOCKED = 2;

const open = (b: Uint8Array = bytes) => WorkbookSession.open(b, 'outline.xlsx', 'en-US');
const reopen = (s: WorkbookSession) => open(s.save());
const info = (s: WorkbookSession, sheet: number) => s.loadResult().sheets[sheet];
const layout = (s: WorkbookSession, sheet: number, op: LayoutOp): EditResult => s.apply({ type: 'layout', sheet, op });
const outline = (s: WorkbookSession, sheet: number, axis: 'rows' | 'cols', levels: Record<number, number>, collapsed: Record<number, boolean> = {}) =>
  layout(s, sheet, { kind: 'outline', axis, levels, collapsed });
const sheetXml = (s: WorkbookSession, n: number) => parts(s.save())[`xl/worksheets/sheet${n}.xml`];
const notes = (s: WorkbookSession) => s.changes().map((c) => `${c.ref} ${english.text(c.note!)}`);

/** The workbook with one sheet's XML changed. */
function edited(n: number, change: (xml: string) => string): Uint8Array {
  const files = unzipSync(bytes);
  const path = `xl/worksheets/sheet${n}.xml`;
  files[path] = strToU8(change(strFromU8(files[path])));
  return zipSync(files);
}

describe('reading outlines', () => {
  it('reads how deep each row and column is grouped, the collapsed groups, and where their summaries are', () => {
    const s = open();
    const report = info(s, REPORT);
    expect(report.outline).toEqual({
      rows: { levels: { 1: 2, 2: 2, 3: 1, 4: 2, 5: 2, 6: 1 }, collapsed: [6], summaryAfter: true },
      cols: { levels: { 1: 1, 2: 1, 3: 1, 5: 1, 6: 1, 7: 1 }, collapsed: [8], summaryAfter: true },
    });
    // The rows and columns a collapsed group hides are the ones marked hidden.
    expect([report.hiddenRows, report.hiddenCols]).toEqual([[4, 5], [5, 6, 7]]);
    const above = info(s, ABOVE).outline!;
    expect([above.rows.summaryAfter, above.cols.summaryAfter]).toEqual([false, false]);
    expect([above.rows.levels, above.cols.levels]).toEqual([{ 2: 1, 3: 1, 4: 1 }, { 1: 1, 2: 1 }]);
    expect(info(s, LOCKED).formatAllowed).toEqual({ rows: true, cols: false });
  });

  it('has no outline on sheets without groups', () => {
    const s = open(edited(3, (xml) => xml.replace(/ outlineLevel="1"/g, '')));
    expect(info(s, LOCKED).outline).toBeUndefined();
    expect(info(s, REPORT).outline).toBeDefined();
  });
});

describe('outline changes', () => {
  it('are saved: rows and columns grouped and ungrouped, groups collapsed and expanded, with the deepest level in <sheetFormatPr>', () => {
    const s = open();
    // Rows 10-11 grouped three deep, the first quarter's columns ungrouped, and South expanded.
    outline(s, REPORT, 'rows', { 9: 3, 10: 3 });
    outline(s, REPORT, 'cols', { 1: 0, 2: 0, 3: 0, 12: 1 });
    outline(s, REPORT, 'rows', {}, { 6: false });
    layout(s, REPORT, { kind: 'hidden', axis: 'rows', lines: { 4: false, 5: false } });
    const xml = sheetXml(s, 1);
    expect(xml).toMatch(/<sheetFormatPr defaultRowHeight="15" outlineLevelRow="3" outlineLevelCol="1"\/>/);
    expect(xml).toMatch(/<row r="10" outlineLevel="3"\/><row r="11" outlineLevel="3"\/>/);
    expect(xml).toMatch(/<row r="7" spans="1:10" outlineLevel="1">/);
    // B:D, no longer grouped, are now written with E, which is like them.
    expect(xml).toMatch(/<col min="2" max="5" width="8.7109375" customWidth="1"\/>/);
    expect(xml).toMatch(/<col min="6" max="8" width="8.7109375" hidden="1" customWidth="1" outlineLevel="1"\/>/);
    // A column that had no <col> of its own gets the sheet's width with its level, as Excel writes it.
    expect(xml).toMatch(/<col min="13" max="13" outlineLevel="1" width="9.140625"\/>/);
    const t = info(reopen(s), REPORT);
    expect(t.outline!.rows).toEqual({ levels: { 1: 2, 2: 2, 3: 1, 4: 2, 5: 2, 6: 1, 9: 3, 10: 3 }, collapsed: [], summaryAfter: true });
    expect(t.outline!.cols.levels).toEqual({ 5: 1, 6: 1, 7: 1, 12: 1 });
    expect(t.hiddenRows).toEqual([]);
  });

  it('take the deepest levels out of <sheetFormatPr> once the groups are gone, and add it to a sheet without one', () => {
    const s = open();
    outline(s, REPORT, 'rows', { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }, { 6: false });
    expect(sheetXml(s, 1)).toMatch(/<sheetFormatPr defaultRowHeight="15" outlineLevelCol="1"\/>/);
    const bare = open(edited(2, (xml) => xml.replace(/<sheetFormatPr[^>]*\/>/, '').replace(/ outlineLevel="1"/g, '')));
    expect(info(bare, ABOVE).outline!.rows.levels).toEqual({});
    outline(bare, ABOVE, 'rows', { 5: 1, 6: 1 });
    const xml = sheetXml(bare, 2);
    expect(xml).toMatch(/<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"\/><cols>/);
    expect(info(reopen(bare), ABOVE).outline!.rows).toEqual({ levels: { 5: 1, 6: 1 }, collapsed: [], summaryAfter: false });
  });

  it('leave the file as it was when there are none', () => {
    const s = open();
    expect(s.save()).toBe(bytes);
    // A change undone leaves nothing to save either.
    const r = outline(s, REPORT, 'rows', { 9: 1 });
    s.apply({ type: 'restore', cells: [], layout: r.layout!.undo });
    expect(s.editedCount).toBe(0);
    expect(s.save()).toBe(bytes);
  });

  it('undo with the change each one returns, and list what differs from the file', () => {
    const s = open();
    const done = [
      outline(s, REPORT, 'rows', { 9: 1, 10: 1, 3: 2 }),
      outline(s, REPORT, 'cols', { 1: 0, 2: 0, 3: 0 }, { 4: true }),
      outline(s, REPORT, 'rows', {}, { 6: false }),
    ];
    expect(done[0].layout!.undo).toEqual([{ sheet: REPORT, op: { kind: 'outline', axis: 'rows', levels: { 3: 1, 9: 0, 10: 0 }, collapsed: {} } }]);
    expect(notes(s)).toEqual(['4:4, 10:11 Grouped 3 rows', '7:7 Expanded 1 group', 'B:D Ungrouped 3 columns', 'E:E Collapsed 1 group']);
    expect(s.editedCount).toBe(4);
    for (const r of done.reverse()) s.apply({ type: 'restore', cells: [], layout: r.layout!.undo });
    expect(s.changes()).toEqual([]);
    expect(info(s, REPORT).outline!.rows.collapsed).toEqual([6]);
  });

  it('are refused past level 7, and on protected sheets unless their protection lets rows or columns be formatted', () => {
    const s = open();
    expect(english.text(outline(s, REPORT, 'rows', { 9: 8 }).rejected!)).toBe('Groups can go at most 7 levels deep, as in Excel.');
    // "Locked" lets rows be formatted, so they can be grouped, hidden and shown, but its columns can't.
    expect(outline(s, LOCKED, 'rows', { 5: 1 }).rejected).toBeUndefined();
    expect(layout(s, LOCKED, { kind: 'hidden', axis: 'rows', lines: { 1: true } }).rejected).toBeUndefined();
    expect(english.text(outline(s, LOCKED, 'cols', { 1: 1 }).rejected!)).toBe('This sheet is protected in Excel, so it is read-only here.');
    expect(english.text(layout(s, LOCKED, { kind: 'hidden', axis: 'cols', lines: { 1: true } }).rejected!)).toBe('This sheet is protected in Excel, so it is read-only here.');
    expect(layout(s, LOCKED, { kind: 'rowHeights', heights: { 1: 30 } }).rejected).toBeDefined();
    const t = info(reopen(s), LOCKED);
    expect([t.outline!.rows.levels, t.hiddenRows]).toEqual([{ 1: 1, 2: 1, 3: 1, 5: 1 }, [1]]);
  });
});

describe('outlines and inserted or deleted rows and columns', () => {
  it('keep each row and column at its level, and rows and columns inserted inside a group join it', () => {
    const s = open();
    // Between two rows of North's lines, and above them (outside any group).
    let next = s.restructure({ type: 'insertRows', sheet: REPORT, at: 2, count: 2 });
    next = next.restructure({ type: 'insertRows', sheet: REPORT, at: 1, count: 1 });
    let o = info(next, REPORT).outline!;
    expect(o.rows.levels).toEqual({ 2: 2, 3: 2, 4: 2, 5: 2, 6: 1, 7: 2, 8: 2, 9: 1 });
    expect(o.rows.collapsed).toEqual([9]);
    next = next.restructure({ type: 'deleteRows', sheet: REPORT, at: 6, count: 2 });
    expect(info(next, REPORT).outline!.rows.levels).toEqual({ 2: 2, 3: 2, 4: 2, 5: 2, 6: 2, 7: 1 });
    // Inside the first quarter's columns, and between May and June once each has its own <col>.
    next = next.restructure({ type: 'insertCols', sheet: REPORT, at: 2, count: 1 });
    next.apply({ type: 'layout', sheet: REPORT, op: { kind: 'colWidths', widths: { 7: 100 } } });
    next = next.restructure({ type: 'insertCols', sheet: REPORT, at: 7, count: 1 });
    o = info(next, REPORT).outline!;
    expect(o.cols.levels).toEqual({ 1: 1, 2: 1, 3: 1, 4: 1, 6: 1, 7: 1, 8: 1, 9: 1 });
    expect(o.cols.collapsed).toEqual([10]);
    // Right of a group's last column is outside it.
    next = next.restructure({ type: 'insertCols', sheet: REPORT, at: 5, count: 1 });
    expect(info(next, REPORT).outline!.cols.levels[5]).toBeUndefined();
  });

  it('lower the deepest level <sheetFormatPr> gives once the deepest rows or columns are deleted', () => {
    const s = open();
    let next = s.restructure({ type: 'deleteRows', sheet: REPORT, at: 1, count: 5 });
    next = next.restructure({ type: 'deleteCols', sheet: REPORT, at: 1, count: 7 });
    expect(info(next, REPORT).outline!.rows.levels).toEqual({ 1: 1 });
    expect(info(next, REPORT).outline!.cols.levels).toEqual({});
    expect(sheetXml(next, 1)).toMatch(/<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="0"\/>/);
  });

  it('carry outline changes made before the rows move, and keep them in the session gone back to', () => {
    const s = open();
    outline(s, REPORT, 'rows', { 9: 1, 10: 1 });
    const next = s.restructure({ type: 'insertRows', sheet: REPORT, at: 0, count: 1 });
    expect(info(next, REPORT).outline!.rows.levels).toMatchObject({ 10: 1, 11: 1 });
    expect(notes(next)).toContain('10:11 Grouped 2 rows');
    // The session the insert was made from, which undo goes back to, still has the group where it was.
    expect(info(s, REPORT).outline!.rows.levels).toMatchObject({ 9: 1, 10: 1 });
  });
});
