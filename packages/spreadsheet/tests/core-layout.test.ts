import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { EditResult, FilterCondition, LayoutOp } from '../src/core/types';
import { english } from '../src/i18n';
import { strToU8, zipSync } from 'fflate';
import { parts } from './helpers';

/**
 * Sheet "Plain": A1:C3 filled, nothing else. Sheet "Laid": column B 20 wide,
 * row 2 30pt high, row 4 hidden, column D hidden, A1:B1 merged, frozen at B2,
 * a table on F1:G3, and its sheet XML with every optional element the layout
 * writes into already there.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const plain = wb.addWorksheet('Plain');
  for (let r = 1; r <= 3; r++) plain.addRow([r, r * 2, `x${r}`]);
  const laid = wb.addWorksheet('Laid', { views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }] });
  for (let r = 1; r <= 5; r++) laid.addRow(['a', 'b', 'c', 'd']);
  laid.getColumn(2).width = 20;
  laid.getColumn(4).hidden = true;
  laid.getRow(2).height = 30;
  laid.getRow(4).hidden = true;
  laid.mergeCells('A1:B1');
  laid.addTable({ name: 'T', ref: 'F1', columns: [{ name: 'N' }, { name: 'M' }], rows: [[1, 2], [3, 4]] });
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const open = (b = bytes) => WorkbookSession.open(b, 'layout.xlsx', 'en-US');
const layout = (s: WorkbookSession, sheet: number, op: LayoutOp): EditResult => s.apply({ type: 'layout', sheet, op });
const reopen = (s: WorkbookSession) => open(s.save());
const info = (s: WorkbookSession, name: string) => s.model.sheets.find((m) => m.info.name === name)!.info;
const sheetXml = (s: WorkbookSession, n: number) => parts(s.save())[`xl/worksheets/sheet${n}.xml`];

describe('layout changes', () => {
  it('are saved: widths, heights, hidden rows and columns, frozen panes and merges', () => {
    const s = open();
    layout(s, 0, { kind: 'colWidths', widths: { 0: 140, 2: 35 } });
    layout(s, 0, { kind: 'rowHeights', heights: { 1: 40, 9: 12 } });
    layout(s, 0, { kind: 'hidden', axis: 'rows', lines: { 2: true, 20: true } });
    layout(s, 0, { kind: 'hidden', axis: 'cols', lines: { 1: true } });
    layout(s, 0, { kind: 'freeze', rows: 1, cols: 2 });
    layout(s, 0, { kind: 'merges', add: [{ r1: 4, c1: 0, r2: 5, c2: 2 }], remove: [] });
    const t = info(reopen(s), 'Plain');
    expect(t.colWidths).toMatchObject({ 0: 140, 2: 35 });
    expect(t.rowHeights).toMatchObject({ 1: 40, 9: 12 });
    expect(t.hiddenRows).toEqual([2, 20]);
    expect(t.hiddenCols).toEqual([1]);
    expect([t.frozenRows, t.frozenCols]).toEqual([1, 2]);
    expect(t.merges).toEqual([{ r1: 4, c1: 0, r2: 5, c2: 2 }]);
    // Rows that held no cells are written for their height or visibility, in order.
    const xml = sheetXml(s, 1);
    expect(xml.indexOf('<row r="10"')).toBeGreaterThan(xml.indexOf('<row r="3"'));
    expect(xml).toMatch(/<row r="21"[^>]*hidden="1"/);
    // Elements the sheet lacked go where the schema puts them.
    expect(xml.indexOf('<cols>')).toBeLessThan(xml.indexOf('<sheetData'));
    expect(xml.indexOf('<mergeCells')).toBeGreaterThan(xml.indexOf('</sheetData>'));
    expect(xml.indexOf('<pane')).toBeGreaterThan(xml.indexOf('<sheetView'));
  });

  it('change what the file already had, keeping the rest', () => {
    const s = open();
    layout(s, 1, { kind: 'colWidths', widths: { 1: null, 2: 70 } });
    layout(s, 1, { kind: 'rowHeights', heights: { 1: null } });
    layout(s, 1, { kind: 'hidden', axis: 'rows', lines: { 3: false } });
    layout(s, 1, { kind: 'hidden', axis: 'cols', lines: { 3: false } });
    layout(s, 1, { kind: 'freeze', rows: 0, cols: 0 });
    layout(s, 1, { kind: 'merges', add: [], remove: [{ r1: 0, c1: 0, r2: 0, c2: 1 }] });
    const t = info(reopen(s), 'Laid');
    expect(t.colWidths[1]).toBeUndefined();
    expect(t.colWidths[2]).toBe(70);
    expect(t.rowHeights[1]).toBeUndefined();
    expect([t.hiddenRows, t.hiddenCols, t.merges]).toEqual([[], [], []]);
    expect([t.frozenRows, t.frozenCols]).toEqual([0, 0]);
    expect(sheetXml(s, 2)).not.toMatch(/<mergeCells|<pane/);
  });

  it('undo with the change each one returns, and list what differs from the file', () => {
    const s = open();
    const done = [
      layout(s, 1, { kind: 'colWidths', widths: { 1: 200 } }),
      layout(s, 1, { kind: 'hidden', axis: 'rows', lines: { 6: true, 7: true } }),
      layout(s, 1, { kind: 'freeze', rows: 3, cols: 0 }),
      layout(s, 1, { kind: 'merges', add: [{ r1: 2, c1: 0, r2: 2, c2: 2 }], remove: [{ r1: 0, c1: 0, r2: 0, c2: 1 }] }),
    ];
    expect(s.changes().map((c) => `${c.ref} ${english.text(c.note!)}`)).toEqual([
      'B:B Changed the width of 1 column',
      '7:8 Hid 2 rows',
      'A4 Froze the rows above and the columns left of A4',
      'A3:C3 Merged the cells',
      'A1:B1 Unmerged the cells',
    ]);
    expect(s.editedCount).toBe(5);
    for (const r of done.reverse()) s.apply({ type: 'restore', cells: [], layout: r.layout!.undo });
    expect(s.changes()).toEqual([]);
    expect(s.editedCount).toBe(0);
    const t = info(reopen(s), 'Laid');
    expect([t.colWidths[1], t.frozenRows, t.merges, t.hiddenRows]).toEqual([140, 1, [{ r1: 0, c1: 0, r2: 0, c2: 1 }], [3]]);
  });

  it('merge cells only once the others are cleared, and undo puts both back', () => {
    const s = open();
    const r = s.apply({
      type: 'batch',
      ops: [
        { type: 'clear', sheet: 0, rects: [{ r1: 0, c1: 1, r2: 0, c2: 2 }] },
        { type: 'layout', sheet: 0, op: { kind: 'merges', add: [{ r1: 0, c1: 0, r2: 0, c2: 2 }], remove: [] } },
      ],
    });
    expect(r.undo).toHaveLength(2);
    // The cells it covers can't be typed into.
    expect(s.apply({ type: 'input', sheet: 0, row: 0, col: 1, text: '9' }).undo).toEqual([]);
    s.apply({ type: 'restore', cells: r.undo, layout: r.layout!.undo });
    expect([s.current(0, 0, 1).content, s.current(0, 0, 2).content]).toEqual([
      { kind: 'number', value: 2 },
      { kind: 'string', value: 'x1' },
    ]);
    expect(info(reopen(s), 'Plain').merges).toEqual([]);
  });

  it('are refused over a table, on a protected sheet, and on sheets that are not worksheets', () => {
    const s = open();
    const merge: LayoutOp = { kind: 'merges', add: [{ r1: 0, c1: 5, r2: 1, c2: 5 }], remove: [] };
    expect(english.text(layout(s, 1, merge).rejected!)).toBe("Cells in an Excel table can't be merged (T).");
    expect(english.text(layout(s, 2, { kind: 'colWidths', widths: { 0: 10 } }).rejected!)).toBe('This sheet is protected in Excel, so it is read-only here.');
    // Freezing only changes the view, which protection allows.
    expect(layout(s, 2, { kind: 'freeze', rows: 1, cols: 0 }).rejected).toBeUndefined();
    expect(s.changes()).toHaveLength(1);
  });

  it('follow rows and columns inserted after them', () => {
    const s = open();
    layout(s, 0, { kind: 'colWidths', widths: { 1: 150 } });
    layout(s, 0, { kind: 'merges', add: [{ r1: 1, c1: 1, r2: 2, c2: 1 }], remove: [] });
    const next = s.restructure({ type: 'insertCols', sheet: 0, at: 0, count: 1 });
    const t = info(next, 'Plain');
    expect(t.colWidths[2]).toBe(150);
    expect(t.merges).toEqual([{ r1: 1, c1: 2, r2: 2, c2: 2 }]);
    // What changed before is still listed.
    expect(next.changes().map((c) => english.text(c.note!))).toEqual(expect.arrayContaining(['Changed the width of 1 column', 'Merged the cells']));
  });
});

describe('filters', () => {
  /** Region!A1:B6 with values; Table!A1:B4 table "Sales" without a filter of its own picked. */
  async function filterBook(): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    const region = wb.addWorksheet('Region');
    region.addRow(['Region', 'Sales']);
    for (const [r, n] of [['North', 1], ['South', 2], ['North', 3], ['East', 4], ['', 5]] as const) region.addRow([r, n]);
    wb.addWorksheet('Table').addTable({ name: 'Sales', ref: 'A1', columns: [{ name: 'Who', filterButton: true }, { name: 'Amount' }], rows: [['Ann', 1], ['Bob', 2], ['Ann', 3]] });
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }

  it('are saved as Excel saves them: the values picked, and the rows they hide', async () => {
    const s = open(await filterBook());
    const rect = { r1: 0, c1: 0, r2: 5, c2: 1 };
    layout(s, 0, { kind: 'filter', filter: { rect, columns: { 0: ['North', ''] } }, rows: [2, 4] });
    const xml = sheetXml(s, 1);
    expect(xml).toMatch(/<autoFilter ref="A1:B6"><filterColumn colId="0"><filters blank="1"><filter val="North"\/><\/filters><\/filterColumn><\/autoFilter>/);
    const t = info(open(s.save()), 'Region');
    expect(t.filter).toEqual({ rect, columns: { 0: ['North', ''] } });
    expect(t.hiddenRows).toEqual([2, 4]);
    expect(s.changes().map((c) => english.text(c.note!))).toEqual(['Filtered, hiding 2 rows']);
  });

  it('cleared, keep the filter buttons and show every row; reopened, the file filter can be changed', async () => {
    const s = open(await filterBook());
    const rect = { r1: 0, c1: 0, r2: 5, c2: 1 };
    layout(s, 0, { kind: 'filter', filter: { rect, columns: { 0: ['South'] } }, rows: [1, 3, 4, 5] });
    const t = open(s.save());
    // Hidden by the filter, not by hand: clearing it shows them.
    const r = layout(t, 0, { kind: 'filter', filter: { rect, columns: {} }, rows: [] });
    expect(r.layout!.undo).toEqual([{ sheet: 0, op: { kind: 'filter', filter: { rect, columns: { 0: ['South'] } }, rows: [1, 3, 4, 5] } }]);
    expect(sheetXml(t, 1)).toMatch(/<autoFilter ref="A1:B6"\/>/);
    const u = info(open(t.save()), 'Region');
    expect([u.hiddenRows, u.filter, u.autoFilter]).toEqual([[], undefined, rect]);
    expect(t.changes().map((c) => english.text(c.note!))).toEqual(['Cleared the filter']);
  });

  it('turned off, leave the file without one, and say so', async () => {
    const s = open(await filterBook());
    const rect = { r1: 0, c1: 0, r2: 5, c2: 1 };
    layout(s, 0, { kind: 'filter', filter: { rect, columns: {} }, rows: [] });
    expect(sheetXml(s, 1)).toMatch(/<autoFilter ref="A1:B6"\/>/);
    expect(s.changes().map((c) => english.text(c.note!))).toEqual(['Filter buttons added']);
    expect(s.loadResult().sheets[0].autoFilter).toEqual(rect);
    const t = open(s.save());
    expect(info(t, 'Region').autoFilter).toEqual(rect);
    layout(t, 0, { kind: 'filter', filter: null, rows: [] });
    // As the editor now has it, before saving: no buttons, and none to bring back from the file.
    expect(t.loadResult().sheets[0]).toMatchObject({ filter: undefined, autoFilter: undefined });
    expect(sheetXml(t, 1)).not.toMatch(/<autoFilter/);
    expect(t.changes().map((c) => english.text(c.note!))).toEqual(['Filter buttons removed']);
    expect(info(open(t.save()), 'Region').autoFilter).toBeUndefined();
  });

  it("by condition, are saved as Excel saves them and read back: custom rules, Top 10 and the average", async () => {
    const s = open(await filterBook());
    const rect = { r1: 0, c1: 0, r2: 5, c2: 1 };
    const conditions: Record<number, FilterCondition> = {
      0: { kind: 'custom', and: false, rules: [{ operator: 'equal', value: 'N*' }, { operator: 'notEqual', value: 'East' }] },
      1: { kind: 'top10', top: false, percent: true, count: 20, threshold: 1 },
    };
    layout(s, 0, { kind: 'filter', filter: { rect, columns: {}, conditions }, rows: [2, 3, 4, 5] });
    const xml = sheetXml(s, 1);
    expect(xml).toContain(
      '<filterColumn colId="0"><customFilters><customFilter val="N*"/><customFilter operator="notEqual" val="East"/></customFilters></filterColumn>' +
        '<filterColumn colId="1"><top10 top="0" percent="1" val="20" filterVal="1"/></filterColumn>',
    );
    expect(info(reopen(s), 'Region').filter).toEqual({ rect, columns: {}, conditions });
    const above: Record<number, FilterCondition> = { 1: { kind: 'average', above: true, average: 3 } };
    layout(s, 0, { kind: 'filter', filter: { rect, columns: { 0: ['North'] }, conditions: above }, rows: [2, 4, 5] });
    expect(sheetXml(s, 1)).toContain('<filterColumn colId="1"><dynamicFilter type="aboveAverage" val="3"/></filterColumn>');
    expect(info(reopen(s), 'Region').filter).toEqual({ rect, columns: { 0: ['North'] }, conditions: above });
  });

  it("the editor can't change, such as dates in a period, are kept as the file has them", async () => {
    const files = parts(await filterBook());
    const kept = '<autoFilter ref="A1:B6"><filterColumn colId="0"><dynamicFilter type="today"/></filterColumn></autoFilter>';
    const sheet = files['xl/worksheets/sheet1.xml'].replace('</sheetData>', `</sheetData>${kept}`);
    const tablePath = 'xl/tables/table1.xml';
    const table = files[tablePath].replace('<filterColumn colId="0" hiddenButton="0"/>', '<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>');
    // The package's folder entries go, as they hold nothing.
    const entries = Object.entries({ ...files, 'xl/worksheets/sheet1.xml': sheet, [tablePath]: table }).filter(([k]) => !k.endsWith('/'));
    const zipped = zipSync(Object.fromEntries(entries.map(([k, v]) => [k, strToU8(v)])));
    const s = open(zipped);
    expect([info(s, 'Region').filter, info(s, 'Region').filterKept]).toEqual([undefined, true]);
    expect(info(s, 'Table').tables[0].filterKept).toBe(true);
    // Another filter on the sheet leaves it as it was.
    layout(s, 0, { kind: 'filter', filter: { rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, columns: {}, table: 'Elsewhere' }, rows: [] });
    expect(sheetXml(s, 1)).toContain(kept);
  });

  it("in a table, go in the table's part", async () => {
    const s = open(await filterBook());
    const rect = { r1: 0, c1: 0, r2: 3, c2: 1 };
    layout(s, 1, { kind: 'filter', filter: { rect, columns: { 0: ['Ann'] }, table: 'Sales' }, rows: [2] });
    const files = parts(s.save());
    const table = Object.entries(files).find(([k]) => k.startsWith('xl/tables/'))![1];
    expect(table).toMatch(/<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="Ann"\/><\/filters><\/filterColumn><\/autoFilter>/);
    expect(files['xl/worksheets/sheet2.xml']).not.toMatch(/<autoFilter/);
    const t = info(open(s.save()), 'Table');
    expect([t.filter, t.hiddenRows]).toEqual([{ rect, columns: { 0: ['Ann'] }, table: 'Sales' }, [2]]);
    expect(t.tables[0].filterButtons).toBe(true);
    // The sheet's own filter, which it hasn't, is left as it was.
    expect(s.loadResult().sheets[1].autoFilter).toBeUndefined();
  });
});
