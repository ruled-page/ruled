import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8 } from 'fflate';
import { it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parseA1, quoteSheetName } from '../src/core/refs';
import type { CommentChange, LayoutOp, NameChange } from '../src/core/types';
import { builtInLists } from '../src/state/fillLists';
import { open, sample, sheetIndex } from './helpers';

/**
 * Writes edited copies of the samples and of the workbooks in tests/corpus,
 * named <original>.<what was done>.xlsx, for checking with other tools:
 * `npm run validate:saves` validates each against the Office Open XML
 * schemas and compares it with its original.
 */
// oxlint-disable-next-line vitest/expect-expect -- writes files for npm run validate:saves to check
it.skipIf(!process.env.ROUNDTRIP_OUT)('writes edited workbooks', () => {
  const dir = process.env.ROUNDTRIP_OUT!;
  const at = (r: string) => parseA1(r)!;
  const write = (name: string, s: WorkbookSession) => writeFileSync(join(dir, name), s.save());
  const corpus = (name: string) => WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', name))), name, 'en-US', { recalc: 'on' });

  const k = open('kitchen-sink.xlsx');
  const data = sheetIndex(k, 'Data');
  const notes = sheetIndex(k, 'Notes');
  k.apply({ type: 'input', sheet: data, ...at('D2'), text: '500' });
  k.apply({ type: 'pasteText', sheet: notes, ...at('A20'), values: [['12%', '2026-03-15', '$1,234.50', '=B10*2', 'hello & <goodbye>', 'TRUE', '=CONCAT("a","b")']] });
  k.apply({ type: 'fill', sheet: notes, src: { r1: 19, c1: 0, r2: 19, c2: 6 }, dst: { r1: 19, c1: 0, r2: 23, c2: 6 } });
  k.apply({ type: 'clear', sheet: sheetIndex(k, 'Summary'), rects: [{ r1: 13, c1: 1, r2: 13, c2: 1 }] });
  write('kitchen-sink.edited.xlsx', k);

  // Sizes, hidden rows and columns, frozen panes and merges, on sheets that had some of each and on one that had none.
  for (const [name, sheets] of [['kitchen-sink.xlsx', ['Data', 'Notes', 'Summary']], ['orders.xlsx', []]] as const) {
    const l = open(name);
    for (const s of sheets.length ? sheets.map((n) => sheetIndex(l, n)) : [0]) {
      const layout = (op: LayoutOp) => l.apply({ type: 'layout', sheet: s, op });
      layout({ kind: 'colWidths', widths: { 0: 120, 1: null, 5: 30 } });
      layout({ kind: 'rowHeights', heights: { 0: 32, 1: null, 40: 10 } });
      layout({ kind: 'hidden', axis: 'rows', lines: { 2: true, 60: true } });
      layout({ kind: 'hidden', axis: 'cols', lines: { 3: true } });
      layout({ kind: 'freeze', rows: 1, cols: 1 });
      layout({ kind: 'merges', add: [{ r1: 30, c1: 0, r2: 31, c2: 1 }], remove: [] });
      const rects = [{ r1: 0, c1: 0, r2: 4, c2: 3 }];
      l.apply({ type: 'format', sheet: s, rects, change: { bold: true, fill: '#FFF2CC', hAlign: 'center', wrap: true, numFmt: '#,##0.00' } });
      l.apply({ type: 'format', sheet: s, rects, change: { fontFamily: 'Georgia', fontSizePt: 13, color: '#1F4E79' }, borders: { where: 'outside', edge: { width: 2, pattern: 'dashed' } } });
      l.apply({ type: 'format', sheet: s, rects: [{ r1: 5, c1: 0, r2: 6, c2: 1 }], change: {}, clear: true });
    }
    write(name.replace('.xlsx', '.layout.xlsx'), l);
  }

  // Filters: one on a sheet, and one in a table, which goes in the table's part.
  const f = open('orders.xlsx');
  f.apply({ type: 'layout', sheet: 0, op: { kind: 'filter', filter: { rect: { r1: 0, c1: 0, r2: 3, c2: 3 }, columns: { 0: ['Bolts', ''] } }, rows: [2, 3] } });
  write('orders.filtered.xlsx', f);

  // Fill Series: dates by month along a row, a growth series down to a stop value, and German day names continued.
  const sr = open('orders.xlsx');
  sr.apply({ type: 'input', sheet: 0, ...at('A50'), text: '2026-01-31' });
  sr.apply({ type: 'series', sheet: 0, rect: { r1: 49, c1: 0, r2: 49, c2: 5 }, series: { by: 'rows', type: 'date', dateUnit: 'month', step: 1, trend: false } });
  sr.apply({ type: 'input', sheet: 0, ...at('H50'), text: '3' });
  sr.apply({ type: 'series', sheet: 0, rect: { r1: 49, c1: 7, r2: 59, c2: 7 }, series: { by: 'columns', type: 'growth', dateUnit: 'day', step: 2, stop: 1000, trend: false } });
  sr.apply({ type: 'input', sheet: 0, ...at('J50'), text: 'Montag' });
  sr.apply({ type: 'fill', sheet: 0, src: { r1: 49, c1: 9, r2: 49, c2: 9 }, dst: { r1: 49, c1: 9, r2: 55, c2: 9 }, lists: builtInLists(['de']) });
  write('orders.series.xlsx', sr);
  const tf = corpus('templates.xlsx');
  const orders = tf.model.sheets.find((m) => m.info.tables.some((t) => t.name === 'Orders'))!;
  const table = orders.info.tables.find((t) => t.name === 'Orders')!;
  const rect = { ...table.rect, r2: table.rect.r2 - table.totalsRows };
  tf.apply({ type: 'layout', sheet: orders.info.index, op: { kind: 'filter', filter: { rect, columns: { [rect.c1]: ['Nuts'] }, table: 'Orders' }, rows: [rect.r1 + 1] } });
  write('templates.filtered.xlsx', tf);

  // Defined names: one changed, one deleted, one renamed where formulas use it, and new ones for the workbook and for a sheet.
  const n = open('orders.xlsx');
  const name = (change: NameChange) => n.apply({ type: 'name', change });
  name({ from: { name: 'MaxDiscount' }, to: { name: 'MaxDiscount', formula: 'Lists!$I$3', comment: 'Highest discount & "rules"\non two lines' } });
  name({ from: { name: 'Gizmo' }, to: null });
  name({ from: { name: 'Widget' }, to: { name: 'Widgets', formula: 'Lists!$E$2:$E$4' } });
  name({ to: { name: 'Everything', formula: 'Lists!$C$2:$G$5' } });
  name({ to: { name: 'First', formula: `${quoteSheetName(n.model.sheets[0].info.name)}!$A$1`, localSheet: 0 } });
  write('orders.names.xlsx', n);
  // A workbook without names gets the element that holds them.
  for (const [file, rename] of [['budget.xlsx', ''], ['kitchen-sink.xlsx', 'TaxRate']] as const) {
    const b = open(file);
    if (rename) b.apply({ type: 'name', change: { from: { name: rename }, to: { name: 'Tax_Rate', formula: b.model.names[0].formula } } });
    b.apply({ type: 'name', change: { to: { name: 'Start', formula: `${quoteSheetName(b.model.sheets[0].info.name)}!$A$1:$B$2` } } });
    write(file.replace('.xlsx', '.names.xlsx'), b);
  }

  // Links changed, added to sheets that had none (one of them with no relationships yet), and removed.
  const links = (s: WorkbookSession, sheet: number, op: Omit<Extract<LayoutOp, { kind: 'links' }>, 'kind'>) => s.apply({ type: 'layout', sheet, op: { kind: 'links', ...op } });
  const range = (a: string, b = a) => ({ r1: at(a).row, c1: at(a).col, r2: at(b).row, c2: at(b).col });
  const h = open('kitchen-sink.xlsx');
  const summary = sheetIndex(h, 'Summary');
  links(h, summary, { add: [{ rect: range('A19'), target: 'https://example.com/changed page', tooltip: 'Changed "tip"' }], remove: h.model.sheets[summary].info.links });
  links(h, data, { add: [{ rect: range('B2', 'C3'), location: "'Summary'!A1", display: 'Summary' }, { rect: range('D5'), target: 'mailto:someone@example.com?subject=Prices & more' }], remove: [] });
  h.apply({ type: 'linkFormat', sheet: data, rects: [range('B2', 'D5')], on: true });
  h.apply({ type: 'linkFormat', sheet: data, rects: [range('C3')], on: false });
  h.apply({ type: 'copy', srcSheet: data, src: range('B2'), sheet: notes, dst: range('A1', 'A3') });
  write('kitchen-sink.links.xlsx', h);
  const o = open('orders.xlsx');
  links(o, 0, { add: [{ rect: range('A1'), target: 'https://example.com/orders' }, { rect: range('B1'), location: 'A10' }], remove: [] });
  write('orders.links.xlsx', o);
  const tl = corpus('templates.xlsx');
  const linked = sheetIndex(tl, 'Lines');
  links(tl, linked, { add: [], remove: tl.model.sheets[linked].info.links });
  links(tl, sheetIndex(tl, 'Request'), { add: [{ rect: range('A1'), target: 'https://example.com/request', display: 'Request' }], remove: [] });
  write('templates.links.xlsx', tl);

  const p = open('pivot.xlsx');
  p.apply({ type: 'input', sheet: 0, ...at('E2'), text: '=C2*D2*2' });
  p.apply({ type: 'input', sheet: 0, ...at('C10'), text: '77' });
  write('pivot.edited.xlsx', p);

  // Formulas Excel stores in forms of its own: dynamic arrays, LET and LAMBDA parameters.
  const t = corpus('templates.xlsx');
  const lines = sheetIndex(t, 'Lines');
  const arrays = sheetIndex(t, 'Arrays');
  t.ensureEngine();
  t.apply({ type: 'input', sheet: lines, ...at('C3'), text: '350' });
  t.apply({ type: 'input', sheet: arrays, ...at('H1'), text: '=SEQUENCE(3)' });
  t.apply({ type: 'input', sheet: arrays, ...at('I1'), text: '=SORT(Orders[Qty],,-1)' });
  t.apply({ type: 'input', sheet: arrays, ...at('J1'), text: '=LET(a,2,b,a*Budget,b+1)' });
  t.apply({ type: 'input', sheet: arrays, ...at('K1'), text: '=MAP(Orders[Qty],LAMBDA(q,q*2))' });
  t.apply({ type: 'input', sheet: arrays, ...at('L1'), text: '=A2:A4&"!"' });
  // A name defined here follows the row inserted below.
  t.apply({ type: 'name', change: { to: { name: 'Quantities', formula: 'Lines!$C$2:$C$4', localSheet: lines } } });
  write('templates.typed.xlsx', t);
  write('templates.row-inserted.xlsx', t.restructure({ type: 'insertRows', sheet: lines, at: 2, count: 1 }));

  // Notes and threaded comments: changed, added to sheets that had none, and removed.
  const note = (s: WorkbookSession, sheet: string, ref: string, change: CommentChange) => s.apply({ type: 'comment', sheet: sheetIndex(s, sheet), ...at(ref), change });
  const c = corpus('comments.xlsx');
  note(c, 'Review', 'B2', { action: 'setNote', text: 'Agreed.\nRevisit in March.' });
  note(c, 'Review', 'C3', { action: 'reply', text: 'Thanks & noted <ok>.', author: 'Ann Analyst' });
  note(c, 'Review', 'C3', { action: 'editEntry', id: c.noteAt(sheetIndex(c, 'Review'), 2, 2)!.entries[1].id!, text: 'Yes, since Tuesday.' });
  note(c, 'Review', 'C3', { action: 'resolve', resolved: true });
  note(c, 'Review', 'A6', { action: 'delete' });
  note(c, 'Review', 'D5', { action: 'startThread', text: 'New line?', author: 'Cara' });
  note(c, 'Plain', 'B2', { action: 'setNote', text: 'Counted twice.', author: 'Cara' });
  note(c, 'Plain', 'A1', { action: 'startThread', text: 'Units?', author: 'Cara' });
  note(c, 'Plain', 'A1', { action: 'reply', text: 'Each.', author: 'Dee' });
  note(c, 'Objects', 'A1', { action: 'setNote', text: 'Allowed here.' });
  write('comments.edited.xlsx', c);
  const k2 = open('kitchen-sink.xlsx');
  note(k2, 'Summary', 'A4', { action: 'delete' });
  note(k2, 'Data', 'B3', { action: 'setNote', text: 'Checked.', author: 'Cara' });
  write('kitchen-sink.notes.xlsx', k2);
  const t2 = corpus('templates.xlsx');
  note(t2, 'Arrays', 'D2', { action: 'setNote', text: 'Largest line.', author: 'Cara' });
  note(t2, 'Arrays', 'A1', { action: 'startThread', text: 'Sorted?', author: 'Cara' });
  write('templates.commented.xlsx', t2);

  // Sheets added, renamed, deleted, moved, hidden and unhidden, with what refers to them following.
  const ks = () => WorkbookSession.open(sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx', 'en-US', { recalc: 'on' });
  const renamed = ks();
  write('kitchen-sink.sheet-renamed.xlsx', renamed.restructure({ type: 'renameSheet', sheet: sheetIndex(renamed, 'Summary'), name: "Q4 Summary's" }));
  const deleted = ks();
  write('kitchen-sink.data-deleted.xlsx', deleted.restructure({ type: 'deleteSheet', sheet: sheetIndex(deleted, 'Data') }));
  const summaryDeleted = ks();
  write('kitchen-sink.summary-deleted.xlsx', summaryDeleted.restructure({ type: 'deleteSheet', sheet: sheetIndex(summaryDeleted, 'Summary') }));
  let arranged = ks();
  arranged = arranged.restructure({ type: 'addSheet', at: 1, name: 'Sheet1' });
  arranged.apply({ type: 'input', sheet: 1, ...at('A1'), text: '=Data!F2*2' });
  arranged = arranged.restructure({ type: 'moveSheet', sheet: sheetIndex(arranged, 'Monthly Chart'), to: 0 });
  arranged = arranged.restructure({ type: 'hideSheet', sheet: sheetIndex(arranged, 'Summary'), hidden: true });
  arranged = arranged.restructure({ type: 'hideSheet', sheet: sheetIndex(arranged, 'Settings'), hidden: false });
  arranged = arranged.restructure({ type: 'deleteSheet', sheet: sheetIndex(arranged, 'Monthly Chart') });
  write('kitchen-sink.sheets-arranged.xlsx', arranged);
  const pivots = () => WorkbookSession.open(sample('pivot.xlsx'), 'pivot.xlsx');
  write('pivot.source-renamed.xlsx', pivots().restructure({ type: 'renameSheet', sheet: 0, name: 'Raw data' }));
  write('pivot.source-deleted.xlsx', pivots().restructure({ type: 'deleteSheet', sheet: 0 }));
  write('pivot.pivot-deleted.xlsx', pivots().restructure({ type: 'deleteSheet', sheet: 1 }));
  const tables = corpus('templates.xlsx');
  write('templates.table-sheet-renamed.xlsx', tables.restructure({ type: 'renameSheet', sheet: lines, name: 'Order lines' }));
  write('templates.table-sheet-deleted.xlsx', tables.restructure({ type: 'deleteSheet', sheet: lines }));
  write('templates.sheet-added.xlsx', tables.restructure({ type: 'addSheet', at: 3, name: 'Sheet1' }));

  const d = corpus('dates-1904.xlsx');
  d.apply({ type: 'input', sheet: 0, ...at('A3'), text: '2026-05-01' });
  d.apply({ type: 'input', sheet: 0, ...at('B4'), text: '120' });
  write('dates-1904.edited.xlsx', d);

  // A CSV file saved as the workbook it opens as, unchanged and edited.
  const csv = () =>
    WorkbookSession.open(strToU8('Item;Price;Due;Share;Done\r\nBolts;1,50;31.12.2025;12%;TRUE\r\n"two\nlines";=1+2;;;#N/A\r\n'), 'prices.csv', 'de-DE');
  write('prices-csv.unchanged.xlsx', csv());
  const fromCsv = csv();
  fromCsv.apply({ type: 'input', sheet: 0, ...at('C3'), text: '=B2*2' });
  fromCsv.apply({ type: 'input', sheet: 0, ...at('F1'), text: '7%' });
  write('prices-csv.edited.xlsx', fromCsv);
});
