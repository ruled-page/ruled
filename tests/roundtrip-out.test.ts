import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8 } from 'fflate';
import { it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parseA1 } from '../src/core/refs';
import type { CommentChange, LayoutOp } from '../src/core/types';
import { open, sheetIndex } from './helpers';

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
