import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parseA1 } from '../src/core/refs';
import { open, sheetIndex } from './helpers';

/**
 * Writes edited copies of the samples and of the workbooks in tests/corpus,
 * named <original>.<what was done>.xlsx, for checking with other tools:
 * `npm run validate:saves` validates each against the Office Open XML
 * schemas and compares it with its original.
 */
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

  const d = corpus('dates-1904.xlsx');
  d.apply({ type: 'input', sheet: 0, ...at('A3'), text: '2026-05-01' });
  d.apply({ type: 'input', sheet: 0, ...at('B4'), text: '120' });
  write('dates-1904.edited.xlsx', d);
});
