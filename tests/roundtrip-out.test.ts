import { writeFileSync } from 'node:fs';
import { it } from 'vitest';
import { parseA1 } from '../src/core/refs';
import { open, sheetIndex } from './helpers';

// Writes edited copies of the samples for checking with other tools.
it.skipIf(!process.env.ROUNDTRIP_OUT)('writes edited samples', () => {
  const dir = process.env.ROUNDTRIP_OUT!;
  const at = (r: string) => parseA1(r)!;
  const k = open('kitchen-sink.xlsx');
  const data = sheetIndex(k, 'Data');
  const notes = sheetIndex(k, 'Notes');
  k.apply({ type: 'input', sheet: data, ...at('D2'), text: '500' });
  k.apply({ type: 'pasteText', sheet: notes, ...at('A20'), values: [['12%', '2026-03-15', '$1,234.50', '=B10*2', 'hello & <goodbye>', 'TRUE', '=CONCAT("a","b")']] });
  k.apply({ type: 'fill', sheet: notes, src: { r1: 19, c1: 0, r2: 19, c2: 6 }, dst: { r1: 19, c1: 0, r2: 23, c2: 6 } });
  k.apply({ type: 'clear', sheet: sheetIndex(k, 'Summary'), rects: [{ r1: 13, c1: 1, r2: 13, c2: 1 }] });
  writeFileSync(dir + '/kitchen-sink-edited.xlsx', k.save());
  const p = open('pivot.xlsx');
  p.apply({ type: 'input', sheet: 0, ...at('E2'), text: '=C2*D2*2' });
  p.apply({ type: 'input', sheet: 0, ...at('C10'), text: '77' });
  writeFileSync(dir + '/pivot-edited.xlsx', p.save());
});
