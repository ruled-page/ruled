import { describe, expect, it } from 'vitest';
import { findCell, parseSharedStrings, scanSheet } from '../src/core/sheetScanner';
import { CellFlag, VT } from '../src/core/types';

const sheet = (rows: string, p = '') =>
  `<${p}worksheet xmlns${p ? ':' + p.slice(0, -1) : ''}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><${p}sheetData>${rows}</${p}sheetData></${p}worksheet>`;

/** Each cell as [row, col, type, value, formula]. */
function cells(xml: string, sst: string[] = []) {
  const { grid } = scanSheet(xml, sst);
  const out: [number, number, number, unknown, string | undefined][] = [];
  for (let row = 0; row < grid.rowCount; row++) {
    for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) {
      const t = grid.types[i];
      out.push([row, grid.cols[i], t, t === VT.Number || t === VT.Bool ? grid.nums[i] : t === VT.None ? null : grid.strs[i], grid.formulas.get(i)]);
    }
  }
  return out;
}

describe('reading cells written by other programs', () => {
  it('places rows and cells that leave out their address after the one before', () => {
    expect(cells(sheet('<row><c><v>1</v></c><c><v>2</v></c></row><row r="4"><c r="C4"><v>3</v></c><c><v>4</v></c></row><row><c><v>5</v></c></row>'))).toEqual([
      [0, 0, VT.Number, 1, undefined],
      [0, 1, VT.Number, 2, undefined],
      [3, 2, VT.Number, 3, undefined],
      [3, 3, VT.Number, 4, undefined],
      [4, 0, VT.Number, 5, undefined],
    ]);
  });

  it('sorts cells written out of order, and keeps the later of two at one address', () => {
    expect(cells(sheet('<row r="2"><c r="B2"><v>2</v></c><c r="A2"><v>1</v></c><c r="A2"><v>9</v></c></row><row r="1"><c r="A1"><v>0</v></c></row>'))).toEqual([
      [0, 0, VT.Number, 0, undefined],
      [1, 0, VT.Number, 9, undefined],
      [1, 1, VT.Number, 2, undefined],
    ]);
  });

  it("reads ISO dates, spaced and single-quoted attributes, and namespace prefixes", () => {
    const xml = sheet(
      `<x:row r="1"><x:c r = 'A1' t = 'd'><x:v>2026-01-05</x:v></x:c><x:c r="B1" t="d"><x:v>2026-01-05T12:00:00</x:v></x:c><x:c r="C1" t="d"><x:v>2026-01-05T12:00:00+02:00</x:v></x:c>` +
        `<x:c r="D1" t="b"><x:v>true</x:v></x:c><x:c r="E1" t="e"><x:v> #N/A </x:v></x:c><x:c r="F1"><x:v> </x:v></x:c></x:row>`,
      'x:',
    );
    const read = cells(xml).map(([, col, t, v]) => [col, t, v]);
    expect(read).toEqual([
      [0, VT.Number, 46027],
      [1, VT.Number, 46027.5],
      [2, VT.Number, expect.closeTo(46027 + 10 / 24, 9)],
      [3, VT.Bool, 1],
      [4, VT.Error, '#N/A'],
      [5, VT.None, null],
    ]);
  });

  it('reads shared and inline strings, without phonetic guides, and counts rich text', () => {
    const sst = parseSharedStrings(
      '<sst><si><t>plain</t></si><si><r><t>ri</t></r><r><rPr><b/></rPr><t>ch</t></r></si><si/><si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh></si></sst>',
    );
    expect(sst.strings).toEqual(['plain', 'rich', '', '漢字']);
    expect([...sst.rich]).toEqual([1]);
    expect(parseSharedStrings(undefined).strings).toEqual([]);
    const xml = sheet('<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>99</v></c><c r="C1" t="inlineStr"><is><t>in</t><t>line</t></is></c><c r="D1" t="inlineStr"/></row>');
    const scan = scanSheet(xml, sst.strings, sst.rich);
    expect(scan.richTextCells).toBe(1);
    expect(cells(xml, sst.strings).map(([, , , v]) => v)).toEqual(['rich', '', 'inline', null]);
  });

  it('marks followers of a shared formula whose master is missing as not calculable', () => {
    const { grid } = scanSheet(sheet('<row r="1"><c r="A1"><f t="shared" si="3"/><v>1</v></c></row>'), []);
    expect(grid.flags[0] & CellFlag.Unsupported).toBe(CellFlag.Unsupported);
    expect(findCell(grid, 0, 0)).toBe(0);
    expect(findCell(grid, 5, 5)).toBe(-1);
  });

  it('does not mistake longer tag names for rows and cells', () => {
    const xml = sheet('<rowx r="1"/><row r="2"><cx/><c r="A2"><v>7</v></c></row>');
    expect(cells(xml)).toEqual([[1, 0, VT.Number, 7, undefined]]);
  });
});
