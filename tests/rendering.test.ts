import { describe, expect, it } from 'vitest';
import { mergeNotes, parsePersons, parseThreads } from '../src/core/comments';
import { parseA1 } from '../src/core/refs';
import { WorkbookSession } from '../src/core/session';
import { StyleBook, tintColor } from '../src/core/styles';
import { builtinTableLook, tableCellFormat } from '../src/core/tableStyles';
import type { TableInfo } from '../src/core/types';
import { effectiveStyle } from '../src/state/cellStyle';
import { excelTextWidth, fitNumber, rowOverflows } from '../src/state/overflow';
import { SheetView } from '../src/state/sheetView';
import { english } from '../src/i18n';
import { sample } from './helpers';

function load(file: string) {
  const s = WorkbookSession.open(sample(file), file);
  const r = s.loadResult();
  const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  return { s, r, views };
}

/** Roughly 7px per character, enough to tell what spills where. */
const measure = (text: string) => text.length * 7;

describe('formatting.xlsx', () => {
  const { r, views } = load('formatting.xlsx');
  const borders = views[0];
  const styleOf = (ref: string) => {
    const { row, col } = parseA1(ref)!;
    return r.styles[borders.styleAt(row, col)];
  };

  it('reads each border style', () => {
    // No colour in the file: Excel's automatic colour, which the theme decides.
    const thin = { width: 1, pattern: 'solid', color: undefined };
    expect(styleOf('B3').border).toEqual({ left: thin, right: thin, top: thin, bottom: thin });
    const kinds = ['B5', 'B7', 'B9', 'B11', 'B13', 'B17'].map((ref) => {
      const top = styleOf(ref).border?.top;
      return `${top?.width} ${top?.pattern}`;
    });
    expect(kinds).toEqual(['2 solid', '3 solid', '3 double', '1 dashed', '1 dotted', '2 dashed']);
    expect(styleOf('D5').border?.left?.color).toBe('#C00000');
    expect(styleOf('A3').border).toBeUndefined();
  });

  it('gives empty cells their row or column format, the row winning', () => {
    expect(styleOf('H5').fill).toBe('#E2EFDA');
    expect(styleOf('B38').fill).toBe('#DDEBF7');
    expect(styleOf('H38').fill).toBe('#DDEBF7');
    expect(styleOf('C5').fill).toBeUndefined();
  });

  it('spills text into empty neighbours the way Excel does', () => {
    const at = (ref: string) => {
      const { row } = parseA1(ref)!;
      return rowOverflows(borders, row, (_c, rec) => r.styles[rec.style], false, measure).map((o) => [o.col, o.from, o.to, o.align]);
    };
    expect(at('A25')).toEqual([[0, 0, 4, 'left']]);
    expect(at('F26')).toEqual([[5, 3, 5, 'right']]);
    expect(at('C27')).toEqual([[2, 1, 3, 'center']]);
    // Blocked by B28, wrapped, a number, and a styled but empty cell that text runs through.
    expect(at('A28')).toEqual([]);
    expect(at('A29')).toEqual([]);
    expect(at('A30')).toEqual([]);
    expect(at('A31')).toEqual([[0, 0, 1, 'left']]);
  });

  it('reads notes with their authors', () => {
    expect(borders.noteAt(33, 0)).toEqual({
      row: 33,
      col: 0,
      kind: 'note',
      entries: [{ author: 'Alex Analyst', text: 'This is a note on A34.\nIt has two lines.' }],
    });
    expect(borders.noteAt(33, 1)?.entries[0]).toMatchObject({ author: 'Sam Reviewer', text: 'Checked against the ledger.' });
    expect(english.text(r.findings.find((f) => f.kind === 'comment')?.detail)).toMatch(/^2 notes\. Marked/);
  });
});

describe('fitting numbers into narrow columns', () => {
  it('drops decimals, then uses scientific notation, for General numbers', () => {
    expect(fitNumber('3.14159265', 3.14159265, true, 7 * 5, measure)).toBe('3.142');
    expect(fitNumber('123456789012345', 123456789012345, true, 7 * 11, measure)).toBe('1.23457E+14');
    expect(fitNumber('0.000012345', 0.000012345, true, 7 * 8, measure)).toBe('1.23E-05');
  });

  it("judges fit with Excel's Calibri widths, not the browser's font", () => {
    const style = { ...new StyleBook(undefined, undefined, 'en-US').get(0), bold: true };
    // Column width 13 is 96px; Excel shows this total in full.
    expect(excelTextWidth('$157,925.25', style)).toBeLessThan(96 - 9);
    expect(excelTextWidth('123456789012345', style)).toBeGreaterThan(64 - 9);
  });

  it('shows # for formatted numbers that do not fit', () => {
    expect(fitNumber('$12,345.00', 12345, false, 7 * 6, measure)).toBe('######');
    expect(fitNumber('12', 12, false, 7 * 6, measure)).toBe('12');
  });
});

describe('threaded comments', () => {
  const persons = parsePersons(`<personList xmlns="x">
    <person displayName="Jo Smith" id="{P1}" userId="jo" providerId="AD"/>
    <person displayName="Kim Lee" id="{P2}" userId="kim" providerId="AD"/>
  </personList>`);
  const threads = parseThreads(
    `<ThreadedComments xmlns="x">
      <threadedComment ref="B2" dT="2026-03-04T10:00:00.00" personId="{P1}" id="{T1}"><text>Is this right?</text></threadedComment>
      <threadedComment ref="B2" dT="2026-03-05T09:00:00.00" personId="{P2}" id="{T2}" parentId="{T1}"><text>Yes &amp; checked.</text></threadedComment>
      <threadedComment ref="C9" personId="{P2}" id="{T3}" done="1"><text>Done</text></threadedComment>
    </ThreadedComments>`,
    persons,
  );

  it('groups replies under the first comment', () => {
    expect(threads).toEqual([
      {
        row: 1,
        col: 1,
        kind: 'thread',
        resolved: false,
        entries: [
          { author: 'Jo Smith', text: 'Is this right?', date: '2026-03-04', id: '{T1}', personId: '{P1}', time: '2026-03-04T10:00:00.00' },
          { author: 'Kim Lee', text: 'Yes & checked.', date: '2026-03-05', id: '{T2}', personId: '{P2}', time: '2026-03-05T09:00:00.00' },
        ],
      },
      { row: 8, col: 2, kind: 'thread', resolved: true, entries: [{ author: 'Kim Lee', text: 'Done', date: undefined, id: '{T3}', personId: '{P2}' }] },
    ]);
  });

  it("replace the placeholder notes Excel writes for older versions", () => {
    const notes = [
      { row: 1, col: 1, kind: 'note' as const, entries: [{ text: '[Threaded comment] ...' }] },
      { row: 4, col: 0, kind: 'note' as const, entries: [{ text: 'A real note' }] },
    ];
    expect(mergeNotes(notes, threads).map((n) => [n.row, n.col, n.kind])).toEqual([
      [1, 1, 'thread'],
      [8, 2, 'thread'],
      [4, 0, 'note'],
    ]);
  });
});

describe('theme tints', () => {
  it("match the colours Excel's palette shows", () => {
    const cases: [string, number, string][] = [
      ['#4472C4', 0.8, '#D9E1F2'], ['#4472C4', 0.6, '#B4C6E7'], ['#4472C4', 0.4, '#8EA9DB'],
      ['#4472C4', -0.25, '#305496'], ['#4472C4', -0.5, '#203764'], ['#ED7D31', 0.8, '#FCE4D6'],
      ['#70AD47', -0.25, '#548235'], ['#FFC000', 0.8, '#FFF2CC'], ['#A5A5A5', 0.8, '#EDEDED'],
      ['#000000', 0.5, '#808080'], ['#FFFFFF', -0.15, '#D9D9D9'], ['#000000', 0.35, '#595959'],
    ];
    expect(cases.map(([c, t]) => tintColor(c, t))).toEqual(cases.map(([, , want]) => want));
  });
});

describe('table styles', () => {
  const { r, views } = load('formatting.xlsx');
  const sheet = views[2];
  const theme = new StyleBook(undefined, undefined, 'en-US').theme;

  it('reads every table with its style options', () => {
    expect(sheet.info.tables.map((t) => [t.name, t.style, t.totalsRows, t.showRowStripes, t.showColumnStripes, t.showFirstColumn])).toEqual([
      ['Styled1', 'TableStyleMedium2', 1, true, false, false],
      ['Styled2', 'TableStyleLight9', 0, true, false, false],
      ['Styled3', 'TableStyleDark2', 0, true, true, true],
      ['Styled4', 'TableStyleMedium14', 0, true, false, false],
      ['Styled5', 'TableStyleLight1', 0, true, false, false],
      ['Styled6', 'TableStyleLight15', 0, false, true, false],
    ]);
    expect(sheet.info.tables.every((t) => t.look)).toBe(true);
  });

  it('draws headers, stripes and total rows from the theme accent', () => {
    // XlsxWriter files use the Office 2007 theme, whose accent 1 is 4F81BD.
    const at = (row: number, col: number) => effectiveStyle(sheet, r.styles, row, col);
    expect(at(0, 0)).toMatchObject({ fill: '#4F81BD', color: '#FFFFFF', bold: true });
    expect(at(1, 1).fill).toBe('#DCE6F1');
    expect(at(2, 1).fill).toBeUndefined();
    expect(at(2, 1).border?.top?.color).toBe('#95B3D7');
    expect(at(5, 0)).toMatchObject({ bold: true, border: { top: { pattern: 'double' } } });
    expect(at(7, 0).fill).toBeUndefined();
  });

  it('layers elements from the whole table up to the header', () => {
    const t: TableInfo = {
      name: 'T', rect: { r1: 0, c1: 0, r2: 4, c2: 2 }, headerRows: 1, totalsRows: 0, columns: [],
      showFirstColumn: true, showLastColumn: false, showRowStripes: true, showColumnStripes: false,
      look: builtinTableLook('TableStyleDark1', theme),
    };
    expect(tableCellFormat(t, 0, 0)).toMatchObject({ fill: '#000000', bold: true });
    expect(tableCellFormat(t, 1, 0)?.border?.right).toMatchObject({ width: 2, color: '#FFFFFF' });
    expect(tableCellFormat(t, 1, 1)?.fill).not.toBe(tableCellFormat(t, 2, 1)?.fill);
    expect(tableCellFormat(t, 5, 1)).toBeUndefined();
    expect(builtinTableLook('TableStyleMedium29', theme)).toBeDefined();
    expect(builtinTableLook('Fancy', theme)).toBeUndefined();
  });

  it('uses custom table styles defined in the workbook', () => {
    const styles = new StyleBook(
      `<styleSheet xmlns="x"><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
        <dxfs count="2">
          <dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><bgColor rgb="FF7030A0"/></patternFill></fill></dxf>
          <dxf><fill><patternFill><bgColor theme="4" tint="0.6"/></patternFill></fill></dxf>
        </dxfs>
        <tableStyles count="1"><tableStyle name="Purple" pivot="0" count="2">
          <tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="1" size="2"/>
        </tableStyle></tableStyles></styleSheet>`,
      undefined,
      'en-US',
    );
    const look = styles.tableStyles.get('Purple')!;
    expect(look.rowStripe).toBe(2);
    const t: TableInfo = {
      name: 'T', rect: { r1: 0, c1: 0, r2: 6, c2: 1 }, headerRows: 1, totalsRows: 0, columns: [],
      showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false, look,
    };
    expect(tableCellFormat(t, 0, 0)).toEqual({ fill: '#7030A0', color: '#FFFFFF', bold: true });
    // Stripes two rows deep: rows 1-2 striped, 3-4 not, 5-6 striped.
    expect([1, 2, 3, 4, 5].map((row) => tableCellFormat(t, row, 0)?.fill ?? '-')).toEqual(['#B4C6E7', '#B4C6E7', '-', '-', '#B4C6E7']);
  });
});

describe('conditional formatting', () => {
  const { s, r } = load('formatting.xlsx');
  const sheet = 1;
  const result = s.conditionalFormats(sheet);
  const at = (ref: string) => {
    const { row, col } = parseA1(ref)!;
    const c = result.cells.find((x) => x.row === row && x.col === col);
    return c && { ...c, style: c.style === undefined ? undefined : result.styles[c.style] };
  };
  // The sample's columns, as gen_formatting.py fills rows 2-31 (i = 0..29).
  const score = (i: number) => (i * 37) % 100;
  const late = (i: number) => i % 5 === 1 || i % 5 === 4;

  it('reads every rule on the sheet', () => {
    expect(r.sheets[sheet].conditionalFormats.map((c) => c.type)).toEqual([
      'cellIs', 'cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet',
      'containsText', 'containsBlanks', 'duplicateValues', 'top10', 'aboveAverage',
    ]);
    expect(result.skipped).toBe(0);
  });

  it('applies cell-value and formula rules in priority order', () => {
    for (let i = 0; i < 30; i++) {
      const b = at(`B${i + 2}`)?.style;
      const fill = late(i) ? '#FCE4D6' : undefined;
      // High scores are green and bold; low ones red and italic, over the late fill; the rest show only the late fill.
      const want: Record<string, unknown> = score(i) >= 80 ? { fill: '#C6EFCE', bold: true } : score(i) < 30 ? { color: '#9C0006', italic: true, fill } : { fill };
      const got = Object.fromEntries(Object.keys(want).map((k) => [k, b?.[k as keyof typeof b]]));
      expect(got, `B${i + 2}`).toEqual(want);
      expect(at(`A${i + 2}`)?.style?.fill).toBe(late(i) ? '#FCE4D6' : undefined);
    }
  });

  it('colours scales, sizes bars and picks icons from the range', () => {
    // Column D holds (i*13)%50: 0 is the minimum (red), 49 the maximum (green).
    expect(at('D2')?.style?.fill).toBe('#F8696B');
    expect(at('D25')?.style?.fill).toBe('#63BE7B');
    expect(at('E2')?.bar).toEqual({ from: 0, to: 0.1, color: '#638EC6', gradient: true });
    expect(at('E26')?.bar?.to).toBeCloseTo(0.9);
    // Column F holds (i*17)%30, 0-29: icons split at 33% and 67% of the span.
    const icon = (i: number) => at(`F${i + 2}`)?.icon?.index;
    expect([icon(0), icon(1), icon(2), icon(3)]).toEqual([0, 1, 0, 2]);
  });

  it('handles text, blanks, duplicates, top-N and averages', () => {
    // G7 contains "error"; G3 does too, but it is a late row and that rule has priority.
    expect(at('G7')?.style?.fill).toBe('#FFEB9C');
    expect(at('G3')?.style?.fill).toBe('#FCE4D6');
    expect(at('G4')?.style?.fill).toBe('#EDEDED');
    expect(at('G2')?.style?.fill).toBeUndefined();
    expect(at('H2')?.style?.strike).toBe(true);
    const top = [...Array(30).keys()].map((i) => (i * 23) % 61).sort((a, b) => b - a).slice(0, 3);
    const tops = [...Array(30).keys()].filter((i) => at(`I${i + 2}`)?.style?.bold).map((i) => (i * 23) % 61);
    expect(tops.sort((a, b) => b - a)).toEqual(top);
    const avg = [...Array(30).keys()].reduce((s, i) => s + ((i * 11) % 40), 0) / 30;
    const above = [...Array(30).keys()].filter((i) => at(`J${i + 2}`)?.style?.fill === '#C6EFCE');
    expect(above.every((i) => (i * 11) % 40 > avg)).toBe(true);
  });

  it('follows edits', () => {
    const t = WorkbookSession.open(sample('formatting.xlsx'), 'formatting.xlsx');
    t.apply({ type: 'input', sheet, row: 1, col: 2, text: 'Late' });
    const after = t.conditionalFormats(sheet);
    const a2 = after.cells.find((c) => c.row === 1 && c.col === 0);
    expect(a2 && after.styles[a2.style!].fill).toBe('#FCE4D6');
  });
});
