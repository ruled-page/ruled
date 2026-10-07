import { describe, expect, it } from 'vitest';
import { mergeNotes, parsePersons, parseThreads } from '../src/core/comments';
import { parseA1 } from '../src/core/refs';
import { WorkbookSession } from '../src/core/session';
import { fitNumber, rowOverflows } from '../src/state/overflow';
import { SheetView } from '../src/state/sheetView';
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
    expect(styleOf('B3').border).toEqual({
      left: { width: 1, pattern: 'solid', color: '#000000' },
      right: { width: 1, pattern: 'solid', color: '#000000' },
      top: { width: 1, pattern: 'solid', color: '#000000' },
      bottom: { width: 1, pattern: 'solid', color: '#000000' },
    });
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
      return rowOverflows(borders, row, r.styles, false, measure).map((o) => [o.col, o.from, o.to, o.align]);
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
    expect(r.findings.find((f) => f.kind === 'comment')?.detail).toMatch(/^2 notes\. Marked/);
  });
});

describe('fitting numbers into narrow columns', () => {
  it('drops decimals, then uses scientific notation, for General numbers', () => {
    expect(fitNumber('3.14159265', 3.14159265, true, 7 * 5, measure)).toBe('3.142');
    expect(fitNumber('123456789012345', 123456789012345, true, 7 * 11, measure)).toBe('1.23457E+14');
    expect(fitNumber('0.000012345', 0.000012345, true, 7 * 8, measure)).toBe('1.23E-05');
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
          { author: 'Jo Smith', text: 'Is this right?', date: '2026-03-04' },
          { author: 'Kim Lee', text: 'Yes & checked.', date: '2026-03-05' },
        ],
      },
      { row: 8, col: 2, kind: 'thread', resolved: true, entries: [{ author: 'Kim Lee', text: 'Done', date: undefined }] },
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
