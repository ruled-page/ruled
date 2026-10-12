import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { mergeNotes, parseNotes, parsePersons, parseThreads } from '../src/core/comments';
import { distinctTexts, mergeEditResults } from '../src/core/editResult';
import { planFill, type FillSource } from '../src/core/fill';
import { readCustomProperties, setCustomProperties } from '../src/core/properties';
import { cellKey, keyCol, keyRow, quoteSheetName } from '@sitepen/formula-engine/refs';
import type { CellContent, EditResult } from '../src/core/types';
import { decodeEntities, decodeOoxmlText, encodeOoxmlText, findTagEnd, parseXml, textOf } from '../src/core/xml';
import { OoxmlPackage, resolveTarget } from '../src/core/zip';
import { english, errorText, msg } from '../src/i18n';

function pkg(files: Record<string, string>): OoxmlPackage {
  return OoxmlPackage.open(zipSync(Object.fromEntries(Object.entries({ '[Content_Types].xml': '<Types></Types>', ...files }).map(([k, v]) => [k, strToU8(v)]))));
}

describe('XML', () => {
  it('reads text through comments, CDATA, declarations and character references', () => {
    const root = parseXml('<?xml version="1.0"?><!DOCTYPE x><r><!-- note --><t>a &#65;&#x42; &amp; &unknown; </t><t><![CDATA[<raw> & ]]></t></r>');
    expect(textOf(root)).toBe('a AB & &unknown; <raw> & ');
    expect(decodeEntities('&lt;&gt;&quot;&apos;')).toBe('<>"\'');
  });

  it('closes what it can of broken markup', () => {
    expect(parseXml('<a><b>text</c></a></a>').children[0]).toMatchObject({ name: 'b', text: 'text' });
    expect(findTagEnd('<a b="x>y"', 0)).toBe(9);
    expect(parseXml('just text')).toMatchObject({ name: '#document', children: [] });
  });

  it('escapes characters cell text cannot hold, and literal escapes, both ways', () => {
    const s = 'tab\there\u0001 _x0041_';
    expect(encodeOoxmlText(s)).toBe('tab\there_x0001_ _x005F_x0041_');
    expect(decodeOoxmlText('_x0001_ and _x0041_')).toBe('\u0001 and A');
  });
});

describe('references', () => {
  it('quote sheet names when Excel would', () => {
    expect(['Data', 'my_sheet.2', 'My Sheet', "Bob's", 'A1', '2024', 'R1C1x'].map(quoteSheetName)).toEqual(['Data', 'my_sheet.2', "'My Sheet'", "'Bob''s'", "'A1'", "'2024'", 'R1C1x']);
  });

  it('turn a cell key back into its row and column', () => {
    const key = cellKey(1_048_575, 16_383);
    expect([keyRow(key), keyCol(key)]).toEqual([1_048_575, 16_383]);
  });
});

describe('packages', () => {
  it('resolve relationship targets, absolute or relative', () => {
    expect(resolveTarget('xl/worksheets/sheet1.xml', '/xl/media/a.png')).toBe('xl/media/a.png');
    expect(resolveTarget('xl/worksheets/sheet1.xml', '../drawings/./d.xml')).toBe('xl/drawings/d.xml');
  });

  it('refuse a file with too many parts before unzipping it', () => {
    const bytes = zipSync({ '[Content_Types].xml': strToU8('<Types/>'), a: strToU8('1'), b: strToU8('2') });
    let error: unknown;
    try {
      OoxmlPackage.open(bytes, { maxParts: 2 });
    } catch (e) {
      error = e;
    }
    expect(english.text(errorText(error))).toBe('This file has more than 2 parts, which is more than a workbook should.');
    let notPackage: unknown;
    try {
      OoxmlPackage.open(zipSync({ a: strToU8('1') }));
    } catch (e) {
      notPackage = e;
    }
    expect(english.text(errorText(notPackage))).toBe('This file is not a valid Office Open XML package.');
  });

  it('read a part once without keeping it, and keep edits apart in copies', () => {
    const p = pkg({ 'a.xml': '<a/>' });
    expect(p.textOnce('a.xml')).toBe('<a/>');
    expect(p.textOnce('missing.xml')).toBeUndefined();
    p.setText('a.xml', '<b/>');
    expect(p.textOnce('a.xml')).toBe('<b/>');
    const copy = p.clone();
    copy.setText('a.xml', '<c/>');
    copy.remove('[Content_Types].xml');
    expect([p.text('a.xml'), copy.text('a.xml'), copy.has('[Content_Types].xml'), p.has('[Content_Types].xml')]).toEqual(['<b/>', '<c/>', false, true]);
    expect(p.rels('a.xml')).toEqual([]);
  });
});

describe('custom document properties', () => {
  const ROOT_RELS =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>';

  it('read as text, skipping properties without a value', () => {
    const p = pkg({
      '_rels/.rels': ROOT_RELS,
      'docProps/custom.xml': '<Properties><property name="A" pid="2"><vt:lpwstr>one</vt:lpwstr></property><property name="Empty" pid="3"/><property pid="4"><vt:i4>5</vt:i4></property></Properties>',
    });
    expect(readCustomProperties(p)).toEqual({ A: 'one' });
    expect(readCustomProperties(pkg({}))).toEqual({});
  });

  it('update a property in place and number new ones after the highest', () => {
    const p = pkg({
      '_rels/.rels': ROOT_RELS,
      'docProps/custom.xml':
        '<Properties xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="x" pid="5" name="A"><v:lpwstr>old</v:lpwstr></property></Properties>',
    });
    setCustomProperties(p, { A: 'new & improved', B: '2' });
    const xml = p.text('docProps/custom.xml')!;
    expect(xml).toContain('pid="5" name="A"><v:lpwstr>new &amp; improved</v:lpwstr>');
    expect(xml).toContain('pid="6" name="B"><v:lpwstr>2</v:lpwstr>');
    expect(readCustomProperties(p)).toEqual({ A: 'new & improved', B: '2' });
  });

  it('create the part, its relationship and its content type when they are missing', () => {
    const p = pkg({ '_rels/.rels': ROOT_RELS });
    setCustomProperties(p, {});
    expect(p.has('docProps/custom.xml')).toBe(false);
    setCustomProperties(p, { Schema: 'tasks@2' });
    expect(readCustomProperties(p)).toEqual({ Schema: 'tasks@2' });
    expect(p.text('[Content_Types].xml')).toContain('PartName="/docProps/custom.xml"');
    // A part without the value-types namespace gets it.
    const bare = pkg({ '_rels/.rels': ROOT_RELS, 'docProps/custom.xml': '<Properties></Properties>' });
    setCustomProperties(bare, { A: '1' });
    expect(bare.text('docProps/custom.xml')).toMatch(/^<Properties xmlns:vt="[^"]+"><property fmtid="[^"]+" pid="2" name="A"><vt:lpwstr>1<\/vt:lpwstr>/);
  });
});

describe('notes and comments', () => {
  it('read authors, and drop the author name Excel puts at the start of a note', () => {
    const notes = parseNotes(
      '<comments><authors><author>Ann</author><author></author></authors><commentList>' +
        '<comment ref="B2" authorId="0"><text><r><t>Ann:</t></r><r><t>\nCheck this</t></r></text></comment>' +
        '<comment ref="C3" authorId="1"><text><t>No author</t></text></comment>' +
        '<comment ref="nowhere" authorId="0"><text><t>Lost</t></text></comment>' +
        '<comment ref="D4"><text><t>Bob: keeps his name</t></text></comment>' +
        '</commentList></comments>',
    );
    expect(notes).toEqual([
      { row: 1, col: 1, kind: 'note', entries: [{ author: 'Ann', text: 'Check this' }] },
      { row: 2, col: 2, kind: 'note', entries: [{ author: undefined, text: 'No author' }] },
      { row: 3, col: 3, kind: 'note', entries: [{ author: undefined, text: 'Bob: keeps his name' }] },
    ]);
  });

  it('gather replies under their thread, which replaces the placeholder note', () => {
    const persons = parsePersons('<personList><person id="p1" displayName="Ann"/><person displayName="Nobody"/><person id="p2"/></personList>');
    expect([...persons]).toEqual([
      ['p1', 'Ann'],
      ['p2', ''],
    ]);
    expect(parsePersons(undefined).size).toBe(0);
    const threads = parseThreads(
      '<ThreadedComments>' +
        '<threadedComment ref="B2" dT="2026-01-05T10:00:00.00" personId="p1" id="t1" done="1"><text>Is this right?</text></threadedComment>' +
        '<threadedComment ref="B2" dT="2026-01-06T09:00:00.00" personId="p2" id="t2" parentId="t1"><text>Yes</text></threadedComment>' +
        '<threadedComment ref="bad" id="t3"><text>Lost</text></threadedComment>' +
        '<threadedComment ref="E5" id="t4"/>' +
        '</ThreadedComments>',
      persons,
    );
    expect(threads).toEqual([
      {
        row: 1,
        col: 1,
        kind: 'thread',
        resolved: true,
        entries: [
          { author: 'Ann', text: 'Is this right?', date: '2026-01-05', id: 't1', personId: 'p1', time: '2026-01-05T10:00:00.00' },
          { author: undefined, text: 'Yes', date: '2026-01-06', id: 't2', personId: 'p2', time: '2026-01-06T09:00:00.00' },
        ],
      },
      { row: 4, col: 4, kind: 'thread', resolved: false, entries: [{ author: undefined, text: '', date: undefined, id: 't4' }] },
    ]);
    const placeholder = { row: 1, col: 1, kind: 'note' as const, entries: [{ text: '[Threaded comment]' }] };
    const other = { row: 9, col: 0, kind: 'note' as const, entries: [{ text: 'Plain note' }] };
    expect(mergeNotes([placeholder, other], threads).map((n) => [n.kind, n.row])).toEqual([
      ['thread', 1],
      ['thread', 4],
      ['note', 9],
    ]);
  });
});

describe('edit results', () => {
  const snapshot = (value: number) => ({ sheet: 0, row: 0, col: 0, content: { kind: 'number' as const, value }, style: 0 });
  const result = (n: number, extra: Partial<EditResult> = {}): EditResult => ({
    updates: [{ sheet: 0, row: 0, col: 0, t: 1, n, style: 0, flags: 0 }],
    undo: [snapshot(-n)],
    redo: [snapshot(n)],
    skipped: 1,
    skippedReasons: [msg('core.reason.protected')],
    warnings: [],
    newStyles: [],
    editedCount: n,
    ...extra,
  });

  it('combine into one step: the last update to a cell wins, undo runs backwards', () => {
    const merged = mergeEditResults([result(1), result(2, { rejected: 'no' }), result(3, { rejected: 'later' })]);
    expect(merged.updates.map((u) => u.n)).toEqual([3]);
    expect(merged.redo.map((c) => c.content)).toEqual([1, 2, 3].map((value) => ({ kind: 'number', value })));
    // Undo restores the last edit's "before" first and the first edit's last, so the first wins.
    expect(merged.undo.map((c) => c.content)).toEqual([-3, -2, -1].map((value) => ({ kind: 'number', value })));
    expect(merged.skipped).toBe(3);
    expect(merged.skippedReasons).toEqual([msg('core.reason.protected')]);
    expect(merged.editedCount).toBe(3);
    expect(merged.rejected).toBe('no');
    expect(mergeEditResults([]).editedCount).toBe(0);
    expect(distinctTexts(['a', 'a', msg('core.value.empty'), msg('core.value.empty')])).toEqual(['a', msg('core.value.empty')]);
  });
});

describe('fill series', () => {
  const src = (...values: (number | string | boolean)[]): FillSource[] =>
    values.map((v) => ({ content: (typeof v === 'number' ? { kind: 'number', value: v } : typeof v === 'string' ? { kind: 'string', value: v } : { kind: 'bool', value: v }) as CellContent, isDate: false }));
  const filled = (sources: FillSource[], positions: number[]) => planFill(sources, positions).map((p) => (p.kind === 'copy' ? `copy ${p.source}` : (p.content as { value: unknown }).value));

  it('copies when there is nothing to continue', () => {
    expect(filled([], [1])).toEqual([]);
    expect(filled(src(1, 'a'), [2, 3])).toEqual(['copy 0', 'copy 1']);
    expect(filled(src(5), [1, 2])).toEqual(['copy 0', 'copy 0']);
    expect(filled(src(true, false), [2])).toEqual(['copy 0']);
    expect(filled(src('plain'), [1])).toEqual(['copy 0']);
  });

  it('steps dates one day at a time, and numbers along their trend, both ways', () => {
    expect(filled([{ content: { kind: 'number', value: 46000 }, isDate: true }], [1, 2, -1])).toEqual([46001, 46002, 45999]);
    expect(filled(src(0.1, 0.2), [2, 3, -1])).toEqual([0.3, 0.4, 0]);
  });

  it('continues month and day names in their case, and copies ones that skip about', () => {
    expect(filled(src('jan'), [1, 12])).toEqual(['feb', 'jan']);
    expect(filled(src('MONDAY', 'WEDNESDAY'), [2, 3])).toEqual(['FRIDAY', 'SUNDAY']);
    expect(filled(src('Mar', 'Apr', 'Jun'), [3])).toEqual(['copy 0']);
  });

  it('counts numbered text, keeping leading zeros, below zero too', () => {
    expect(filled(src('Item 007'), [1, 3])).toEqual(['Item 008', 'Item 010']);
    expect(filled(src('Q1x', 'Q3x'), [2])).toEqual(['Q5x']);
    expect(filled(src('v2'), [-3])).toEqual(['v-1']);
    expect(filled(src('a1', 'b2'), [2])).toEqual(['copy 0']);
  });
});
