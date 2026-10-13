import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { applyCommentChange, commentTime, describeCommentChange, newGuid, noteText } from '../src/core/commentEdits';
import { parseA1, toA1 } from '@ruled/formula/refs';
import { WorkbookSession } from '../src/core/session';
import type { CellNote, CommentChange, EditResult } from '../src/core/types';
import { english } from '../src/i18n';
import { parts, sample, sheetIndex } from './helpers';

const corpusBytes = (name: string) => new Uint8Array(readFileSync(join(__dirname, 'corpus', name)));
const corpus = (name: string) => WorkbookSession.open(corpusBytes(name), name);
const comment = (s: WorkbookSession, sheet: string, ref: string, change: CommentChange): EditResult =>
  s.apply({ type: 'comment', sheet: sheetIndex(s, sheet), ...parseA1(ref)!, change });
/** The workbook saved and read again, as anything else opening it would read it. */
const reopened = (s: WorkbookSession) => WorkbookSession.open(s.save(), s.fileName);
const noteOn = (s: WorkbookSession, sheet: string, ref: string) => s.model.sheets[sheetIndex(s, sheet)].info.notes.find((n) => toA1(n.row, n.col) === ref);
/** Replaces parts of a workbook, for files the samples don't have. */
function withParts(bytes: Uint8Array, edits: Record<string, (xml: string) => string>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [path, edit] of Object.entries(edits)) files[path] = strToU8(edit(strFromU8(files[path])));
  return zipSync(files);
}

const ANN = '{6B1F2D3A-0C4E-4F7B-9A21-3E5D7C8B9A10}';
const C3_ROOT = '{2F4A6C8E-1B3D-4E5F-8A9B-0C1D2E3F4A5B}';
const C3_REPLY = '{7A8B9C0D-2E3F-4A5B-9C6D-1E2F3A4B5C6D}';
const A6_ROOT = '{9E8D7C6B-5A4F-4E3D-8C2B-1A0F9E8D7C6B}';
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

describe('notes', () => {
  it('are added to a sheet with none, with a comments part, a hidden box and the links Excel needs', () => {
    const s = corpus('comments.xlsx');
    const r = comment(s, 'Plain', 'B2', { action: 'setNote', text: 'Counted twice.\nCheck.', author: 'Cara' });
    const note: CellNote = { row: 1, col: 1, kind: 'note', entries: [{ author: 'Cara', text: 'Counted twice.\nCheck.' }] };
    expect(r.noteRedo).toEqual([{ sheet: 1, row: 1, col: 1, note }]);
    expect(r.noteUndo).toEqual([{ sheet: 1, row: 1, col: 1, note: null }]);
    expect(s.noteAt(1, 1, 1)).toEqual(note);
    expect(s.loadResult().sheets[1].notes).toEqual([note]);

    const bytes = s.save();
    const xml = parts(bytes);
    const before = parts(corpusBytes('comments.xlsx'));
    expect(xml['xl/worksheets/_rels/sheet2.xml.rels']).toContain('relationships/comments" Target="../comments3.xml"/>');
    expect(xml['xl/worksheets/_rels/sheet2.xml.rels']).toContain('relationships/vmlDrawing" Target="../drawings/vmlDrawing3.vml"/>');
    expect(xml['[Content_Types].xml']).toContain('<Override PartName="/xl/comments3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
    // After <pageMargins>, as a worksheet orders its parts.
    expect(xml['xl/worksheets/sheet2.xml']).toMatch(/<pageMargins [^>]*\/><legacyDrawing r:id="rId2"\/><\/worksheet>$/);
    expect(xml['xl/comments3.xml']).toContain('<authors><author>Cara</author></authors>');
    expect(xml['xl/comments3.xml']).toContain('<b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>Cara:</t>');
    const vml = xml['xl/drawings/vmlDrawing3.vml'];
    expect(vml).toContain('<o:idmap v:ext="edit" data="3"/>');
    expect(vml).toContain('<v:shape id="_x0000_s3073" type="#_x0000_t202"');
    expect(vml).toContain('visibility:hidden');
    expect(vml).toContain('<x:Anchor>2, 15, 0, 2, 4, 15, 4, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column>');
    // The other sheets' notes are as they were.
    for (const part of ['xl/comments1.xml', 'xl/comments2.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/threadedComments/threadedComment1.xml', 'xl/persons/person.xml']) {
      expect(xml[part]).toBe(before[part]);
    }
    expect(WorkbookSession.open(bytes, 'c.xlsx').model.sheets[1].info.notes).toEqual([note]);
  });

  it('are edited in place, keeping who wrote them and the other notes as written', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Review', 'B2', { action: 'setNote', text: 'Agreed.', author: 'Someone else' });
    const xml = parts(s.save());
    const before = parts(corpusBytes('comments.xlsx'));
    const placeholders = before['xl/comments1.xml'].slice(before['xl/comments1.xml'].indexOf('<comment ref="C3"'), before['xl/comments1.xml'].indexOf('</commentList>'));
    expect(xml['xl/comments1.xml']).toContain(`<comment ref="B2" authorId="0" shapeId="0"><text><r>`);
    expect(xml['xl/comments1.xml']).toContain(placeholders);
    expect(xml['xl/drawings/vmlDrawing1.vml']).toBe(before['xl/drawings/vmlDrawing1.vml']);
    expect(noteOn(reopened(s), 'Review', 'B2')?.entries).toEqual([{ author: 'Ann Analyst', text: 'Agreed.' }]);
  });

  it('without a writer have no author', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Plain', 'A2', { action: 'setNote', text: ' Spaces kept ' });
    const xml = parts(s.save());
    expect(xml['xl/comments3.xml']).toContain('<authors><author></author></authors>');
    expect(xml['xl/comments3.xml']).toContain('<t xml:space="preserve"> Spaces kept </t></r></text>');
    expect(noteOn(reopened(s), 'Plain', 'A2')?.entries).toEqual([{ author: undefined, text: ' Spaces kept ' }]);
  });

  it('go with their box, and the last one takes its comments part and drawing with it', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    comment(s, 'Summary', 'A4', { action: 'delete' });
    const xml = parts(s.save());
    expect(xml['xl/comments1.xml']).toBeUndefined();
    expect(xml['xl/drawings/vmlDrawing1.vml']).toBeUndefined();
    expect(xml['xl/worksheets/sheet1.xml']).not.toContain('legacyDrawing');
    expect(xml['xl/worksheets/_rels/sheet1.xml.rels']).not.toMatch(/comments|vmlDrawing/);
    expect(xml['xl/worksheets/_rels/sheet1.xml.rels']).toContain('relationships/drawing"');
    expect(xml['[Content_Types].xml']).not.toContain('comments1.xml');
    expect(reopened(s).model.sheets[sheetIndex(s, 'Summary')].info.notes).toEqual([]);
  });

  it('leave a drawing that has other shapes in it', () => {
    const checkbox = '<v:shape id="_x0000_s1030" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"><x:Row>9</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>';
    const bytes = withParts(sample('kitchen-sink.xlsx'), { 'xl/drawings/vmlDrawing1.vml': (xml) => xml.replace('</xml>', `${checkbox}</xml>`) });
    const s = WorkbookSession.open(bytes, 'k.xlsx');
    comment(s, 'Summary', 'A4', { action: 'delete' });
    comment(s, 'Summary', 'B6', { action: 'setNote', text: 'New' });
    const xml = parts(s.save());
    const vml = xml['xl/drawings/vmlDrawing1.vml'];
    expect(vml).toContain(checkbox);
    expect(vml).not.toContain('<x:Row>3</x:Row>');
    // Numbered after the shapes already there.
    expect(vml).toContain('<v:shape id="_x0000_s1031" type="#_x0000_t202"');
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<legacyDrawing r:id="rId3"/>');
    expect(reopened(s).model.sheets[sheetIndex(s, 'Summary')].info.notes.map((n) => [toA1(n.row, n.col), n.entries[0].text])).toEqual([['B6', 'New']]);
  });

  it('can be added where a sheet has no relationships at all', () => {
    const s = corpus('templates.xlsx');
    comment(s, 'Arrays', 'D2', { action: 'setNote', text: 'Biggest by total', author: 'Cara' });
    const xml = parts(s.save());
    expect(xml['xl/worksheets/_rels/sheet3.xml.rels']).toMatch(/^<\?xml[^>]*\?>\r?\n<Relationships xmlns="[^"]+">.*vmlDrawing1\.vml"\/><\/Relationships>$/);
    expect(xml['[Content_Types].xml']).toContain('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');
    expect(noteOn(reopened(s), 'Arrays', 'D2')?.entries[0]).toEqual({ author: 'Cara', text: 'Biggest by total' });
  });
});

describe('workbooks laid out unlike Excel writes them', () => {
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const THREADS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
  const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  const file = (): Uint8Array =>
    zipSync({
      '[Content_Types].xml': strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
      ),
      '_rels/.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', `${REL}/officeDocument`, 'xl/workbook.xml')}</Relationships>`),
      'xl/workbook.xml': strToU8(
        `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Bare" sheetId="1" r:id="rId1"/><sheet name="Odd" sheetId="2" r:id="rId2"/><sheet name="Broken" sheetId="3" r:id="rId3"/></sheets></workbook>`,
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', `${REL}/worksheet`, 'worksheets/sheet1.xml')}${rel('rId2', `${REL}/worksheet`, 'worksheets/sheet2.xml')}${rel('rId3', `${REL}/worksheet`, 'worksheets/sheet3.xml')}${rel('rId4', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml')}</Relationships>`,
      ),
      'xl/persons/person.xml': strToU8(`<personList xmlns="${THREADS}"/>`),
      // A prefixed namespace, no r: namespace, no cells, and an extLst the drawing must go before.
      'xl/worksheets/sheet1.xml': strToU8(`<x:worksheet xmlns:x="${MAIN}"><x:sheetData/><x:extLst><x:ext uri="{X}"/></x:extLst></x:worksheet>`),
      'xl/worksheets/sheet2.xml': strToU8(
        `<x:worksheet xmlns:x="${MAIN}" xmlns:r="${REL}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:legacyDrawing r:id="rId1"/></x:worksheet>`,
      ),
      'xl/worksheets/_rels/sheet2.xml.rels': strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', `${REL}/vmlDrawing`, '../drawings/odd.vml')}${rel('rId2', `${REL}/comments`, '../comments9.xml')}${rel('rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/odd.xml')}</Relationships>`,
      ),
      'xl/comments9.xml': strToU8(
        `<x:comments xmlns:x="${MAIN}"><x:authors><x:author>Ann</x:author></x:authors><x:commentList><x:comment ref="B2" authorId="0"><x:text><x:t>Old</x:t></x:text></x:comment><x:comment ref="D4" authorId="0"><x:text><x:t>Going</x:t></x:text></x:comment><x:comment ref="C3" authorId="0"><x:text><x:t>Thread</x:t></x:text></x:comment></x:commentList></x:comments>`,
      ),
      // No shape layout, no note shape type, and a box for D4 only.
      'xl/drawings/odd.vml': strToU8(
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1100"><x:ClientData ObjectType="Note"><x:Row>3</x:Row><x:Column>3</x:Column></x:ClientData></v:shape></xml>',
      ),
      'xl/threadedComments/odd.xml': strToU8(
        `<ThreadedComments xmlns="${THREADS}"><threadedComment ref="C3" personId="{P}" id="{T}" xmlns:q="urn:q"><text>Only</text></threadedComment></ThreadedComments>`,
      ),
      // A drawing named that isn't there.
      'xl/worksheets/sheet3.xml': strToU8(`<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetData/><legacyDrawing r:id="rId7"/></worksheet>`),
    });

  it('are written to all the same', () => {
    const s = WorkbookSession.open(file(), 'odd.xlsx');
    comment(s, 'Bare', 'A1', { action: 'setNote', text: 'Bare' });
    comment(s, 'Odd', 'A1', { action: 'setNote', text: 'First' });
    comment(s, 'Odd', 'D4', { action: 'delete' });
    comment(s, 'Odd', 'C3', { action: 'editEntry', id: '{T}', text: 'Edited' });
    comment(s, 'Broken', 'A1', { action: 'startThread', text: 'Hello', author: 'Ann' });
    const xml = parts(s.save());

    expect(xml['xl/worksheets/sheet1.xml']).toBe(
      `<x:worksheet xmlns:x="${MAIN}" xmlns:r="${REL}"><x:sheetData/><x:legacyDrawing r:id="rId2"/><x:extLst><x:ext uri="{X}"/></x:extLst></x:worksheet>`,
    );
    expect(xml['xl/comments9.xml']).toMatch(/^<x:comments [^>]+><x:authors><x:author>Ann<\/x:author><x:author>tc=\{T\}<\/x:author><x:author><\/x:author><\/x:authors><x:commentList><x:comment ref="A1" authorId="2"/);
    expect(xml['xl/comments9.xml']).toContain('<x:comment ref="B2" authorId="0"><x:text><x:t>Old</x:t></x:text></x:comment>');
    expect(xml['xl/comments9.xml']).not.toContain('D4');
    // The thread's placeholder has no xr:uid, as the part doesn't declare its namespace.
    expect(xml['xl/comments9.xml']).toContain('<x:comment ref="C3" authorId="1" shapeId="0"><x:text><x:t xml:space="preserve">[Threaded comment]');
    const vml = xml['xl/drawings/odd.vml'];
    expect(vml).not.toContain('<x:Row>3</x:Row>');
    expect(vml).toMatch(/^<xml [^>]+><o:shapelayout v:ext="edit" xmlns:o="urn:schemas-microsoft-com:office:office"><o:idmap v:ext="edit" data="2"\/><\/o:shapelayout>/);
    expect(vml).toContain('<v:shapetype id="_x0000_t202" xmlns:o="urn:schemas-microsoft-com:office:office"');
    expect(vml).toContain('<v:shape id="_x0000_s2049" type="#_x0000_t202" xmlns:o="urn:schemas-microsoft-com:office:office"');
    expect(xml['xl/threadedComments/odd.xml']).toBe(`<ThreadedComments xmlns="${THREADS}"><threadedComment ref="C3" personId="{P}" id="{T}" xmlns:q="urn:q"><text>Edited</text></threadedComment></ThreadedComments>`);
    expect(xml['xl/worksheets/sheet3.xml']).toMatch(/^<worksheet [^>]+><sheetData\/><legacyDrawing r:id="rId\d"\/><\/worksheet>$/);
    expect(xml['xl/persons/person.xml']).toMatch(/^<personList xmlns="[^"]+"><person displayName="Ann" id="\{[^"]+\}" userId="Ann" providerId="None"\/><\/personList>$/);

    const back = WorkbookSession.open(s.save(), 'odd.xlsx');
    expect(back.model.sheets.map((m) => m.info.notes.map((n) => [toA1(n.row, n.col), n.kind, n.entries[0].text]))).toEqual([
      [['A1', 'note', 'Bare']],
      [
        ['C3', 'thread', 'Edited'],
        ['A1', 'note', 'First'],
        ['B2', 'note', 'Old'],
      ],
      [['A1', 'thread', 'Hello']],
    ]);
  });

  it('lose a threads part that has nothing left in it', () => {
    const s = WorkbookSession.open(file(), 'odd.xlsx');
    comment(s, 'Odd', 'C3', { action: 'delete' });
    const xml = parts(s.save());
    expect(xml['xl/threadedComments/odd.xml']).toBeUndefined();
    expect(xml['xl/worksheets/_rels/sheet2.xml.rels']).not.toContain('threadedComment');
  });
});

describe('threaded comments', () => {
  it('take replies with ids, dates and people as Excel writes them', () => {
    const s = corpus('comments.xlsx');
    const r = comment(s, 'Review', 'C3', { action: 'reply', text: 'Agreed & <noted>.', author: 'Ann Analyst' });
    const reply = r.noteRedo![0].note!.entries[2];
    expect(reply).toMatchObject({ author: 'Ann Analyst', text: 'Agreed & <noted>.', personId: ANN });
    expect(reply.id).toMatch(GUID);
    expect(reply.time).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d\d$/);
    expect(reply.date).toBe(reply.time!.slice(0, 10));

    const xml = parts(s.save());
    const before = parts(corpusBytes('comments.xlsx'));
    const threads = xml['xl/threadedComments/threadedComment1.xml'];
    // The comments already there are kept as written; the reply follows them.
    const kept = before['xl/threadedComments/threadedComment1.xml'];
    expect(threads).toContain(kept.slice(kept.indexOf('<threadedComment '), kept.indexOf('<threadedComment ref="A6"')));
    expect(threads).toContain(`<threadedComment ref="C3" dT="${reply.time}" personId="${ANN}" id="${reply.id}" parentId="${C3_ROOT}"><text>Agreed &amp; &lt;noted&gt;.</text></threadedComment><threadedComment ref="A6"`);
    expect(xml['xl/comments1.xml']).toContain('Comment:\n    Does this include the Q3 adjustment?\nReply:\n    Yes, since Monday.\nReply:\n    Agreed &amp; &lt;noted&gt;.</t>');
    expect(xml['xl/comments1.xml']).toContain(`<comment ref="C3" authorId="1" shapeId="0" xr:uid="${C3_ROOT}">`);
    // Ann is already one of the people.
    expect(xml['xl/persons/person.xml']).toBe(before['xl/persons/person.xml']);
    expect(noteOn(reopened(s), 'Review', 'C3')?.entries.map((e) => [e.author, e.text])).toEqual([
      ['Ann Analyst', 'Does this include the Q3 adjustment?'],
      ['Bob Reviewer', 'Yes, since Monday.'],
      ['Ann Analyst', 'Agreed & <noted>.'],
    ]);
  });

  it('have replies edited and deleted, and go when their first comment is deleted', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Review', 'C3', { action: 'editEntry', id: C3_REPLY, text: 'Yes, since Tuesday.' });
    let xml = parts(s.save());
    expect(xml['xl/threadedComments/threadedComment1.xml']).toContain(
      `<threadedComment ref="C3" dT="2026-09-14T11:40:02.17" personId="{C2D4E6F8-1A3B-4C5D-8E9F-0A1B2C3D4E5F}" id="${C3_REPLY}" parentId="${C3_ROOT}"><text>Yes, since Tuesday.</text></threadedComment>`,
    );
    comment(s, 'Review', 'C3', { action: 'deleteEntry', id: C3_REPLY });
    comment(s, 'Review', 'A6', { action: 'deleteEntry', id: A6_ROOT });
    xml = parts(s.save());
    expect(xml['xl/threadedComments/threadedComment1.xml']).not.toContain(C3_REPLY);
    expect(xml['xl/threadedComments/threadedComment1.xml']).not.toContain('ref="A6"');
    expect(xml['xl/comments1.xml']).not.toContain('ref="A6"');
    expect(xml['xl/drawings/vmlDrawing1.vml']).not.toContain('<x:Row>5</x:Row>');
    const back = reopened(s);
    expect(noteOn(back, 'Review', 'C3')?.entries.map((e) => e.text)).toEqual(['Does this include the Q3 adjustment?']);
    expect(noteOn(back, 'Review', 'A6')).toBeUndefined();
  });

  it('are resolved and reopened on their first comment', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Review', 'C3', { action: 'resolve', resolved: true });
    comment(s, 'Review', 'A6', { action: 'resolve', resolved: false });
    const threads = parts(s.save())['xl/threadedComments/threadedComment1.xml'];
    expect(threads).toContain(`id="${C3_ROOT}" done="1"><text>`);
    expect(threads).toContain(`parentId="${C3_ROOT}"><text>`);
    expect(threads).toContain(`id="${A6_ROOT}"><text>`);
    const back = reopened(s);
    expect([noteOn(back, 'Review', 'C3')?.resolved, noteOn(back, 'Review', 'A6')?.resolved]).toEqual([true, false]);
  });

  it('are started on a sheet with none, with the writer added to the people', () => {
    const s = corpus('comments.xlsx');
    const r = comment(s, 'Plain', 'A1', { action: 'startThread', text: 'Why two?', author: 'Cara' });
    const first = r.noteRedo![0].note!.entries[0];
    expect(first.personId).toMatch(GUID);
    const xml = parts(s.save());
    expect(xml['xl/worksheets/_rels/sheet2.xml.rels']).toContain('2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment2.xml"/>');
    expect(xml['[Content_Types].xml']).toContain('<Override PartName="/xl/threadedComments/threadedComment2.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>');
    expect(xml['xl/threadedComments/threadedComment2.xml']).toContain(`<threadedComment ref="A1" dT="${first.time}" personId="${first.personId}" id="${first.id}"><text>Why two?</text></threadedComment></ThreadedComments>`);
    expect(xml['xl/persons/person.xml']).toContain(`<person displayName="Cara" id="${first.personId}" userId="Cara" providerId="None"/></personList>`);
    expect(xml['xl/comments3.xml']).toContain(`<author>tc=${first.id}</author>`);
    expect(xml['xl/comments3.xml']).toContain(`xr:uid="${first.id}"`);
    // A second comment by the same person names them again.
    comment(s, 'Plain', 'A1', { action: 'reply', text: 'Typo.', author: 'Cara' });
    expect(s.noteAt(1, 0, 0)?.entries[1].personId).toBe(first.personId);
    expect(noteOn(reopened(s), 'Plain', 'A1')).toMatchObject({ kind: 'thread', resolved: false, entries: [{ author: 'Cara', text: 'Why two?' }, { author: 'Cara', text: 'Typo.' }] });
  });

  it('are started in a workbook without people, which then gets its list of them', () => {
    const s = corpus('templates.xlsx');
    comment(s, 'Arrays', 'A1', { action: 'startThread', text: 'Sorted?', author: 'Dee' });
    const xml = parts(s.save());
    expect(xml['xl/_rels/workbook.xml.rels']).toContain('2017/10/relationships/person" Target="persons/person.xml"/>');
    expect(xml['[Content_Types].xml']).toContain('<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>');
    expect(xml['xl/persons/person.xml']).toMatch(/^<\?xml[^>]*\?>\r?\n<personList xmlns="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2018\/threadedcomments" xmlns:x="[^"]+"><person displayName="Dee" id="\{[^"]+\}" userId="Dee" providerId="None"\/><\/personList>$/);
    expect(noteOn(reopened(s), 'Arrays', 'A1')?.entries[0]).toMatchObject({ author: 'Dee', text: 'Sorted?' });
  });

  it("aren't changed in ways that don't apply", () => {
    const s = corpus('comments.xlsx');
    const refused = (ref: string, change: CommentChange) => {
      const r = comment(s, 'Review', ref, change);
      expect(r.noteRedo).toBeUndefined();
      return english.text(r.rejected);
    };
    expect(refused('B2', { action: 'reply', text: 'Hi', author: 'Ann Analyst' })).toBe('This cell has no comment thread.');
    expect(refused('A6', { action: 'reply', text: 'Hi', author: 'Ann Analyst' })).toBe('This thread is resolved. Reopen it to reply.');
    expect(refused('B2', { action: 'startThread', text: 'Hi', author: 'Ann Analyst' })).toBe('This cell already has a note or comments.');
    expect(refused('C3', { action: 'setNote', text: 'Hi' })).toBe('This cell has a comment thread, not a note.');
    expect(refused('C3', { action: 'reply', text: '  ', author: 'Ann Analyst' })).toBe('A comment needs some text.');
    expect(refused('C3', { action: 'editEntry', id: '{nope}', text: 'Hi' })).toBe('That comment is no longer in the thread.');
    expect(refused('C3', { action: 'deleteEntry', id: '{nope}' })).toBe('That comment is no longer in the thread.');
    expect(s.editedCount).toBe(0);
    expect(s.save()).toEqual(corpusBytes('comments.xlsx'));
  });
});

describe('protected sheets and merged cells', () => {
  it("keep their notes as they are unless their protection lets objects be edited, as Excel's does", () => {
    const s = corpus('comments.xlsx');
    expect(s.model.sheets.map((m) => [m.info.name, m.info.protected, !!m.info.objectsLocked])).toEqual([
      ['Review', false, false],
      ['Plain', false, false],
      ['Locked', true, true],
      ['Objects', true, false],
    ]);
    expect(english.text(comment(s, 'Locked', 'A1', { action: 'setNote', text: 'Mine now' }).rejected)).toBe(
      'This sheet is protected in Excel, which keeps its notes and comments as they are.',
    );
    // Locked cells or not: notes are objects.
    expect(comment(s, 'Objects', 'A1', { action: 'setNote', text: 'Allowed' }).noteRedo).toHaveLength(1);
  });

  it('take a note on the first cell of a merge only', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    expect(english.text(comment(s, 'Summary', 'C1', { action: 'setNote', text: 'x' }).rejected)).toBe('Part of a merged cell. Edit its top-left cell.');
    expect(comment(s, 'Summary', 'A1', { action: 'setNote', text: 'x' }).rejected).toBeUndefined();
  });

  it("can't be changed on a chart sheet", () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const chart = s.model.sheets.findIndex((m) => m.info.kind === 'chartsheet');
    expect(english.text(s.apply({ type: 'comment', sheet: chart, row: 0, col: 0, change: { action: 'setNote', text: 'x' } }).rejected)).toBe('This sheet cannot be edited here.');
  });
});

describe('undo, the change list and structural changes', () => {
  it('undo and redo put notes and threads back exactly', () => {
    const s = corpus('comments.xlsx');
    const reply = comment(s, 'Review', 'C3', { action: 'reply', text: 'Done.', author: 'Ann Analyst' });
    const added = s.noteAt(0, 2, 2);
    const undone = s.apply({ type: 'restore', cells: reply.undo, notes: reply.noteUndo });
    expect(undone.noteRedo).toEqual(reply.noteUndo);
    expect(s.editedCount).toBe(0);
    expect(s.save()).toEqual(corpusBytes('comments.xlsx'));
    s.apply({ type: 'restore', cells: reply.redo, notes: reply.noteRedo });
    expect(s.noteAt(0, 2, 2)).toEqual(added);
    // Restoring the same state again changes nothing.
    expect(s.apply({ type: 'restore', cells: [], notes: reply.noteRedo }).noteRedo).toEqual([]);
  });

  it('a batch undoes its changes to one cell in reverse', () => {
    const s = corpus('comments.xlsx');
    const r = s.apply({
      type: 'batch',
      ops: [
        { type: 'comment', sheet: 1, row: 0, col: 0, change: { action: 'setNote', text: 'One' } },
        { type: 'comment', sheet: 1, row: 0, col: 0, change: { action: 'setNote', text: 'Two' } },
      ],
    });
    expect(r.noteRedo?.map((n) => n.note?.entries[0].text)).toEqual(['One', 'Two']);
    expect(r.noteUndo?.map((n) => n.note?.entries[0].text ?? null)).toEqual(['One', null]);
    s.apply({ type: 'restore', cells: [], notes: r.noteUndo });
    expect(s.noteAt(1, 0, 0)).toBeNull();
  });

  it('a restored thread by someone new to the file adds them to its people', () => {
    const s = corpus('templates.xlsx');
    const note: CellNote = { row: 0, col: 0, kind: 'thread', resolved: false, entries: [{ author: 'Eve', text: 'Hi', id: newGuid(), personId: newGuid(), time: commentTime(new Date(Date.UTC(2026, 0, 2, 3, 4, 5, 670))) }] };
    s.apply({ type: 'restore', cells: [], notes: [{ sheet: 2, row: 0, col: 0, note }] });
    const xml = parts(s.save());
    expect(xml['xl/threadedComments/threadedComment1.xml']).toContain('dT="2026-01-02T03:04:05.67"');
    expect(xml['xl/persons/person.xml']).toContain(`<person displayName="Eve" id="${note.entries[0].personId}" userId="Eve" providerId="None"/>`);
  });

  it('are listed with the changed cells and counted with them', () => {
    const s = corpus('comments.xlsx');
    s.apply({ type: 'input', sheet: 0, row: 1, col: 1, text: '1250' });
    comment(s, 'Review', 'B2', { action: 'setNote', text: 'Raised.' });
    comment(s, 'Review', 'C3', { action: 'reply', text: 'OK', author: 'Cara' });
    comment(s, 'Review', 'C3', { action: 'resolve', resolved: true });
    comment(s, 'Review', 'A6', { action: 'delete' });
    comment(s, 'Plain', 'B1', { action: 'startThread', text: 'Units?', author: 'Cara' });
    comment(s, 'Plain', 'A2', { action: 'setNote', text: 'Gone soon' });
    comment(s, 'Plain', 'A2', { action: 'delete' });
    expect(s.editedCount).toBe(5);
    expect(s.changes().map((c) => [`${c.sheet}!${c.ref}`, c.before, c.after, c.comment && english.text(c.comment)])).toEqual([
      ['Review!B2', '1200', '1250', undefined],
      ['Review!B2', 'Agreed with finance.\nRevisit in January.', 'Raised.', 'Note edited'],
      [
        'Review!C3',
        'Ann Analyst: Does this include the Q3 adjustment?\nBob Reviewer: Yes, since Monday.',
        'Ann Analyst: Does this include the Q3 adjustment?\nBob Reviewer: Yes, since Monday.\nCara: OK',
        'Comment thread: 1 reply added, resolved',
      ],
      ['Review!A6', 'Bob Reviewer: Rename this row?', '', 'Comment thread deleted'],
      ['Plain!B1', '', 'Cara: Units?', 'Comment thread started'],
    ]);
  });

  it('follow their cells when rows are inserted, against what the file had', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Review', 'B2', { action: 'setNote', text: 'Raised.' });
    const next = s.restructure({ type: 'insertRows', sheet: 0, at: 0, count: 1 });
    expect(next.noteAt(0, 2, 1)?.entries[0].text).toBe('Raised.');
    expect(next.changes().map((c) => [c.ref, c.before, c.after, c.note ?? c.comment])).toEqual([
      ['1:1', '', '', { key: 'core.change.insertedRows', params: { count: 1 } }],
      ['B3', 'Agreed with finance.\nRevisit in January.', 'Raised.', { key: 'core.change.noteEdited' }],
    ]);
    // Put back as the file had it, the note is no longer a change.
    next.apply({ type: 'comment', sheet: 0, row: 2, col: 1, change: { action: 'setNote', text: 'Agreed with finance.\nRevisit in January.' } });
    expect(next.changes().map((c) => c.ref)).toEqual(['1:1']);
    expect(next.editedCount).toBe(1);
  });

  it('are gone with the rows they were on', () => {
    const s = corpus('comments.xlsx');
    comment(s, 'Review', 'B2', { action: 'setNote', text: 'Raised.' });
    const next = s.restructure({ type: 'deleteRows', sheet: 0, at: 1, count: 1 });
    expect(next.changes().map((c) => c.ref)).toEqual(['2:2']);
  });
});

describe('describing changes', () => {
  const thread = (texts: string[], resolved = false): CellNote => ({
    row: 0,
    col: 0,
    kind: 'thread',
    resolved,
    entries: texts.map((text, i) => ({ text, id: `t${i}`, author: i ? undefined : 'Ann' })),
  });

  it('say what happened to a thread', () => {
    const say = (a: CellNote | null, b: CellNote | null) => english.text(describeCommentChange(a, b));
    expect(say(thread(['a', 'b', 'c']), { ...thread(['a', 'B']), resolved: true })).toBe('Comment thread: 1 comment edited, 1 reply deleted, resolved');
    expect(say(thread(['a'], true), thread(['a']))).toBe('Comment thread: reopened');
    expect(say(thread(['a']), { ...thread(['a']), entries: [{ ...thread(['a']).entries[0], time: 'later' }] })).toBe('Comment thread edited');
    expect(say({ row: 0, col: 0, kind: 'note', entries: [{ text: 'x' }] }, thread(['a']))).toBe('Comment thread edited');
    expect(say(null, { row: 0, col: 0, kind: 'note', entries: [{ text: 'x' }] })).toBe('Note added');
    expect(say({ row: 0, col: 0, kind: 'note', entries: [{ text: 'x' }] }, null)).toBe('Note deleted');
    expect(noteText(thread(['a', 'b']))).toBe('Ann: a\nb');
    expect(noteText(null)).toBe('');
  });

  it('keep a note with no entries, and stamp new comments with the given id and time', () => {
    const ctx = { person: () => '{P}', newId: () => '{ID}', now: () => '2026-01-02T03:04:05.06' };
    expect(applyCommentChange({ row: 0, col: 0, kind: 'note', entries: [] }, 0, 0, { action: 'setNote', text: 'x', author: 'Ann' }, ctx)?.entries).toEqual([
      { author: undefined, text: 'x' },
    ]);
    expect(applyCommentChange(null, 0, 0, { action: 'startThread', text: 'x', author: 'Ann' }, ctx)?.entries).toEqual([
      { author: 'Ann', text: 'x', date: '2026-01-02', id: '{ID}', personId: '{P}', time: '2026-01-02T03:04:05.06' },
    ]);
    expect(newGuid()).toMatch(GUID);
  });
});
