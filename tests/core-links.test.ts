import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import { linkAddress } from '../src/core/links';
import { WorkbookSession } from '../src/core/session';
import type { Hyperlink, LayoutOp } from '../src/core/types';
import { english } from '../src/i18n';
import { parts } from './helpers';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * "Plain": A1:B3 filled, no links. "Links": A1 links to a web page, A2 to a
 * place in the workbook, with conditional formatting, validation and a
 * picture-free drawing-free tail. "Locked" is protected.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const plain = wb.addWorksheet('Plain');
  for (let r = 1; r <= 3; r++) plain.addRow([`a${r}`, r]);
  const links = wb.addWorksheet('Links');
  links.getCell('A1').value = { text: 'Example', hyperlink: 'https://example.com/a', tooltip: 'Go there' };
  links.getCell('A2').value = { text: 'Back', hyperlink: '#Plain!B2' };
  links.getCell('A1').font = { color: { theme: 10 }, underline: true };
  links.getCell('A3').value = 'Not a link';
  links.getCell('B1').value = { formula: 'HYPERLINK("https://example.com/b","B")', result: 'B' };
  links.getCell('B2').value = 'https://example.com/from-cell';
  links.getCell('B3').value = { formula: 'HYPERLINK(B2,"From B2")', result: 'From B2' };
  links.addConditionalFormatting({ ref: 'C1:C5', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['5'], priority: 1, style: { font: { bold: true } } }] });
  links.getCell('D1').dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0], allowBlank: true };
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const open = (b = bytes, recalc: 'on' | 'off' = 'on') => WorkbookSession.open(b, 'links.xlsx', 'en-US', { recalc });
const links = (s: WorkbookSession, sheet: number, op: Omit<Extract<LayoutOp, { kind: 'links' }>, 'kind'>) => s.apply({ type: 'layout', sheet, op: { kind: 'links', ...op } });
const reopen = (s: WorkbookSession) => open(s.save());
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const sheetXml = (b: Uint8Array, n: number) => parts(b)[`xl/worksheets/sheet${n}.xml`];
const sheetRels = (b: Uint8Array, n: number) => parts(b)[`xl/worksheets/_rels/sheet${n}.xml.rels`];

describe('reading links', () => {
  it('reads links to addresses through the sheet’s relationships, and places from their location', () => {
    const info = open().model.sheets[1].info;
    expect(info.links).toEqual([
      { rect: rect(0, 0), id: 0, target: 'https://example.com/a', tooltip: 'Go there' },
      // ExcelJS writes a place as an address starting "#", and as the location too.
      { rect: rect(1, 0), id: 1, target: '#Plain!B2', location: '#Plain!B2' },
    ]);
    expect(info.links.map(linkAddress)).toEqual(['https://example.com/a', 'Plain!B2']);
    expect(open().model.findings.find((f) => f.kind === 'hyperlink')).toMatchObject({ sheet: 'Links', detail: { key: 'links.finding.detail', params: { count: 2 } } });
  });

  it('reads the corpus’s links and their display text', () => {
    const s = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))), 'templates.xlsx');
    expect(s.model.sheets.find((m) => m.info.name === 'Lines')!.info.links).toEqual([
      { rect: rect(0, 6), id: 0, target: 'https://example.com/catalogue' },
      { rect: rect(1, 6), id: 1, location: 'Request!B3', display: 'Back to the request' },
    ]);
  });

  it("marks fonts in the theme's hyperlink colour, as Excel's Hyperlink style has", () => {
    const s = open();
    const grid = s.model.sheets[1].scan!.grid;
    expect(s.model.styles.get(grid.styles[grid.rowPtr[0]])).toMatchObject({ linkFont: true, underline: true, color: '#0000FF' });
    expect(s.model.styles.get(grid.styles[grid.rowPtr[2]]).linkFont).toBeUndefined();
  });

  it('passes the defined names on, for links that go to one', () => {
    expect(open().loadResult().names).toEqual([]);
  });
});

describe('saving links', () => {
  it('adds <hyperlinks> to a sheet that had none, in its place, with a relationship for an address', () => {
    const s = open();
    links(s, 0, { add: [{ rect: rect(0, 0), target: 'https://example.com/new page', tooltip: 'A "tip"\non two lines', display: 'New' }, { rect: rect(1, 0, 2, 1), location: "'Links'!A1" }], remove: [] });
    expect(s.editedCount).toBe(2);
    const saved = s.save();
    const xml = sheetXml(saved, 1);
    expect(xml).toMatch(/<\/sheetData>.*<hyperlinks><hyperlink ref="A1" r:id="rId\d+" tooltip="A &quot;tip&quot;&#10;on two lines" display="New"\/><hyperlink ref="A2:B3" location="'Links'!A1"\/><\/hyperlinks><pageMargins/);
    const id = /<hyperlink ref="A1" r:id="(rId\d+)"/.exec(xml)![1];
    expect(sheetRels(saved, 1)).toContain(`<Relationship Id="${id}" Type="${REL}/hyperlink" Target="https://example.com/new%20page" TargetMode="External"/>`);
    expect(reopen(s).model.sheets[0].info.links).toEqual([
      { rect: rect(0, 0), id: 0, target: 'https://example.com/new%20page', tooltip: 'A "tip"\non two lines', display: 'New' },
      { rect: rect(1, 0, 2, 1), id: 1, location: "'Links'!A1" },
    ]);
  });

  it('goes after conditional formatting and validation, and keeps unchanged links as they were', () => {
    const s = open();
    const [ext, place] = s.model.sheets[1].info.links;
    links(s, 1, { add: [{ rect: rect(4, 0), target: 'mailto:someone@example.com' }], remove: [] });
    const xml = sheetXml(s.save(), 2);
    const before = sheetXml(bytes, 2);
    const box = (x: string) => /<hyperlinks>.*<\/hyperlinks>/.exec(x)![0];
    expect(box(xml).startsWith(box(before).replace('</hyperlinks>', ''))).toBe(true);
    expect(xml.indexOf('<dataValidations')).toBeLessThan(xml.indexOf('<hyperlinks>'));
    expect(xml.indexOf('<conditionalFormatting')).toBeLessThan(xml.indexOf('<hyperlinks>'));
    expect(reopen(s).model.sheets[1].info.links).toEqual([ext, place, { rect: rect(4, 0), id: 2, target: 'mailto:someone@example.com' }]);
  });

  it('changes a link: the old relationship goes and a new one comes', () => {
    const s = open();
    const [ext] = s.model.sheets[1].info.links;
    links(s, 1, { add: [{ ...ext, id: undefined, target: 'https://example.com/changed' }], remove: [ext] });
    const saved = s.save();
    const rels = sheetRels(saved, 2);
    expect(rels).not.toContain('https://example.com/a"');
    expect(rels).toContain('https://example.com/changed');
    expect(open(saved).model.sheets[1].info.links.map((l) => linkAddress(l))).toEqual(['Plain!B2', 'https://example.com/changed']);
    expect(s.changes().filter((c) => c.note).map((c) => [c.ref, english.text(c.note!)])).toEqual([['A1', 'Link changed to https://example.com/changed']]);
  });

  it('removes <hyperlinks> with the last link, and the relationships that went with them', () => {
    const s = open();
    const info = s.model.sheets[1].info;
    links(s, 1, { add: [], remove: info.links });
    expect(s.changes().filter((c) => c.note).map((c) => english.text(c.note!))).toEqual(['Link removed', 'Link removed']);
    const saved = s.save();
    expect(sheetXml(saved, 2)).not.toContain('hyperlink');
    expect(sheetRels(saved, 2) ?? '').not.toContain('/hyperlink"');
    expect(open(saved).model.sheets[1].info.links).toEqual([]);
  });

  it('undoes as one step, and saves the file as it was once undone', () => {
    const s = open();
    const r = links(s, 1, { add: [{ rect: rect(9, 9), target: 'https://example.com/x' }], remove: [s.model.sheets[1].info.links[0]] });
    expect(r.layout?.undo).toEqual([{ sheet: 1, op: { kind: 'links', add: [s.model.sheets[1].info.links[0]], remove: [{ rect: rect(9, 9), target: 'https://example.com/x' }] } }]);
    s.apply({ type: 'restore', cells: r.undo, layout: r.layout!.undo });
    expect(s.editedCount).toBe(0);
    expect(s.save()).toBe(bytes);
  });

  it('declares the relationships namespace, and makes the relationships part, for a sheet that had neither', () => {
    const files = unzipSync(bytes);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(files['xl/worksheets/sheet1.xml'])
        .replace(` xmlns:r="${REL}"`, '')
        .replace('</sheetData>', '</sheetData><!-- a comment --><?pi here?>'),
    );
    delete files['xl/worksheets/_rels/sheet1.xml.rels'];
    const bare = zipSync(files);
    expect(sheetXml(bare, 1)).not.toContain('xmlns:r=');
    const s = open(bare);
    links(s, 0, { add: [{ rect: rect(0, 1), target: 'https://example.com/' }], remove: [] });
    const saved = s.save();
    expect(sheetXml(saved, 1)).toMatch(new RegExp(`^<\\?xml[^>]*\\?>\\s*<worksheet [^>]* xmlns:r="${REL}">`));
    expect(sheetRels(saved, 1)).toMatch(/^<\?xml[^>]*\?>\s*<Relationships xmlns="[^"]+"><Relationship Id="rId1" [^>]*TargetMode="External"\/><\/Relationships>$/);
    expect(open(saved).model.sheets[0].info.links[0].target).toBe('https://example.com/');
  });

  it('fills a relationships part that closes itself', () => {
    const files = unzipSync(bytes);
    files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(`<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
    const s = open(zipSync(files));
    links(s, 0, { add: [{ rect: rect(0, 1), target: 'https://example.com/' }], remove: [] });
    expect(sheetRels(s.save(), 1)).toMatch(/<Relationships [^>]*><Relationship Id="rId1" [^>]*\/><\/Relationships>$/);
  });

  it("can't change a protected sheet's links", () => {
    const s = open();
    const r = links(s, 2, { add: [{ rect: rect(0, 0), target: 'https://example.com/' }], remove: [] });
    expect(r.rejected).toBeDefined();
    expect(s.editedCount).toBe(0);
  });
});

describe('links and cells', () => {
  it("stay when a cell's contents are cleared, as Excel's Clear Contents leaves them", () => {
    const s = open();
    s.apply({ type: 'clear', sheet: 1, rects: [rect(0, 0, 1, 0)] });
    expect(reopen(s).model.sheets[1].info.links).toHaveLength(2);
  });

  it('are copied with their cells, replacing the links where they are pasted', () => {
    const s = open();
    const r = s.apply({ type: 'copy', srcSheet: 1, src: rect(0, 0, 1, 0), sheet: 1, dst: rect(1, 0, 2, 0) });
    expect(r.layout?.applied[0].op).toMatchObject({ kind: 'links' });
    const after = reopen(s).model.sheets[1].info.links.map((l) => [l.rect.r1, linkAddress(l)]);
    expect(after).toEqual([[0, 'https://example.com/a'], [1, 'https://example.com/a'], [2, 'Plain!B2']]);
    // Undoing the paste puts the links back too.
    s.apply({ type: 'restore', cells: r.undo, layout: r.layout!.undo });
    expect(reopen(s).model.sheets[1].info.links.map((l) => l.rect.r1)).toEqual([0, 1]);
  });

  it('are copied to another sheet', () => {
    const s = open();
    s.apply({ type: 'copy', srcSheet: 1, src: rect(0, 0), sheet: 0, dst: rect(5, 5) });
    expect(reopen(s).model.sheets[0].info.links).toMatchObject([{ rect: rect(5, 5), target: 'https://example.com/a', tooltip: 'Go there' }]);
  });

  it('move with their rows when the rows are sorted', () => {
    const s = open();
    s.apply({ type: 'sort', sheet: 1, rect: rect(0, 0, 2, 0), keys: [{ col: 0, descending: false }], header: false });
    // "Back", "Example", "Not a link".
    const info = reopen(s).model.sheets[1].info;
    expect(info.links.map((l) => [l.rect.r1, linkAddress(l)]).sort()).toEqual([[0, 'Plain!B2'], [1, 'https://example.com/a']]);
  });

  it('move with cut cells on their sheet', () => {
    const s = open();
    const moved = s.restructure({ type: 'move', sheet: 1, src: rect(0, 0), toSheet: 1, row: 9, col: 3 });
    expect(moved.model.sheets[1].info.links.map((l) => [l.rect.r1, l.rect.c1, linkAddress(l)])).toEqual([
      [9, 3, 'https://example.com/a'],
      [1, 0, 'Plain!B2'],
    ]);
  });

  it('follow their cells when rows are inserted, links added here included', () => {
    const s = open();
    links(s, 1, { add: [{ rect: rect(5, 1), location: 'Plain!A1' }], remove: [] });
    const next = s.restructure({ type: 'insertRows', sheet: 1, at: 1, count: 2 });
    expect(next.model.sheets[1].info.links.map((l) => [l.rect.r1, l.rect.c1, linkAddress(l)])).toEqual([
      [0, 0, 'https://example.com/a'],
      [3, 0, 'Plain!B2'],
      [7, 1, 'Plain!A1'],
    ]);
    expect(next.changes().map((c) => english.text(c.note ?? ''))).toContain('Link to Plain!A1 added');
    // Deleting a row takes its link away, and the relationship goes when it is saved again.
    const gone = next.restructure({ type: 'deleteRows', sheet: 1, at: 0, count: 1 });
    expect(gone.model.sheets[1].info.links.map((l) => l.rect.r1)).toEqual([2, 6]);
  });
});

describe('HYPERLINK formulas', () => {
  const at = (u: { sheet: number; row: number; col: number }[], row: number, col: number) => u.find((x) => x.sheet === 1 && x.row === row && x.col === col) as { link?: string } | undefined;

  it('link to the address the engine works out, and say so when it changes', () => {
    const s = open();
    const report = s.ensureEngine();
    expect(at(report.updates, 0, 1)?.link).toBe('https://example.com/b');
    expect(at(report.updates, 2, 1)?.link).toBe('https://example.com/from-cell');
    // The address changes while the text shown doesn't.
    const r = s.apply({ type: 'input', sheet: 1, row: 1, col: 1, text: 'https://example.com/changed' });
    expect(at(r.updates, 2, 1)?.link).toBe('https://example.com/changed');
    // A cell that stops being a link stops being told.
    const plain = s.apply({ type: 'input', sheet: 1, row: 2, col: 1, text: '5' });
    expect(at(plain.updates, 2, 1)?.link).toBeUndefined();
    expect(at(s.apply({ type: 'input', sheet: 1, row: 1, col: 1, text: 'x' }).updates, 2, 1)).toBeUndefined();
  });

  it('typed in, link at once', () => {
    const s = open();
    s.ensureEngine();
    const r = s.apply({ type: 'input', sheet: 0, row: 5, col: 0, text: '=HYPERLINK("#Links!A1","Go")' });
    expect(r.updates.find((u) => u.row === 5)).toMatchObject({ link: '#Links!A1', s: 'Go' });
  });

  it('link to what the formula writes when recalculation is off', () => {
    const s = open(bytes, 'off');
    s.ensureEngine();
    const r = s.apply({ type: 'input', sheet: 0, row: 5, col: 0, text: '=HYPERLINK("https://example.com/off")' });
    expect(r.updates.find((u) => u.row === 5)?.link).toBe('https://example.com/off');
    expect(s.apply({ type: 'input', sheet: 0, row: 6, col: 0, text: '=HYPERLINK(A1)' }).updates.find((u) => u.row === 6)?.link).toBeUndefined();
  });
});

describe('the list of changes', () => {
  it('lists links added and removed by range', () => {
    const s = open();
    const added: Hyperlink = { rect: rect(3, 3, 4, 4), target: 'https://example.com/d' };
    links(s, 1, { add: [added], remove: [s.model.sheets[1].info.links[1]] });
    expect(s.changes().map((c) => [c.sheet, c.ref, english.text(c.note ?? '')])).toEqual([
      ['Links', 'A2', 'Link removed'],
      ['Links', 'D4:E5', 'Link to https://example.com/d added'],
    ]);
  });
});
