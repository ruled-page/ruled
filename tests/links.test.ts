import { describe, expect, it } from 'vitest';
import {
  applyLinkChange,
  checkAddress,
  copyLinks,
  isHyperlinkFormula,
  linkAddress,
  linkDestination,
  linkDifferences,
  literalHyperlinkTarget,
  replaceLinks,
  resolvePlace,
  sortLinks,
  splitPlace,
  typedAddress,
} from '../src/core/links';
import type { Hyperlink } from '../src/core/types';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

describe('which addresses may be opened', () => {
  it('opens web pages and email', () => {
    expect(checkAddress('https://example.com/a?b=1#c')).toEqual({ kind: 'web', url: 'https://example.com/a?b=1#c' });
    expect(checkAddress('HTTP://Example.COM')).toEqual({ kind: 'web', url: 'http://example.com/' });
    expect(checkAddress('mailto:someone@example.com?subject=Hi')).toEqual({ kind: 'email', url: 'mailto:someone@example.com?subject=Hi' });
  });

  it.each([
    'javascript:alert(1)',
    'JavaScript:alert(1)',
    ' javascript:alert(1)',
    'java\tscript:alert(1)',
    'java\nscript:alert(1)',
    '\u0001javascript:alert(1)',
    'vbscript:msgbox(1)',
    'data:text/html,<script>alert(1)</script>',
    'file:///C:/Windows/System32/calc.exe',
    'file://server/share/x.xlsx',
    '\\\\server\\share\\book.xlsx',
    'C:\\Users\\me\\book.xlsx',
    '[Book2.xlsx]Sheet1!A1',
    '//example.com/path',
    'ms-excel:ofe|u|https://example.com/x.xlsx',
    'http://',
    'blob:https://example.com/123',
    '',
  ])('refuses %j', (address) => {
    expect(checkAddress(address)).toEqual({ kind: 'refused', address });
  });

  it('reads the address as the browser does, so spacing can neither hide nor change a scheme', () => {
    // Browsers drop surrounding spaces and control characters, and tabs and line breaks anywhere.
    expect(checkAddress('  https://exa\nmple.com\t ')).toEqual({ kind: 'web', url: 'https://example.com/' });
    // Lookalike letters in a domain show as what they are.
    expect(checkAddress('https://exаmple.com/').kind).toBe('web');
    expect((checkAddress('https://exаmple.com/') as { url: string }).url).toBe('https://xn--exmple-4nf.com/');
  });

  it("goes to places in the workbook, a link's location, or a HYPERLINK address starting #", () => {
    expect(linkDestination({ location: 'Sheet2!A1' })).toEqual({ kind: 'place', location: 'Sheet2!A1' });
    expect(linkDestination({ target: '#Sheet2!A1' })).toEqual({ kind: 'place', location: 'Sheet2!A1' });
    expect(linkDestination({ target: 'https://example.com/page', location: 'part' })).toEqual({ kind: 'web', url: 'https://example.com/page#part' });
    expect(linkDestination({ target: 'javascript:alert(1)' })).toEqual({ kind: 'refused', address: 'javascript:alert(1)' });
    expect(linkDestination({})).toBeUndefined();
    expect(linkDestination({ target: '#' })).toBeUndefined();
    expect(linkDestination({ target: ' ', location: ' ' })).toBeUndefined();
  });

  it('takes addresses typed in, adding the scheme Excel would', () => {
    expect(typedAddress(' www.example.com/a b ')).toBe('https://www.example.com/a%20b');
    expect(typedAddress('someone@example.com')).toBe('mailto:someone@example.com');
    expect(typedAddress('http://example.com')).toBe('http://example.com/');
    expect(typedAddress('MAILTO:a@b.c?subject=Hi there')).toBe('mailto:a@b.c?subject=Hi%20there');
    for (const bad of ['', '  ', 'javascript:alert(1)', 'file:///etc/passwd', 'ftp://example.com', 'https://']) expect(typedAddress(bad), bad).toBeUndefined();
  });

  it('reads as the address, with any part of the page', () => {
    expect(linkAddress({ target: 'https://example.com', location: 'x' })).toBe('https://example.com#x');
    expect(linkAddress({ target: 'https://example.com' })).toBe('https://example.com');
    expect(linkAddress({ location: 'Sheet2!A1' })).toBe('Sheet2!A1');
    expect(linkAddress({})).toBe('');
  });
});

describe('places in the workbook', () => {
  const sheets = ['Data', 'Q1 \'24', 'Other'];
  const names = [
    { name: 'Total', formula: 'Data!$B$10' },
    { name: 'Total', formula: 'Other!$C$3', localSheet: 2 },
    { name: 'Moving', formula: 'OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)' },
    { name: 'Hop', formula: 'Total' },
  ];

  it('splits a sheet from its cells, as Excel quotes it', () => {
    expect(splitPlace('Data!A1')).toEqual({ sheet: 'Data', ref: 'A1' });
    expect(splitPlace("'Q1 ''24'!B2:C4")).toEqual({ sheet: "Q1 '24", ref: 'B2:C4' });
    expect(splitPlace('B5')).toEqual({ ref: 'B5' });
    for (const bad of ['', "'Data!A1", "'Data'A1", "''!A1", 'Data!', '!A1', 'a!b!c', "'x'!a!b"]) expect(splitPlace(bad), bad).toBeUndefined();
  });

  it('resolve to a sheet and cells, from a reference or a defined name', () => {
    expect(resolvePlace('Other!B2:C4', sheets, names, 0)).toEqual({ sheet: 2, rect: rect(1, 1, 3, 2) });
    expect(resolvePlace("'Q1 ''24'!$A$1", sheets, names, 0)).toEqual({ sheet: 1, rect: rect(0, 0) });
    expect(resolvePlace('data!A3', sheets, names, 2)).toEqual({ sheet: 0, rect: rect(2, 0) });
    // On the link's own sheet when it names none.
    expect(resolvePlace('C7', sheets, names, 2)).toEqual({ sheet: 2, rect: rect(6, 2) });
    expect(resolvePlace('=Data!A1', sheets, names, 2)).toEqual({ sheet: 0, rect: rect(0, 0) });
    // A name local to the link's sheet comes first.
    expect(resolvePlace('total', sheets, names, 0)).toEqual({ sheet: 0, rect: rect(9, 1) });
    expect(resolvePlace('Total', sheets, names, 2)).toEqual({ sheet: 2, rect: rect(2, 2) });
  });

  it("aren't found when the sheet or name isn't there, or the name is worked out by a formula", () => {
    for (const bad of ['Missing!A1', 'Moving', 'Nowhere', 'Hop', 'Data!Total', 'Data!A1:B2:C3', '']) expect(resolvePlace(bad, sheets, names, 0), bad).toBeUndefined();
  });
});

describe('HYPERLINK formulas', () => {
  it('are found by their first call', () => {
    expect(isHyperlinkFormula('HYPERLINK("https://example.com")')).toBe(true);
    expect(isHyperlinkFormula('  hyperlink ( A1 , "x")')).toBe(true);
    expect(isHyperlinkFormula('HYPERLINKS(1)')).toBe(false);
    expect(isHyperlinkFormula('IF(A1,HYPERLINK("x"),"")')).toBe(false);
    expect(isHyperlinkFormula(undefined)).toBe(false);
    expect(isHyperlinkFormula('HYPERLINK')).toBe(false);
  });

  it('give their address when the formula writes it as text', () => {
    expect(literalHyperlinkTarget('HYPERLINK("https://example.com/?q=""x""", "Go")')).toBe('https://example.com/?q="x"');
    expect(literalHyperlinkTarget('HYPERLINK( "#Data!A1" )')).toBe('#Data!A1');
    expect(literalHyperlinkTarget('HYPERLINK(A1,"Go")')).toBeUndefined();
    expect(literalHyperlinkTarget('HYPERLINK("a"&B1)')).toBeUndefined();
    expect(literalHyperlinkTarget('HYPERLINK("unclosed')).toBeUndefined();
    expect(literalHyperlinkTarget('SUM("x")')).toBeUndefined();
  });
});

describe('changing a sheet’s links', () => {
  const big: Hyperlink = { rect: rect(0, 0, 2, 2), target: 'https://example.com', tooltip: 'Tip', id: 0 };
  const other: Hyperlink = { rect: rect(5, 0), location: 'Data!A1', display: 'Back', id: 1 };

  it('gives the selected cells a link, keeping the parts of a link they were in', () => {
    const change = replaceLinks([big, other], [rect(1, 1)], { target: 'https://new.example' });
    expect(change.remove).toEqual([big]);
    expect(change.add).toEqual([
      { rect: rect(0, 0, 0, 2), target: 'https://example.com', tooltip: 'Tip' },
      { rect: rect(2, 0, 2, 2), target: 'https://example.com', tooltip: 'Tip' },
      { rect: rect(1, 0), target: 'https://example.com', tooltip: 'Tip' },
      { rect: rect(1, 2), target: 'https://example.com', tooltip: 'Tip' },
      { rect: rect(1, 1), target: 'https://new.example' },
    ]);
    const { links, added, removed } = applyLinkChange([big, other], change);
    expect(removed).toEqual([big]);
    expect(added).toHaveLength(5);
    expect(links[0]).toBe(other);
  });

  it('takes links away when there is none to give, across several ranges', () => {
    expect(replaceLinks([big, other], [rect(0, 0, 2, 2), rect(5, 0, 6, 1)])).toEqual({ add: [], remove: [big, other] });
    expect(replaceLinks([other], [rect(9, 9)])).toEqual({ add: [], remove: [] });
  });

  it('copies links with their cells, tiled, replacing what was there', () => {
    // A link over the whole source covers the whole destination.
    const one = copyLinks([other], rect(5, 0), [big], rect(1, 1, 3, 1));
    expect(one.remove).toEqual([big]);
    expect(one.add.at(-1)).toEqual({ rect: rect(1, 1, 3, 1), location: 'Data!A1', display: 'Back' });
    // Part of the source: once per tile, cut at the destination's edge.
    const tiled = copyLinks([{ rect: rect(0, 0), target: 'x' }], rect(0, 0, 1, 0), [], rect(10, 0, 14, 0));
    expect(tiled.add.map((l) => l.rect)).toEqual([rect(10, 0), rect(12, 0), rect(14, 0)]);
    const cut = copyLinks([{ rect: rect(0, 0, 2, 0), target: 'x' }], rect(1, 0, 3, 1), [], rect(10, 0, 11, 1));
    expect(cut.add.map((l) => l.rect)).toEqual([rect(10, 0, 11, 0)]);
    expect(copyLinks([other], rect(0, 0), [], rect(1, 1))).toEqual({ add: [], remove: [] });
  });

  it('stops copying links a long way past the source', () => {
    const many = copyLinks([{ rect: rect(0, 0), target: 'x' }], rect(0, 0, 0, 1), [], rect(0, 0, 30_000, 1));
    expect(many.add).toHaveLength(10_000);
  });

  it('moves links on one row with their rows when sorted', () => {
    const a: Hyperlink = { rect: rect(1, 0), target: 'a', id: 0 };
    const b: Hyperlink = { rect: rect(2, 1), target: 'b' };
    const tall: Hyperlink = { rect: rect(1, 2, 2, 2), target: 'tall' };
    const outside: Hyperlink = { rect: rect(2, 0, 2, 5), target: 'wide' };
    const change = sortLinks([a, b, tall, outside], rect(1, 0, 3, 2), 1, [3, 2, 1]);
    expect(change.remove).toEqual([a]);
    expect(change.add).toEqual([{ rect: rect(3, 0), target: 'a' }]);
  });

  it('describes how the links differ from the file’s', () => {
    expect(linkDifferences([big, other], [other, { rect: big.rect, target: 'https://new.example' }, { rect: rect(9, 0), location: 'Data!B2' }])).toEqual([
      { ref: 'A1:C3', address: 'https://new.example', was: true },
      { ref: 'A10', address: 'Data!B2', was: false },
    ]);
    expect(linkDifferences([big], [])).toEqual([{ ref: 'A1:C3', address: null, was: true }]);
  });

  it("doesn't add a link twice", () => {
    const fresh = { rect: rect(4, 4), target: 'x' };
    expect(applyLinkChange([other], { add: [fresh, { ...fresh }, other], remove: [] }).added).toEqual([fresh]);
  });
});
