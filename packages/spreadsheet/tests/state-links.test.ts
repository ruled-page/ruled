import { describe, expect, it } from 'vitest';
import { CellFlag, VT, type Hyperlink } from '../src/core/types';
import { effectiveStyle } from '../src/state/cellStyle';
import { cellLink, linkStyle } from '../src/state/links';
import { sheet, style } from './state-fixtures';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const web: Hyperlink = { rect: rect(0, 0), target: 'https://example.com', tooltip: 'Tip', id: 0 };
const tall: Hyperlink = { rect: rect(2, 1, 4, 2), location: 'Other!A1', id: 1 };

describe("a sheet view's links", () => {
  it('finds the link on a cell, on one row or several', () => {
    const view = sheet([[0, 0, 'Example']], { links: [web, tall, { rect: rect(0, 3), target: 'https://second.example' }] });
    expect(view.linkAt(0, 0)).toBe(web);
    expect(view.linkAt(3, 2)).toBe(tall);
    expect(view.linkAt(0, 3)?.target).toBe('https://second.example');
    expect(view.linkAt(0, 1)).toBeUndefined();
    expect(view.linkAt(5, 1)).toBeUndefined();
  });

  it('follows links the worker adds and removes', () => {
    const view = sheet([], { links: [web] });
    const before = view.layoutVersion;
    view.applyLayout({ kind: 'links', add: [{ rect: rect(1, 1), target: 'https://new.example' }], remove: [web] });
    expect(view.linkAt(0, 0)).toBeUndefined();
    expect(view.linkAt(1, 1)?.target).toBe('https://new.example');
    expect(view.layoutVersion).toBe(before + 1);
  });

  it('takes the link of a cell from a merge, or from a HYPERLINK formula', () => {
    const view = sheet(
      [
        [0, 0, 'Example'],
        [6, 0, 'Literal', 0, 'HYPERLINK("https://literal.example","Literal")'],
        [7, 0, 'From A1', 0, 'HYPERLINK(A1,"From A1")'],
        [8, 0, 3, 0, 'SUM(1,2)'],
      ],
      { links: [web, tall], merges: [rect(0, 0, 1, 1)] },
    );
    expect(cellLink(view, 1, 1)).toEqual({ target: 'https://example.com', tooltip: 'Tip', link: web });
    expect(cellLink(view, 2, 2)?.location).toBe('Other!A1');
    expect(cellLink(view, 6, 0)).toEqual({ target: 'https://literal.example' });
    // Its address isn't known until the worker works it out.
    expect(cellLink(view, 7, 0)).toEqual({});
    view.apply({ sheet: 0, row: 7, col: 0, t: VT.String, s: 'From A1', style: 0, f: 'HYPERLINK(A1,"From A1")', flags: CellFlag.Formula, link: 'https://worked-out.example' });
    expect(cellLink(view, 7, 0)).toEqual({ target: 'https://worked-out.example' });
    expect(cellLink(view, 8, 0)).toBeUndefined();
    expect(cellLink(view, 9, 0)).toBeUndefined();
  });
});

describe('how a link is drawn', () => {
  it("is underlined in the link colour, unless the cell's font has a colour of its own", () => {
    const plain = style();
    expect(linkStyle(plain, '#0563c1')).toEqual({ ...plain, color: '#0563c1', underline: true });
    expect(linkStyle(plain, '#0563c1')).toBe(linkStyle(plain, '#0563c1'));
    const red = style({ color: '#FF0000' });
    expect(linkStyle(red, '#0563c1')).toBe(red);
  });

  it("applies to a linked cell's own format when the grid asks for it", () => {
    const view = sheet([[0, 0, 'Example'], [0, 1, 'Plain']], { links: [web] });
    const styles = [style()];
    expect(effectiveStyle(view, styles, 0, 0, undefined, '#0563c1')).toMatchObject({ color: '#0563c1', underline: true });
    expect(effectiveStyle(view, styles, 0, 0)).toBe(styles[0]);
    expect(effectiveStyle(view, styles, 0, 1, undefined, '#0563c1')).toBe(styles[0]);
  });

  it("takes the editor's link colour for a font in the workbook theme's, with the font's own underline", () => {
    expect(linkStyle(style({ color: '#0563C1', linkFont: true, underline: false }), '#8ab4f8')).toMatchObject({ color: '#8ab4f8', underline: false });
  });
});
