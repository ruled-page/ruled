import { describe, expect, it } from 'vitest';
import { builtinNumFmt, parseBorder, parseThemeColors, StyleBook, tintColor } from '../src/core/styles';
import { builtinTableLook, tableCellFormat } from '../src/core/tableStyles';
import type { TableInfo, TableLook } from '../src/core/types';
import { parseXml } from '../src/core/xml';

const SS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const THEME = parseThemeColors(undefined);

describe('number formats', () => {
  it("write Excel's short date in the locale's order", () => {
    expect([builtinNumFmt(14, 'en-US'), builtinNumFmt(14, 'en-GB'), builtinNumFmt(22, 'de-DE'), builtinNumFmt(22, 'en-US')]).toEqual(['m/d/yyyy', 'dd/mm/yyyy', 'dd/mm/yyyy h:mm', 'm/d/yyyy h:mm']);
    expect(builtinNumFmt(999, 'en-US')).toBe('General');
  });

  it('reuse a format already made for the same base and code', () => {
    const book = new StyleBook(`<styleSheet ${SS}><numFmts count="1"><numFmt numFmtId="170" formatCode="0.000"/></numFmts><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`, undefined, 'en-US');
    const pct = book.withNumFmt(0, '0%');
    expect(pct.created).toBe(true);
    expect(book.withNumFmt(0, '0%')).toEqual({ index: pct.index, created: false });
    expect(book.withNumFmt(pct.index, '0%')).toEqual({ index: pct.index, created: false });
    // A code the workbook already defines keeps its id.
    const three = book.withNumFmt(0, '0.000');
    expect(book.additions.find((a) => a.xfRaw.includes(`numFmtId="170"`))).toBeDefined();
    expect(book.get(three.index).numFmt).toBe('0.000');
  });
});

describe('colours', () => {
  it('read a theme from system colours, or fall back when it has no colour scheme', () => {
    const theme = parseThemeColors(
      '<a:theme xmlns:a="a"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>' +
        '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2></a:clrScheme></a:themeElements></a:theme>',
    );
    // Light 1 and dark 1 swap places in SpreadsheetML's numbering.
    expect(theme.slice(0, 4)).toEqual(['FFFFFF', '111111', 'EEEEEE', '222222']);
    expect(parseThemeColors('<a:theme xmlns:a="a"/>')).toBe(THEME);
  });

  it('tint colours whose hue wraps around red', () => {
    const light = tintColor('#FF0080', 0.5);
    expect(light.slice(1, 3)).toBe('FF');
    expect(parseInt(light.slice(5, 7), 16)).toBeGreaterThan(parseInt(light.slice(3, 5), 16));
    expect(tintColor('#808080', -0.5)).toBe('#404040');
    expect(tintColor('#123456', 0)).toBe('#123456');
  });

  it('use a workbook palette, gradient fills and inner borders', () => {
    const book = new StyleBook(
      `<styleSheet ${SS}><colors><indexedColors><rgbColor rgb="FF00FF00"/><rgbColor/></indexedColors></colors>` +
        '<fonts count="1"><font><color indexed="0"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><gradientFill><stop position="0"><color rgb="FF0000FF"/></stop><stop position="1"><color rgb="FFFFFFFF"/></stop></gradientFill></fill></fills>' +
        '<cellXfs count="1"><xf fontId="0" fillId="2"/></cellXfs>' +
        '<dxfs count="1"><dxf><border><top style="thin"/><horizontal style="thin"><color rgb="FFFF0000"/></horizontal><vertical style="dashed"/></border></dxf></dxfs>' +
        '<tableStyles count="1"><tableStyle name="Banded"><tableStyleElement type="firstColumnStripe" dxfId="0" size="3"/></tableStyle></tableStyles></styleSheet>',
      undefined,
      'en-US',
    );
    expect(book.get(0)).toMatchObject({ color: '#00FF00', fill: '#0000FF' });
    expect(book.dxfs[0].border).toMatchObject({ horizontal: { width: 1, color: '#FF0000' }, vertical: { pattern: 'dashed' } });
    expect(book.tableStyles.get('Banded')?.colStripe).toBe(3);
    expect(parseBorder(parseXml('<border><left/><right/></border>'), THEME, [])).toBeUndefined();
  });
});

describe('built-in table styles', () => {
  it('cover each family, and none past the last Dark style', () => {
    const medium15 = builtinTableLook('TableStyleMedium15', THEME)!;
    expect(medium15.elements.firstRowStripe).toEqual({ fill: '#D9D9D9' });
    expect(medium15.elements.headerRow?.border?.bottom).toMatchObject({ width: 2, color: '#000000' });
    const dark9 = builtinTableLook('TableStyleDark9', THEME)!;
    expect(dark9.elements.headerRow).toEqual({ fill: '#000000', color: '#FFFFFF', bold: true });
    expect(dark9.elements.totalRow?.border?.top).toMatchObject({ pattern: 'double' });
    expect(builtinTableLook('TableStyleDark12', THEME)).toBeUndefined();
  });

  it('layer header, total, stripe and edge formats, the most specific last', () => {
    const fill = (f: string) => ({ fill: f });
    const look: TableLook = {
      rowStripe: 1,
      colStripe: 2,
      elements: {
        wholeTable: fill('whole'),
        firstColumnStripe: fill('stripe1'),
        secondColumnStripe: fill('stripe2'),
        lastColumn: fill('last'),
        headerRow: fill('header'),
        totalRow: fill('total'),
        firstHeaderCell: fill('header-first'),
        lastHeaderCell: fill('header-last'),
        firstTotalCell: fill('total-first'),
        lastTotalCell: fill('total-last'),
      },
    };
    const t: TableInfo = {
      name: 'T', rect: { r1: 0, c1: 0, r2: 4, c2: 4 }, headerRows: 1, totalsRows: 1, columns: [],
      showFirstColumn: true, showLastColumn: true, showRowStripes: false, showColumnStripes: true, look,
    };
    const at = (row: number) => [0, 1, 2, 3, 4].map((col) => tableCellFormat(t, row, col)?.fill);
    expect(at(0)).toEqual(['header-first', 'header', 'header', 'header', 'header-last']);
    // Columns stripe two at a time; the last column's own format wins.
    expect(at(2)).toEqual(['stripe1', 'stripe1', 'stripe2', 'stripe2', 'last']);
    expect(at(4)).toEqual(['total-first', 'total', 'total', 'total', 'total-last']);
  });
});
