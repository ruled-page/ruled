import type { BorderEdge, CellBorders, DiffStyle, DisplayStyle, TableElement, TableLook } from './types';
import { attr, child, children, descendants, parseXml, type XmlNode } from './xml';

/** Excel's built-in number formats (en-US). Ids 14 and 22 follow the user's locale. */
const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

const DAY_FIRST_LOCALES = /^(en-(GB|AU|NZ|IE|ZA|IN)|fr|de|es|it|pt|nl|da|sv|nb|fi|pl|ru|tr)/i;

export function builtinNumFmt(id: number, locale: string): string {
  if ((id === 14 || id === 22) && DAY_FIRST_LOCALES.test(locale)) {
    return id === 14 ? 'dd/mm/yyyy' : 'dd/mm/yyyy h:mm';
  }
  return BUILTIN_NUMFMTS[id] ?? 'General';
}

/** Builtin id for a format code, when Excel has one. */
function builtinIdFor(code: string): number | undefined {
  for (const [id, c] of Object.entries(BUILTIN_NUMFMTS)) {
    if (c === code && id !== '14' && id !== '22') return Number(id);
  }
  return undefined;
}

const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
];

/** Office default theme, in Excel's theme-index order (lt1, dk1, lt2, dk2, accents, links). */
const DEFAULT_THEME = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
];

export function parseThemeColors(themeXml: string | undefined): string[] {
  if (!themeXml) return DEFAULT_THEME;
  const scheme = descendants(parseXml(themeXml), 'clrScheme')[0];
  if (!scheme) return DEFAULT_THEME;
  const read = (name: string, fallback: string) => {
    const node = child(scheme, name);
    const srgb = child(node, 'srgbClr');
    if (srgb) return attr(srgb, 'val') ?? fallback;
    const sys = child(node, 'sysClr');
    return attr(sys, 'lastClr') ?? fallback;
  };
  // The file lists dk1, lt1, dk2, lt2, but SpreadsheetML theme indexes swap each pair.
  return [
    read('lt1', 'FFFFFF'), read('dk1', '000000'), read('lt2', 'E7E6E6'), read('dk2', '44546A'),
    read('accent1', '4472C4'), read('accent2', 'ED7D31'), read('accent3', 'A5A5A5'),
    read('accent4', 'FFC000'), read('accent5', '5B9BD5'), read('accent6', '70AD47'),
    read('hlink', '0563C1'), read('folHlink', '954F72'),
  ];
}

/**
 * Theme tints the way Excel computes them: in Windows' integer HLS space,
 * where hue, lightness and saturation run 0–240. Float HSL comes out a
 * shade off from the colours Excel shows.
 */
const HLSMAX = 240;
const RGBMAX = 255;
const idiv = (a: number, b: number) => Math.trunc(a / b);

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = idiv((max + min) * HLSMAX + RGBMAX, 2 * RGBMAX);
  if (max === min) return [idiv(HLSMAX * 2, 3), l, 0];
  const s =
    l <= HLSMAX / 2
      ? idiv((max - min) * HLSMAX + idiv(max + min, 2), max + min)
      : idiv((max - min) * HLSMAX + idiv(2 * RGBMAX - max - min, 2), 2 * RGBMAX - max - min);
  const delta = (c: number) => idiv((max - c) * idiv(HLSMAX, 6) + idiv(max - min, 2), max - min);
  let h = r === max ? delta(b) - delta(g) : g === max ? idiv(HLSMAX, 3) + delta(r) - delta(b) : idiv(2 * HLSMAX, 3) + delta(g) - delta(r);
  if (h < 0) h += HLSMAX;
  if (h > HLSMAX) h -= HLSMAX;
  return [h, l, s];
}

function hueToRgb(n1: number, n2: number, h: number): number {
  if (h < 0) h += HLSMAX;
  if (h > HLSMAX) h -= HLSMAX;
  if (h < idiv(HLSMAX, 6)) return n1 + idiv((n2 - n1) * h + idiv(HLSMAX, 12), idiv(HLSMAX, 6));
  if (h < idiv(HLSMAX, 2)) return n2;
  if (h < idiv(HLSMAX * 2, 3)) return n1 + idiv((n2 - n1) * (idiv(HLSMAX * 2, 3) - h) + idiv(HLSMAX, 12), idiv(HLSMAX, 6));
  return n1;
}

function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) {
    const v = idiv(l * RGBMAX + idiv(HLSMAX, 2), HLSMAX);
    return [v, v, v];
  }
  const m2 = l <= HLSMAX / 2 ? idiv(l * (HLSMAX + s) + idiv(HLSMAX, 2), HLSMAX) : l + s - idiv(l * s + idiv(HLSMAX, 2), HLSMAX);
  const m1 = 2 * l - m2;
  const c = (hue: number) => Math.min(RGBMAX, Math.max(0, idiv(hueToRgb(m1, m2, hue) * RGBMAX + idiv(HLSMAX, 2), HLSMAX)));
  return [c(h + idiv(HLSMAX, 3)), c(h), c(h - idiv(HLSMAX, 3))];
}

/** A "#RRGGBB" colour lightened (positive) or darkened (negative), as Excel tints theme colours. */
export function tintColor(color: string, tint: number): string {
  return '#' + applyTint(color.replace('#', '').toUpperCase(), tint);
}

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, l, s] = rgbToHls(parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
  const lum = Math.trunc(tint < 0 ? l * (1 + tint) : l * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint)));
  return hlsToRgb(h, Math.min(HLSMAX, Math.max(0, lum)), s)
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

export function resolveColor(node: XmlNode | undefined, theme: string[], indexed: string[]): string | undefined {
  if (!node || attr(node, 'auto') === '1') return undefined;
  let hex: string | undefined;
  const rgb = attr(node, 'rgb');
  const th = attr(node, 'theme');
  const ix = attr(node, 'indexed');
  if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
  else if (th !== undefined) hex = theme[Number(th)];
  else if (ix !== undefined) hex = indexed[Number(ix)];
  if (!hex) return undefined;
  const tint = Number(attr(node, 'tint') ?? 0);
  return '#' + applyTint(hex.toUpperCase(), tint);
}

/** Excel border styles drawn as width and pattern. */
const BORDER_STYLES: Record<string, Omit<BorderEdge, 'color'>> = {
  thin: { width: 1, pattern: 'solid' },
  hair: { width: 1, pattern: 'dotted' },
  dotted: { width: 1, pattern: 'dotted' },
  dashed: { width: 1, pattern: 'dashed' },
  dashDot: { width: 1, pattern: 'dashed' },
  dashDotDot: { width: 1, pattern: 'dashed' },
  medium: { width: 2, pattern: 'solid' },
  mediumDashed: { width: 2, pattern: 'dashed' },
  mediumDashDot: { width: 2, pattern: 'dashed' },
  mediumDashDotDot: { width: 2, pattern: 'dashed' },
  slantDashDot: { width: 2, pattern: 'dashed' },
  thick: { width: 3, pattern: 'solid' },
  double: { width: 3, pattern: 'double' },
};

export function parseBorder(node: XmlNode | undefined, theme: string[], indexed: string[]): DiffStyle['border'] {
  if (!node) return undefined;
  const edge = (...names: string[]): BorderEdge | undefined => {
    for (const name of names) {
      const side = child(node, name);
      const kind = BORDER_STYLES[attr(side, 'style') ?? ''];
      if (side && kind) return { ...kind, color: resolveColor(child(side, 'color'), theme, indexed) };
    }
    return undefined;
  };
  const b: DiffStyle['border'] = { left: edge('left', 'start'), right: edge('right', 'end'), top: edge('top'), bottom: edge('bottom') };
  // Inner edges only mean something in table styles and conditional formats.
  if (child(node, 'horizontal')) b.horizontal = edge('horizontal');
  if (child(node, 'vertical')) b.vertical = edge('vertical');
  return Object.values(b).some((e) => e) ? b : undefined;
}

/** The outer edges of a border, which is all a cell format uses. */
function cellBorders(b: DiffStyle['border']): CellBorders | undefined {
  if (!b || !(b.left || b.right || b.top || b.bottom)) return undefined;
  const { left, right, top, bottom } = b;
  return { left, right, top, bottom };
}

/** A differential format (dxf) from styles.xml, as used by conditional formatting and table styles. */
function parseDxf(node: XmlNode, theme: string[], indexed: string[]): DiffStyle {
  const out: DiffStyle = {};
  const font = child(node, 'font');
  if (font) {
    const on = (name: string) => {
      const n = child(font, name);
      return n ? attr(n, 'val') !== '0' && attr(n, 'val') !== 'false' : undefined;
    };
    out.bold = on('b');
    out.italic = on('i');
    out.strike = on('strike');
    const u = child(font, 'u');
    if (u) out.underline = attr(u, 'val') !== 'none';
    out.color = resolveColor(child(font, 'color'), theme, indexed);
  }
  const pattern = child(child(node, 'fill'), 'patternFill');
  if (pattern && attr(pattern, 'patternType') !== 'none') {
    // In a dxf a solid fill's colour is usually bgColor, unlike in cell formats.
    out.fill = resolveColor(child(pattern, 'bgColor'), theme, indexed) ?? resolveColor(child(pattern, 'fgColor'), theme, indexed);
  }
  out.border = parseBorder(child(node, 'border'), theme, indexed);
  for (const k of Object.keys(out) as (keyof DiffStyle)[]) if (out[k] === undefined) delete out[k];
  return out;
}

const flag = (node: XmlNode | undefined) => !!node && attr(node, 'val') !== '0' && attr(node, 'val') !== 'false';

interface FontInfo {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  sizePt: number;
  family: string;
  color?: string;
}

export interface StyleAddition {
  xfRaw: string;
  numFmt?: { id: number; code: string };
}

/**
 * The workbook's cell formats (cellXfs) resolved for display, plus any formats
 * added during the session so pasted dates and percentages display correctly.
 */
export class StyleBook {
  readonly display: DisplayStyle[] = [];
  private readonly xfRaw: string[] = [];
  private readonly customNumFmts = new Map<number, string>();
  private readonly derived = new Map<string, number>();
  readonly additions: StyleAddition[] = [];
  /** Differential formats, by dxfId. */
  readonly dxfs: DiffStyle[] = [];
  /** Custom table styles defined in the workbook, by name. */
  readonly tableStyles = new Map<string, TableLook>();
  /** Theme colours (hex, no "#") in SpreadsheetML theme-index order. */
  readonly theme: string[];
  private readonly indexed: string[] = [...INDEXED_COLORS];
  private nextNumFmtId = 164;

  constructor(stylesXml: string | undefined, themeXml: string | undefined, locale: string) {
    const theme = parseThemeColors(themeXml);
    this.theme = theme;
    if (!stylesXml) {
      this.display.push(defaultStyle());
      this.xfRaw.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
      return;
    }
    const root = parseXml(stylesXml);
    const indexed = this.indexed;
    const custom = children(child(child(root, 'colors'), 'indexedColors'), 'rgbColor');
    custom.forEach((c, i) => {
      const rgb = attr(c, 'rgb');
      if (rgb) indexed[i] = rgb.slice(-6);
    });

    for (const nf of children(child(root, 'numFmts'), 'numFmt')) {
      const id = Number(attr(nf, 'numFmtId'));
      this.customNumFmts.set(id, attr(nf, 'formatCode') ?? 'General');
      this.nextNumFmtId = Math.max(this.nextNumFmtId, id + 1);
    }

    const fonts: FontInfo[] = children(child(root, 'fonts'), 'font').map((f) => ({
      bold: flag(child(f, 'b')),
      italic: flag(child(f, 'i')),
      underline: !!child(f, 'u') && attr(child(f, 'u'), 'val') !== 'none',
      strike: flag(child(f, 'strike')),
      sizePt: Number(attr(child(f, 'sz'), 'val') ?? 11),
      family: attr(child(f, 'name'), 'val') ?? 'Calibri',
      color: resolveColor(child(f, 'color'), theme, indexed),
    }));

    const fills: (string | undefined)[] = children(child(root, 'fills'), 'fill').map((f) => {
      const pattern = child(f, 'patternFill');
      if (pattern) {
        const type = attr(pattern, 'patternType') ?? 'none';
        if (type === 'none' || type === 'gray125') return undefined;
        return resolveColor(child(pattern, 'fgColor'), theme, indexed) ?? resolveColor(child(pattern, 'bgColor'), theme, indexed);
      }
      const stop = descendants(child(f, 'gradientFill'), 'stop')[0];
      return resolveColor(child(stop, 'color'), theme, indexed);
    });

    const borders = children(child(root, 'borders'), 'border').map((b) => parseBorder(b, theme, indexed));
    for (const d of children(child(root, 'dxfs'), 'dxf')) this.dxfs.push(parseDxf(d, theme, indexed));
    for (const ts of children(child(root, 'tableStyles'), 'tableStyle')) {
      const look: TableLook = { elements: {}, rowStripe: 1, colStripe: 1 };
      for (const el of children(ts, 'tableStyleElement')) {
        const type = attr(el, 'type') as TableElement;
        const dxf = this.dxfs[Number(attr(el, 'dxfId') ?? -1)];
        if (dxf) look.elements[type] = dxf;
        const size = Number(attr(el, 'size') ?? 1);
        if (type === 'firstRowStripe') look.rowStripe = size;
        if (type === 'firstColumnStripe') look.colStripe = size;
      }
      this.tableStyles.set(attr(ts, 'name') ?? '', look);
    }

    const cellXfsNode = child(root, 'cellXfs');
    for (const xf of children(cellXfsNode, 'xf')) {
      const font = fonts[Number(attr(xf, 'fontId') ?? 0)] ?? fonts[0];
      const align = child(xf, 'alignment');
      this.display.push({
        numFmt: this.numFmtCode(Number(attr(xf, 'numFmtId') ?? 0), locale),
        bold: font?.bold ?? false,
        italic: font?.italic ?? false,
        underline: font?.underline ?? false,
        strike: font?.strike ?? false,
        fontSizePx: Math.round(((font?.sizePt ?? 11) * 4) / 3 * 10) / 10,
        fontFamily: font?.family ?? 'Calibri',
        color: font?.color && font.color !== '#000000' ? font.color : undefined,
        fill: fills[Number(attr(xf, 'fillId') ?? 0)],
        hAlign: attr(align, 'horizontal'),
        vAlign: attr(align, 'vertical'),
        wrap: attr(align, 'wrapText') === '1' || attr(align, 'wrapText') === 'true',
        indent: Number(attr(align, 'indent') ?? 0),
        border: cellBorders(borders[Number(attr(xf, 'borderId') ?? 0)]),
      });
    }
    if (this.display.length === 0) this.display.push(defaultStyle());

    // Raw <xf> elements, kept verbatim so derived formats inherit everything else.
    const block = /<(\w+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(\w+:)?cellXfs>/.exec(stylesXml);
    if (block) {
      const re = /<(?:\w+:)?xf\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?xf>)/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(block[2]))) this.xfRaw.push(m[0]);
    }
    while (this.xfRaw.length < this.display.length) {
      this.xfRaw.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    }
  }

  /** A colour element (rgb, theme or indexed) resolved against this workbook's palette. */
  color(node: XmlNode | undefined): string | undefined {
    return resolveColor(node, this.theme, this.indexed);
  }

  /** A differential format written inline, as x14 conditional formats do. */
  dxf(node: XmlNode): DiffStyle {
    return parseDxf(node, this.theme, this.indexed);
  }

  private numFmtCode(id: number, locale: string): string {
    return this.customNumFmts.get(id) ?? builtinNumFmt(id, locale);
  }

  get(index: number): DisplayStyle {
    return this.display[index] ?? this.display[0];
  }

  /**
   * The format index for `base` with a different number format, creating it
   * (and the number format) if the workbook does not have one yet.
   */
  withNumFmt(base: number, code: string): { index: number; created: boolean } {
    if (this.get(base).numFmt === code) return { index: base, created: false };
    const key = `${base}|${code}`;
    const existing = this.derived.get(key);
    if (existing !== undefined) return { index: existing, created: false };

    let id = builtinIdFor(code);
    let numFmt: StyleAddition['numFmt'];
    if (id === undefined) {
      for (const [cid, c] of this.customNumFmts) if (c === code) id = cid;
    }
    if (id === undefined) {
      id = this.nextNumFmtId++;
      this.customNumFmts.set(id, code);
      numFmt = { id, code };
    }

    let raw = this.xfRaw[base] ?? this.xfRaw[0];
    raw = /\bnumFmtId="\d+"/.test(raw)
      ? raw.replace(/\bnumFmtId="\d+"/, `numFmtId="${id}"`)
      : raw.replace(/^<((?:\w+:)?xf)\b/, `<$1 numFmtId="${id}"`);
    raw = /\bapplyNumberFormat="[^"]*"/.test(raw)
      ? raw.replace(/\bapplyNumberFormat="[^"]*"/, 'applyNumberFormat="1"')
      : raw.replace(/^<((?:\w+:)?xf)\b/, '<$1 applyNumberFormat="1"');

    const index = this.display.length;
    this.display.push({ ...this.get(base), numFmt: code });
    this.xfRaw.push(raw);
    this.additions.push({ xfRaw: raw, numFmt });
    this.derived.set(key, index);
    return { index, created: true };
  }
}

function defaultStyle(): DisplayStyle {
  return {
    numFmt: 'General',
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontSizePx: 14.7,
    fontFamily: 'Calibri',
    wrap: false,
    indent: 0,
  };
}
