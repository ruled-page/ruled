import type { BorderEdge, CellBorders, DiffStyle, DisplayStyle, FormatChange, TableElement, TableLook } from './types';
import { attr, child, children, descendants, elements, escapeXml, findTagEnd, firstElement, parseAttrs, parseXml, setAttributes, type XmlNode } from './xml';

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
export function builtinIdFor(code: string): number | undefined {
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
  /** The file's fonts, fills and borders as written, which derived formats copy and change. */
  private readonly parts: Record<PartKind, string[]> = { font: [], fill: [], border: [] };
  /** Fonts, fills and borders added during the session, to append to the file's. */
  readonly addedParts: Record<PartKind, string[]> = { font: [], fill: [], border: [] };
  private readonly partIds = new Map<string, number>();

  constructor(
    stylesXml: string | undefined,
    themeXml: string | undefined,
    private readonly locale: string,
  ) {
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
      const protection = child(xf, 'protection');
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
        ...(['0', 'false'].includes(attr(protection, 'locked') ?? '') ? { unlocked: true } : {}),
        ...(['1', 'true'].includes(attr(protection, 'hidden') ?? '') ? { formulaHidden: true } : {}),
      });
    }
    if (this.display.length === 0) this.display.push(defaultStyle());

    // Raw <xf> elements, kept verbatim so derived formats inherit everything else.
    const block = firstElement(stylesXml, 'cellXfs')?.inner;
    if (block) for (const xf of elements(block, 'xf')) this.xfRaw.push(block.slice(xf.start, xf.end));
    for (const kind of ['font', 'fill', 'border'] as const) {
      const list = firstElement(stylesXml, `${kind}s`)?.inner;
      if (list) for (const e of elements(list, kind)) this.parts[kind].push(list.slice(e.start, e.end));
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
    return this.withFormat(base, { numFmt: code });
  }

  /**
   * The format index for `base` changed as `change` says, creating it, and
   * any font, fill, border or number format it needs, if the workbook does
   * not have it yet. Everything else about the base format is kept.
   */
  withFormat(base: number, change: FormatChange): { index: number; created: boolean } {
    const key = `${base}|${JSON.stringify(change)}`;
    const existing = this.derived.get(key);
    if (existing !== undefined) return { index: existing, created: false };
    const shown = this.get(base);
    const next = changedStyle(shown, change);
    if (JSON.stringify(next) === JSON.stringify(shown)) return { index: base, created: false };

    const raw = this.xfRaw[base] ?? this.xfRaw[0];
    const p = /^<(\w+:)?xf\b/.exec(raw)?.[1] ?? '';
    const tagEnd = findTagEnd(raw, 0);
    const open = raw.slice(0, tagEnd + 1);
    const attrs = parseAttrs(open.slice(open.indexOf(' ') + 1, open.endsWith('/>') ? -2 : -1));
    let inner = open.endsWith('/>') ? '' : raw.slice(tagEnd + 1, raw.lastIndexOf('</'));
    const set: Record<string, string | null> = {};
    let numFmt: StyleAddition['numFmt'];

    if (change.numFmt !== undefined && change.numFmt !== shown.numFmt) {
      // The short date and date-time formats show in the user's locale's order, as Excel's built-in ones do.
      let id = [14, 22].find((n) => builtinNumFmt(n, this.locale) === change.numFmt) ?? builtinIdFor(change.numFmt);
      if (id === undefined) for (const [cid, c] of this.customNumFmts) if (c === change.numFmt) id = cid;
      if (id === undefined) {
        id = this.nextNumFmtId++;
        this.customNumFmts.set(id, change.numFmt);
        numFmt = { id, code: change.numFmt };
      }
      Object.assign(set, { numFmtId: String(id), applyNumberFormat: '1' });
    }
    const font = this.partXml('font', Number(attrs.fontId ?? 0)) ?? this.parts.font[0];
    if (FONT_KEYS.some((k) => change[k] !== undefined)) {
      Object.assign(set, { fontId: String(this.part('font', changedFont(font ?? `<${p}font><${p}sz val="11"/><${p}name val="Calibri"/></${p}font>`, change, p))), applyFont: '1' });
    }
    if (change.fill !== undefined) {
      const fill = change.fill && `<${p}fill><${p}patternFill patternType="solid"><${p}fgColor rgb="${argb(change.fill)}"/><${p}bgColor indexed="64"/></${p}patternFill></${p}fill>`;
      Object.assign(set, { fillId: fill ? String(this.part('fill', fill)) : '0', applyFill: '1' });
    }
    if (change.border) {
      const border = this.partXml('border', Number(attrs.borderId ?? 0)) ?? `<${p}border/>`;
      Object.assign(set, { borderId: String(this.part('border', changedBorder(border, change.border, p))), applyBorder: '1' });
    }
    if (change.hAlign !== undefined || change.vAlign !== undefined || change.wrap !== undefined || change.indent !== undefined) {
      inner = changedAlignment(inner, change, p);
      set.applyAlignment = '1';
    }

    const xf = inner ? setAttributes(open.endsWith('/>') ? open.slice(0, -2) + '>' : open, set) + inner + `</${p}xf>` : setAttributes(open.endsWith('/>') ? open : open.slice(0, -1) + '/>', set);
    const index = this.display.length;
    this.display.push(next);
    this.xfRaw.push(xf);
    this.additions.push({ xfRaw: xf, numFmt });
    this.derived.set(key, index);
    return { index, created: true };
  }

  /** A font, fill or border by id, the file's or one added since. */
  private partXml(kind: PartKind, id: number): string | undefined {
    const own = this.parts[kind];
    return id < own.length ? own[id] : this.addedParts[kind][id - own.length];
  }

  /** The id of a font, fill or border written this way, adding it after the file's when it is new. */
  private part(kind: PartKind, xml: string): number {
    const key = `${kind}|${xml}`;
    let id = this.partIds.get(key);
    if (id === undefined) {
      id = this.parts[kind].length + this.addedParts[kind].length;
      this.addedParts[kind].push(xml);
      this.partIds.set(key, id);
    }
    return id;
  }
}

type PartKind = 'font' | 'fill' | 'border';

const FONT_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSizePt', 'color'] as const;
/** A font's children in the order Excel writes them. */
const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const BORDER_SIDES = ['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'] as const;
const BORDER_NAMES: Record<string, string> = { '1solid': 'thin', '2solid': 'medium', '3solid': 'thick', '1dashed': 'dashed', '2dashed': 'mediumDashed', '1dotted': 'dotted', '2dotted': 'dotted', '3double': 'double' };

const argb = (hex: string) => 'FF' + hex.replace('#', '').toUpperCase();

/** An element's children, by local name, as written. */
function childrenByName(inner: string, names: readonly string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const name of names) {
    const e = firstElement(inner, name);
    if (e) out.set(name, inner.slice(e.start, e.end));
  }
  return out;
}

function innerOf(raw: string): string {
  const end = findTagEnd(raw, 0);
  return raw.charCodeAt(end - 1) === 47 ? '' : raw.slice(end + 1, raw.lastIndexOf('</'));
}

function changedFont(raw: string, c: FormatChange, p: string): string {
  const parts = childrenByName(innerOf(raw), FONT_ORDER);
  const toggle = (name: string, on: boolean | undefined) => {
    if (on === true) parts.set(name, `<${p}${name}/>`);
    else if (on === false) parts.delete(name);
  };
  toggle('b', c.bold);
  toggle('i', c.italic);
  toggle('strike', c.strike);
  toggle('u', c.underline);
  if (c.fontSizePt !== undefined) parts.set('sz', `<${p}sz val="${c.fontSizePt}"/>`);
  if (c.color === null) parts.delete('color');
  else if (c.color !== undefined) parts.set('color', `<${p}color rgb="${argb(c.color)}"/>`);
  if (c.fontFamily !== undefined) {
    parts.set('name', `<${p}name val="${escapeXml(c.fontFamily)}"/>`);
    // A scheme ties the font to the theme's, whatever its name says.
    for (const n of ['family', 'charset', 'scheme']) parts.delete(n);
  }
  return `<${p}font>${FONT_ORDER.map((n) => parts.get(n) ?? '').join('')}</${p}font>`;
}

function changedBorder(raw: string, edges: NonNullable<FormatChange['border']>, p: string): string {
  const inner = innerOf(raw);
  const parts = childrenByName(inner, BORDER_SIDES);
  // Excel 2010 and later may write start and end for left and right.
  for (const [from, to] of [['start', 'left'], ['end', 'right']] as const) {
    const e = !parts.has(to) ? firstElement(inner, from) : undefined;
    if (e) parts.set(to, inner.slice(e.start, e.end).replace(new RegExp(`(</?${p.replace(':', '\\:')})${from}\\b`, 'g'), `$1${to}`));
  }
  for (const [side, edge] of Object.entries(edges)) {
    if (edge === undefined) continue;
    const name = edge && BORDER_NAMES[`${edge.width}${edge.pattern}`];
    const color = edge?.color ? `<${p}color rgb="${argb(edge.color)}"/>` : `<${p}color indexed="64"/>`;
    parts.set(side, name ? `<${p}${side} style="${name}">${color}</${p}${side}>` : `<${p}${side}/>`);
  }
  for (const side of ['left', 'right', 'top', 'bottom', 'diagonal']) if (!parts.has(side)) parts.set(side, `<${p}${side}/>`);
  const end = findTagEnd(raw, 0);
  const open = raw.charCodeAt(end - 1) === 47 ? raw.slice(0, end - 1) + '>' : raw.slice(0, end + 1);
  return open + BORDER_SIDES.map((s) => parts.get(s) ?? '').join('') + `</${p}border>`;
}

function changedAlignment(inner: string, c: FormatChange, p: string): string {
  const el = firstElement(inner, 'alignment');
  const tag = el ? `<${el.qname}${el.attrs}/>` : `<${p}alignment/>`;
  const set: Record<string, string | null> = {};
  if (c.hAlign !== undefined) set.horizontal = c.hAlign;
  if (c.vAlign !== undefined) set.vertical = c.vAlign === 'bottom' ? null : c.vAlign;
  if (c.wrap !== undefined) set.wrapText = c.wrap ? '1' : null;
  if (c.indent !== undefined) set.indent = c.indent > 0 ? String(c.indent) : null;
  const next = setAttributes(tag, set);
  const keep = /\s\w+="/.test(next) ? next : '';
  // <alignment> comes first among an xf's children.
  return el ? inner.slice(0, el.start) + keep + inner.slice(el.end) : keep + inner;
}

/** A format as shown once `c` has changed it. */
function changedStyle(s: DisplayStyle, c: FormatChange): DisplayStyle {
  const next: DisplayStyle = { ...s };
  if (c.numFmt !== undefined) next.numFmt = c.numFmt;
  for (const k of ['bold', 'italic', 'underline', 'strike', 'wrap'] as const) if (c[k] !== undefined) next[k] = c[k];
  if (c.fontSizePt !== undefined) next.fontSizePx = Math.round(((c.fontSizePt * 4) / 3) * 10) / 10;
  if (c.fontFamily !== undefined) next.fontFamily = c.fontFamily;
  if (c.color !== undefined) next.color = c.color && c.color.toUpperCase() !== '#000000' ? c.color.toUpperCase() : undefined;
  if (c.fill !== undefined) next.fill = c.fill?.toUpperCase() ?? undefined;
  if (c.hAlign !== undefined) next.hAlign = c.hAlign ?? undefined;
  if (c.vAlign !== undefined) next.vAlign = c.vAlign === 'bottom' ? undefined : (c.vAlign ?? undefined);
  if (c.indent !== undefined) next.indent = c.indent;
  if (c.border) {
    const b: CellBorders = { ...s.border };
    for (const [side, edge] of Object.entries(c.border) as [keyof CellBorders, BorderEdge | null | undefined][]) {
      if (edge === undefined) continue;
      if (edge) b[side] = edge;
      else delete b[side];
    }
    next.border = Object.keys(b).length ? b : undefined;
  }
  for (const k of Object.keys(next) as (keyof DisplayStyle)[]) if (next[k] === undefined) delete next[k];
  return next;
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
