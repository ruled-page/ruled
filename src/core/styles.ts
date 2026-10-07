import type { DisplayStyle } from './types';
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

function hexToHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h /= 6;
  return [h, s, l];
}

function hslToHex(h: number, s: number, l: number): string {
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l;
  let g = l;
  let b = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue(p, q, h + 1 / 3);
    g = hue(p, q, h);
    b = hue(p, q, h - 1 / 3);
  }
  const to = (x: number) => Math.round(x * 255).toString(16).padStart(2, '0');
  return (to(r) + to(g) + to(b)).toUpperCase();
}

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = hexToHsl(hex);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToHex(h, s, Math.min(1, Math.max(0, nl)));
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
  private nextNumFmtId = 164;

  constructor(stylesXml: string | undefined, themeXml: string | undefined, locale: string) {
    const theme = parseThemeColors(themeXml);
    if (!stylesXml) {
      this.display.push(defaultStyle());
      this.xfRaw.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
      return;
    }
    const root = parseXml(stylesXml);
    const indexed = [...INDEXED_COLORS];
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
