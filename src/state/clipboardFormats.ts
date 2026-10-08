import type { BorderEdge, FormatChange } from '../core/types';

/** Larger clipboard HTML is pasted without its formats, as reading it all would hold the page up. */
const MAX_HTML = 4_000_000;

/** CSS colour names Excel and Google Sheets write, as hex. */
const NAMED: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  lime: '#00FF00',
  green: '#008000',
  blue: '#0000FF',
  yellow: '#FFFF00',
  aqua: '#00FFFF',
  cyan: '#00FFFF',
  fuchsia: '#FF00FF',
  magenta: '#FF00FF',
  silver: '#C0C0C0',
  gray: '#808080',
  grey: '#808080',
  maroon: '#800000',
  olive: '#808000',
  navy: '#000080',
  purple: '#800080',
  teal: '#008080',
  orange: '#FFA500',
};

/** A CSS colour as "#RRGGBB"; null for the automatic one; undefined when it can't be read. */
function cssColor(value: string): string | null | undefined {
  const v = value.trim().toLowerCase();
  if (v === 'windowtext' || v === 'auto' || v === 'inherit') return null;
  if (NAMED[v]) return NAMED[v];
  let m = /^#([0-9a-f]{6})$/.exec(v);
  if (m) return '#' + m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (m) return ('#' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/.exec(v);
  if (m) return '#' + [m[1], m[2], m[3]].map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('').toUpperCase();
  return undefined;
}

/** CSS escapes undone: "\0022" and "\#" as Excel writes number formats. */
function unescapeCss(s: string): string {
  return s.replace(/\\([0-9a-fA-F]{1,6})\s?|\\(.)/g, (_, hex: string | undefined, ch: string | undefined) => (hex ? String.fromCodePoint(parseInt(hex, 16)) : (ch ?? '')));
}

/** Excel's named number formats in clipboard HTML. */
const NAMED_FORMATS: Record<string, string> = {
  percent: '0.00%',
  fixed: '0.00',
  standard: '#,##0.00',
  'short date': 'm/d/yyyy',
  'general date': 'm/d/yyyy h:mm',
  'short time': 'h:mm',
  'long time': 'h:mm:ss AM/PM',
  general: 'General',
};

function edge(value: string): BorderEdge | null | undefined {
  const v = value.trim().toLowerCase();
  if (v === 'none' || v === '0' || v.startsWith('0 ') || v.includes('hidden')) return null;
  const parts = v.split(/\s+/);
  const style = parts.find((p) => /^(solid|dashed|dotted|double)$/.test(p));
  if (!style) return undefined;
  const width = parts.map((p) => /^([\d.]+)(pt|px)$/.exec(p)).find((m) => m);
  // Excel's thin, medium and thick are 0.5, 1 and 1.5 points, or 1, 2 and 3 pixels.
  const size = width ? Number(width[1]) * (width[2] === 'pt' ? 2 : 1) : 1;
  const color = parts.map(cssColor).find((c) => c !== undefined);
  return {
    width: style === 'double' ? 3 : size >= 3 ? 3 : size >= 2 ? 2 : 1,
    pattern: style as BorderEdge['pattern'],
    ...(color ? { color } : {}),
  };
}

/** A cell's formats from its CSS declarations; only what they say, so the rest of the destination's format stays. */
function formatsOf(css: string): FormatChange | null {
  const out: FormatChange = {};
  const border: NonNullable<FormatChange['border']> = {};
  for (const decl of css.split(';')) {
    const colon = decl.indexOf(':');
    if (colon < 0) continue;
    const prop = decl.slice(0, colon).trim().toLowerCase();
    const value = decl.slice(colon + 1).trim();
    const v = value.toLowerCase();
    switch (prop) {
      case 'font-weight':
        out.bold = v === 'bold' || v === 'bolder' || Number(v) >= 600;
        break;
      case 'font-style':
        out.italic = v === 'italic' || v === 'oblique';
        break;
      case 'text-decoration':
      case 'text-decoration-line':
        if (v.includes('underline')) out.underline = true;
        if (v.includes('line-through')) out.strike = true;
        if (v === 'none') Object.assign(out, { underline: false, strike: false });
        break;
      case 'color': {
        const c = cssColor(v);
        if (c !== undefined) out.color = c;
        break;
      }
      case 'background':
      case 'background-color': {
        const c = v === 'transparent' || v === 'none' ? null : cssColor(v.split(/\s+/)[0]);
        if (c !== undefined) out.fill = c;
        break;
      }
      case 'text-align':
        if (v === 'left' || v === 'center' || v === 'right' || v === 'justify') out.hAlign = v;
        else if (v === 'general' || v === 'start') out.hAlign = null;
        break;
      case 'vertical-align':
        if (v === 'top' || v === 'bottom') out.vAlign = v;
        else if (v === 'middle') out.vAlign = 'center';
        break;
      case 'white-space':
        out.wrap = v === 'normal' || v === 'pre-wrap';
        break;
      case 'font-family': {
        const family = value.split(',')[0].trim().replace(/^["']|["']$/g, '');
        if (family && family.length < 64) out.fontFamily = family;
        break;
      }
      case 'font-size': {
        const m = /^([\d.]+)(pt|px)$/.exec(v);
        if (m) out.fontSizePt = Math.round(Number(m[1]) * (m[2] === 'px' ? 0.75 : 1) * 2) / 2;
        break;
      }
      case 'mso-number-format': {
        const code = unescapeCss(value.replace(/^"|"$/g, ''));
        out.numFmt = NAMED_FORMATS[code.toLowerCase()] ?? code;
        break;
      }
      case 'border':
      case 'border-top':
      case 'border-bottom':
      case 'border-left':
      case 'border-right': {
        const e = edge(value);
        if (e === undefined) break;
        const sides = prop === 'border' ? (['top', 'bottom', 'left', 'right'] as const) : ([prop.slice(7)] as ('top' | 'bottom' | 'left' | 'right')[]);
        for (const side of sides) border[side] = e;
        break;
      }
    }
  }
  if (Object.keys(border).length) out.border = border;
  return Object.keys(out).length ? out : null;
}

/** The class rules in the HTML's <style> blocks, by class name. */
function classRules(html: string, lower: string): Map<string, string> {
  const rules = new Map<string, string>();
  for (let at = lower.indexOf('<style'); at !== -1; at = lower.indexOf('<style', at + 1)) {
    const end = lower.indexOf('</style', at);
    if (end === -1) break;
    const css = html.slice(html.indexOf('>', at) + 1, end);
    // Rule by rule: a selector, then its declarations in braces. Only single-class selectors (.xl65) count.
    for (let from = 0; ; ) {
      const open = css.indexOf('{', from);
      const close = open === -1 ? -1 : css.indexOf('}', open);
      if (close === -1) break;
      const selector = css.slice(from, open).replace(/<!--|-->/g, '').trim();
      const m = /^\.([\w-]+)$/.exec(selector);
      if (m) rules.set(m[1].toLowerCase(), (rules.get(m[1].toLowerCase()) ?? '') + ';' + css.slice(open + 1, close));
      from = close + 1;
    }
    at = end;
  }
  return rules;
}

const ENTITIES: Record<string, string> = { quot: '"', apos: "'", amp: '&', lt: '<', gt: '>', nbsp: ' ' };

/** An attribute's value in an HTML tag, quoted or not, with its character references read. */
function htmlAttr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}(?:\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+)))?`, 'i').exec(tag);
  if (!m) return undefined;
  return (m[1] ?? m[2] ?? m[3] ?? '').replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (all, ref: string) => {
    if (ref[0] !== '#') return ENTITIES[ref.toLowerCase()] ?? all;
    const code = ref[1] === 'x' || ref[1] === 'X' ? parseInt(ref.slice(2), 16) : Number(ref.slice(1));
    return code > 0 && code < 0x110000 ? String.fromCodePoint(code) : all;
  });
}

/**
 * Calls `fn` for each cell of the first table in clipboard HTML, with its
 * start tag and where it goes in the grid parseClipboardText lays out:
 * merged cells take up their whole span, as they do in the text, and the
 * cells they cover get `fn` with a null tag.
 */
export function eachClipboardCell(html: string, fn: (tag: string | null, row: number, col: number) => void): void {
  const lower = html.toLowerCase();
  /** Rows still covered below, by column, by a cell spanning several rows. */
  const covered: number[] = [];
  let at = 0;
  for (let r = 0; ; r++) {
    const tr = lower.indexOf('<tr', at);
    if (tr === -1) break;
    const end = lower.indexOf('</tr', tr);
    const stop = end === -1 ? lower.length : end;
    let col = 0;
    const skipCovered = () => {
      while (covered[col] > 0) {
        covered[col]--;
        fn(null, r, col++);
      }
    };
    let cell = lower.indexOf('<t', tr + 3);
    while (cell !== -1 && cell < stop) {
      const close = lower.indexOf('>', cell);
      if (close === -1) break;
      // <td> and <th>, not <thead> or <tbody>.
      if (/^<t[dh][\s>]/.test(lower.slice(cell, cell + 4))) {
        skipCovered();
        const tag = html.slice(cell, close);
        const span = Math.max(1, Number(htmlAttr(tag, 'colspan')) || 1);
        const down = Math.max(1, Number(htmlAttr(tag, 'rowspan')) || 1);
        for (let k = 0; k < span; k++) {
          fn(k === 0 ? tag : null, r, col + k);
          if (down > 1) covered[col + k] = down - 1;
        }
        col += span;
      }
      cell = lower.indexOf('<t', close);
    }
    // Columns after the row's last cell that a cell above still covers.
    for (; col < covered.length; col++) {
      if (covered[col] > 0) covered[col]--;
      fn(null, r, col);
    }
    at = stop + 1;
  }
}

/**
 * The formats of each cell in the HTML Excel or Google Sheets puts on the
 * clipboard, laid out as parseClipboardText lays out its text: from the
 * cell's classes in the <style> block (Excel) and its style attribute
 * (Google Sheets). Null for a cell with none, and undefined when there is
 * nothing to read.
 */
export function clipboardFormats(html: string): (FormatChange | null)[][] | undefined {
  if (html.length > MAX_HTML) return undefined;
  const lower = html.toLowerCase();
  if (lower.indexOf('<td') === -1) return undefined;
  const rules = classRules(html, lower);
  const rows: (FormatChange | null)[][] = [];
  let any = false;
  eachClipboardCell(html, (tag, row, col) => {
    let css = '';
    if (tag) {
      for (const name of (htmlAttr(tag, 'class') ?? '').toLowerCase().split(/\s+/)) css += rules.get(name) ?? '';
      css += ';' + (htmlAttr(tag, 'style') ?? '');
    }
    const f = css ? formatsOf(css) : null;
    if (f) any = true;
    (rows[row] ??= [])[col] = f;
  });
  return any ? rows : undefined;
}
