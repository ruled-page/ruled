/**
 * Small, forgiving XML utilities for SpreadsheetML parts. Office files are
 * well-formed and have no DTDs, so a tiny tree builder is enough for the
 * metadata parts; worksheet cell data is scanned separately for speed.
 *
 * Files are untrusted, so nothing here may loop on truncated input or take
 * more than linear time: searches use indexOf rather than regular expressions
 * that backtrack, and a tag that never ends is reported as a damaged file.
 */
import { msg, TextError } from '../i18n/message';

/** The error for a part that ends in the middle of a tag, comment or CDATA section. */
export function damaged(): TextError {
  return new TextError(msg('core.open.damaged'));
}

export interface XmlNode {
  /** Local name without namespace prefix, e.g. "sheet". */
  name: string;
  /** Qualified name as written, e.g. "x:sheet". */
  qname: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/**
 * A copy of text cut from a part's XML, to keep. A browser keeps a longer
 * cut as a view of the string it was cut from, so the formulas and text of
 * a sheet would otherwise keep its whole XML in memory after it is read.
 */
export function detach(s: string): string {
  // Joining makes a new string, and cutting it again copies that, not the XML.
  return s.length < 13 ? s : (' ' + s).slice(1);
}

export function decodeEntities(s: string): string {
  if (s.indexOf('&') === -1) return s;
  return s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|\w+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return code <= 0x10ffff ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e] ?? m;
  });
}

export function escapeXml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));
}

/**
 * SpreadsheetML escapes characters XML cannot carry as `_xHHHH_`. A literal
 * "_x0041_" in cell text is itself escaped as "_x005F_x0041_".
 */
export function decodeOoxmlText(s: string): string {
  if (s.indexOf('_x') === -1) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

export function encodeOoxmlText(s: string): string {
  return s
    .replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_')
    // oxlint-disable-next-line no-control-regex -- these are the characters XML can't hold
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, (c) =>
      `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`,
    );
}

export function localName(qname: string): string {
  const i = qname.indexOf(':');
  return i === -1 ? qname : qname.slice(i + 1);
}

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * Regular expression source for a tag's attributes, a character or a whole
 * quoted value at a time: repeated, it runs to the tag's closing ">" past
 * any ">" a value holds ("a > b"), where [^>]* would stop short.
 */
export const IN_TAG = `(?:[^>"']|"[^"]*"|'[^']*')`;

/** Parses an attribute string such as ` r="A1" s='2'` into a record, in one pass. */
export function parseAttrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  const n = s.length;
  let i = 0;
  while (i < n) {
    while (i < n && (isSpace(s.charCodeAt(i)) || s.charCodeAt(i) === 47)) i++;
    const nameStart = i;
    while (i < n && !isSpace(s.charCodeAt(i)) && s.charCodeAt(i) !== 61 && s.charCodeAt(i) !== 47 && s.charCodeAt(i) !== 62) i++;
    const name = s.slice(nameStart, i);
    while (i < n && isSpace(s.charCodeAt(i))) i++;
    if (s.charCodeAt(i) !== 61) {
      // A name without a value: skip it.
      if (i === nameStart) i++;
      continue;
    }
    i++;
    while (i < n && isSpace(s.charCodeAt(i))) i++;
    const q = s.charCodeAt(i);
    if (q !== 34 && q !== 39) continue;
    const end = s.indexOf(s[i], i + 1);
    if (end === -1) break;
    if (name) out[name] = decodeEntities(s.slice(i + 1, end));
    i = end + 1;
  }
  return out;
}

/** Index of the `>` that closes the tag starting at `start`, skipping quoted values. */
export function findTagEnd(src: string, start: number): number {
  let quote = 0;
  for (let i = start + 1; i < src.length; i++) {
    const ch = src.charCodeAt(i);
    if (quote) {
      if (ch === quote) quote = 0;
    } else if (ch === 34 || ch === 39) {
      quote = ch;
    } else if (ch === 62) {
      return i;
    }
  }
  return src.length - 1;
}

/** Offset just past `close`, searching from `from`; a missing terminator means the part is damaged. */
function past(src: string, close: string, from: number): number {
  const at = src.indexOf(close, from);
  if (at === -1) throw damaged();
  return at + close.length;
}

/**
 * Throws `damaged()` unless every tag ends before the next one starts, and
 * every comment, CDATA section and processing instruction ends. Each part is
 * checked when it is first read, so later scans that run to the next ">" stop
 * within the tag they started in. XML doesn't allow "<" inside a tag.
 */
export function checkTerminated(xml: string): void {
  let lt = xml.indexOf('<');
  while (lt !== -1) {
    if (xml.charCodeAt(lt + 1) === 33 || xml.charCodeAt(lt + 1) === 63) {
      let end: number;
      if (xml.startsWith('<!--', lt)) end = past(xml, '-->', lt + 4);
      else if (xml.startsWith('<![CDATA[', lt)) end = past(xml, ']]>', lt + 9);
      else if (xml.startsWith('<?', lt)) end = past(xml, '?>', lt + 2);
      else end = past(xml, '>', lt);
      lt = xml.indexOf('<', end);
      continue;
    }
    const gt = xml.indexOf('>', lt + 1);
    if (gt === -1) throw damaged();
    const next = xml.indexOf('<', lt + 1);
    if (next !== -1 && next < gt) throw damaged();
    lt = next;
  }
}

export function parseXml(src: string): XmlNode {
  const doc: XmlNode = { name: '#document', qname: '#document', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [doc];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt === -1) break;
    if (lt > i) stack[stack.length - 1].text += decodeEntities(src.slice(i, lt));
    if (src.startsWith('<?', lt)) {
      i = past(src, '?>', lt);
    } else if (src.startsWith('<!--', lt)) {
      i = past(src, '-->', lt + 4);
    } else if (src.startsWith('<![CDATA[', lt)) {
      i = past(src, ']]>', lt + 9);
      stack[stack.length - 1].text += src.slice(lt + 9, i - 3);
    } else if (src.startsWith('<!', lt)) {
      i = past(src, '>', lt);
    } else if (src.charCodeAt(lt + 1) === 47) {
      if (stack.length > 1) stack.pop();
      i = past(src, '>', lt);
    } else {
      const gt = findTagEnd(src, lt);
      if (src.charCodeAt(gt) !== 62) throw damaged();
      const selfClosing = src.charCodeAt(gt - 1) === 47;
      const inner = src.slice(lt + 1, selfClosing ? gt - 1 : gt);
      const ws = inner.search(/\s/);
      const qname = ws === -1 ? inner : inner.slice(0, ws);
      const node: XmlNode = {
        name: localName(qname),
        qname,
        attrs: ws === -1 ? {} : parseAttrs(inner.slice(ws)),
        children: [],
        text: '',
      };
      stack[stack.length - 1].children.push(node);
      if (!selfClosing) stack.push(node);
      i = gt + 1;
    }
  }
  return doc.children[0] ?? doc;
}

export function child(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

export function children(node: XmlNode | undefined, name: string): XmlNode[] {
  return node ? node.children.filter((c) => c.name === name) : [];
}

/** All descendants with the given local name, in document order. Not recursive, so deep nesting can't overflow the stack. */
export function descendants(node: XmlNode | undefined, name: string, out: XmlNode[] = []): XmlNode[] {
  if (!node) return out;
  const stack: XmlNode[] = [...node.children].reverse();
  while (stack.length) {
    const c = stack.pop()!;
    if (c.name === name) out.push(c);
    for (let k = c.children.length - 1; k >= 0; k--) stack.push(c.children[k]);
  }
  return out;
}

/** An element found by `elements`, as offsets into the text it was found in. */
export interface ElementSpan {
  /** Offset of its "<". */
  start: number;
  /** Offset just past it: after "/>" or its closing tag. */
  end: number;
  /** Its name as written, with any prefix: "x:row". */
  qname: string;
  /** The text between its name and ">" or "/>". */
  attrs: string;
  /** The text between its tags, or null when it closes itself. */
  inner: string | null;
}

/**
 * Elements with local name `name` (any prefix) in xml[from, to), in document
 * order. It searches with indexOf, so it takes linear time whatever the input;
 * an element whose tag or closing tag is missing ends the search.
 */
export function elements(xml: string, name: string, from = 0, to = xml.length, limit = Infinity): ElementSpan[] {
  const out: ElementSpan[] = [];
  let pos = from;
  while (out.length < limit) {
    const lt = xml.indexOf('<', pos);
    if (lt === -1 || lt >= to) break;
    let j = lt + 1;
    for (let c = xml.charCodeAt(j); j < to && !isSpace(c) && c !== 47 && c !== 62 && c !== 60; c = xml.charCodeAt(++j));
    const qname = xml.slice(lt + 1, j);
    if (localName(qname) !== name) {
      pos = lt + 1;
      continue;
    }
    const gt = findTagEnd(xml, lt);
    if (xml.charCodeAt(gt) !== 62 || gt >= to) break;
    if (xml.charCodeAt(gt - 1) === 47) {
      out.push({ start: lt, end: gt + 1, qname, attrs: xml.slice(j, gt - 1), inner: null });
      pos = gt + 1;
      continue;
    }
    const close = xml.indexOf(`</${qname}>`, gt + 1);
    if (close === -1 || close + qname.length + 3 > to) break;
    out.push({ start: lt, end: close + qname.length + 3, qname, attrs: xml.slice(j, gt), inner: xml.slice(gt + 1, close) });
    pos = close + qname.length + 3;
  }
  return out;
}

/** The first element called `name`, as `elements` finds it. */
export function firstElement(xml: string, name: string, from = 0, to = xml.length): ElementSpan | undefined {
  return elements(xml, name, from, to, 1)[0];
}

/** Replaces each element called `name`: fn returns its new text, or '' to remove it. */
export function replaceElements(xml: string, name: string, fn: (el: ElementSpan, text: string) => string, limit = Infinity): string {
  const found = elements(xml, name, 0, xml.length, limit);
  if (!found.length) return xml;
  let out = '';
  let pos = 0;
  for (const el of found) {
    out += xml.slice(pos, el.start) + fn(el, xml.slice(el.start, el.end));
    pos = el.end;
  }
  return out + xml.slice(pos);
}

/** How many elements called `name` open in xml, counting only their opening tags. */
export function countElements(xml: string, name: string): number {
  let n = 0;
  let pos = 0;
  for (;;) {
    const lt = xml.indexOf('<', pos);
    if (lt === -1) return n;
    let j = lt + 1;
    for (let c = xml.charCodeAt(j); j < xml.length && !isSpace(c) && c !== 47 && c !== 62 && c !== 60; c = xml.charCodeAt(++j));
    if (localName(xml.slice(lt + 1, j)) === name) n++;
    pos = lt + 1;
  }
}

/** Attribute by local name, ignoring any prefix (`r:id` matches "id"). */
export function attr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  if (name in node.attrs) return node.attrs[name];
  for (const k in node.attrs) {
    if (localName(k) === name) return node.attrs[k];
  }
  return undefined;
}

/** Concatenated text of all `<t>` descendants (rich text runs, drawing text). */
export function textOf(node: XmlNode | undefined): string {
  return descendants(node, 't')
    .map((t) => t.text)
    .join('');
}

/** A start tag with attributes set, or removed where the value is null. Names are the caller's own, not from a file. */
export function setAttributes(tag: string, attrs: Record<string, string | null>): string {
  const selfClosing = tag.endsWith('/>');
  let body = tag.slice(0, selfClosing ? -2 : -1);
  for (const [name, value] of Object.entries(attrs)) {
    body = body.replace(new RegExp(`\\s${name}="[^"]*"`, 'g'), '');
    if (value !== null) body += ` ${name}="${escapeXml(value)}"`;
  }
  return body + (selfClosing ? '/>' : '>');
}
