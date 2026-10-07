/**
 * Small, forgiving XML utilities for SpreadsheetML parts. Office files are
 * well-formed and have no DTDs, so a tiny tree builder is enough for the
 * metadata parts; worksheet cell data is scanned separately for speed.
 */

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

export function decodeEntities(s: string): string {
  if (s.indexOf('&') === -1) return s;
  return s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|\w+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(code);
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
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, (c) =>
      `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`,
    );
}

export function localName(qname: string): string {
  const i = qname.indexOf(':');
  return i === -1 ? qname : qname.slice(i + 1);
}

/** Parses an attribute string such as ` r="A1" s='2'` into a record. */
export function parseAttrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(s))) out[m[1]] = decodeEntities(m[3] ?? m[4] ?? '');
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
      i = src.indexOf('?>', lt) + 2;
    } else if (src.startsWith('<!--', lt)) {
      i = src.indexOf('-->', lt) + 3;
    } else if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt);
      stack[stack.length - 1].text += src.slice(lt + 9, end);
      i = end + 3;
    } else if (src.startsWith('<!', lt)) {
      i = src.indexOf('>', lt) + 1;
    } else if (src.charCodeAt(lt + 1) === 47) {
      if (stack.length > 1) stack.pop();
      i = src.indexOf('>', lt) + 1;
    } else {
      const gt = findTagEnd(src, lt);
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

/** All descendants with the given local name, in document order. */
export function descendants(node: XmlNode | undefined, name: string, out: XmlNode[] = []): XmlNode[] {
  if (!node) return out;
  for (const c of node.children) {
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
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
