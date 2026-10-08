import { msg } from '../i18n/message';
import { applyLinkChange, linkDifferences, sameLink } from './links';
import { parseRange, rangeToA1 } from './refs';
import type { CellChange, Hyperlink } from './types';
import { attr, child, children, elements, escapeXml, findTagEnd, firstElement, localName, parseAttrs, replaceElements, type XmlNode } from './xml';
import { relsPathFor, type OoxmlPackage, type Relationship } from './zip';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HYPERLINK_REL = `${REL_NS}/hyperlink`;
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * A worksheet's <hyperlinks>: each <hyperlink>'s cells, and where it goes,
 * from its relationship (an address outside the workbook) or its location
 * (a place in it). `id` is its position among them, so an unchanged one is
 * saved as it was.
 */
export function readHyperlinks(meta: XmlNode, rels: readonly Relationship[]): Hyperlink[] {
  const out: Hyperlink[] = [];
  children(child(meta, 'hyperlinks'), 'hyperlink').forEach((h, id) => {
    const rect = parseRange(attr(h, 'ref') ?? '');
    if (!rect) return;
    const link: Hyperlink = { rect, id };
    const rid = attr(h, 'id');
    const rel = rid ? rels.find((r) => r.id === rid && r.type === 'hyperlink') : undefined;
    if (rel) link.target = rel.target;
    for (const key of ['location', 'display', 'tooltip'] as const) {
      const v = attr(h, key);
      if (v !== undefined) link[key] = v;
    }
    out.push(link);
  });
  return out;
}

/**
 * One sheet's hyperlinks as edited: the links the file had, and the list as
 * it is now once a link has been added, changed or removed.
 */
export class SheetLinks {
  private current?: Hyperlink[];

  constructor(private readonly file: readonly Hyperlink[]) {}

  get list(): readonly Hyperlink[] {
    return this.current ?? this.file;
  }

  /** Whether the links differ from the file's. */
  get changed(): boolean {
    const now = this.current;
    return !!now && (now.length !== this.file.length || now.some((l) => !this.file.some((f) => sameLink(f, l))));
  }

  /** Makes the change, returning the change that undoes it. */
  apply(op: { add: readonly Hyperlink[]; remove: readonly Hyperlink[] }): { kind: 'links'; add: Hyperlink[]; remove: Hyperlink[] } {
    const { links, added, removed } = applyLinkChange(this.list, op);
    this.current = links;
    return { kind: 'links', add: removed, remove: added };
  }

  /** How the links differ from the file's, for the list of changes. */
  notes(sheet: string): CellChange[] {
    if (!this.current) return [];
    return linkDifferences(this.file, this.current).map(({ ref, address, was }) => ({
      sheet,
      ref,
      before: '',
      after: address ?? '',
      note: address === null ? msg('links.change.removed') : was ? msg('links.change.changed', { address }) : msg('links.change.added', { address }),
    }));
  }
}

// ---------------------------------------------------------------------------
// Writing

/** An element directly inside the worksheet, after <sheetData>. */
interface Sibling {
  name: string;
  start: number;
  end: number;
}

/**
 * The elements after `from` up to the end of their parent, in one pass:
 * comments and processing instructions are stepped over, and each element
 * ends at its closing tag (worksheet elements don't nest in themselves).
 */
function siblings(xml: string, from: number): Sibling[] {
  const out: Sibling[] = [];
  let pos = from;
  for (;;) {
    const lt = xml.indexOf('<', pos);
    if (lt < 0 || xml.charCodeAt(lt + 1) === 47) break;
    if (xml.startsWith('<!--', lt) || xml.startsWith('<?', lt)) {
      const close = xml.indexOf(xml.startsWith('<!--', lt) ? '-->' : '?>', lt + 2);
      if (close < 0) break;
      pos = close + 2;
      continue;
    }
    const gt = findTagEnd(xml, lt);
    if (xml.charCodeAt(gt) !== 62) break;
    const qname = xml.slice(lt + 1, nameEnd(xml, lt + 1, gt));
    let end = gt + 1;
    if (xml.charCodeAt(gt - 1) !== 47) {
      const close = xml.indexOf(`</${qname}>`, gt + 1);
      if (close < 0) break;
      end = close + qname.length + 3;
    }
    out.push({ name: localName(qname), start: lt, end });
    pos = end;
  }
  return out;
}

/** Elements that come after <hyperlinks> in a worksheet, so a new one goes before the first of them. */
const AFTER = new Set([
  'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors',
  'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
]);

/** An attribute value as written in a tag; line breaks are kept as character references, which a parser doesn't fold into spaces. */
const attrValue = (s: string) =>
  escapeXml(
    // oxlint-disable-next-line no-control-regex -- characters XML can't hold
    s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''),
  ).replace(/\r\n|\r|\n/g, '&#10;').replace(/\t/g, '&#9;');

/** The end of an element's name, from just after its "<". */
function nameEnd(xml: string, from: number, to: number): number {
  let j = from;
  for (let c = xml.charCodeAt(j); j < to && c !== 32 && c !== 9 && c !== 10 && c !== 13 && c !== 47 && c !== 62; c = xml.charCodeAt(++j));
  return j;
}

/** The root element's start tag: its name, where its ">" is, whether it closes itself, and its attributes. */
function rootTag(xml: string): { qname: string; end: number; empty: boolean; attrs: Record<string, string> } | undefined {
  let lt = xml.indexOf('<');
  while (lt >= 0 && (xml.charCodeAt(lt + 1) === 63 || xml.charCodeAt(lt + 1) === 33)) lt = xml.indexOf('<', findTagEnd(xml, lt) + 1);
  if (lt < 0) return undefined;
  const gt = findTagEnd(xml, lt);
  const j = nameEnd(xml, lt + 1, gt);
  const empty = xml.charCodeAt(gt - 1) === 47;
  return { qname: xml.slice(lt + 1, j), end: gt, empty, attrs: parseAttrs(xml.slice(j, empty ? gt - 1 : gt)) };
}

/** The relationship id not yet used in a part's relationships: rId1, rId2… as Excel numbers them. */
function freeIds(used: Iterable<string>): () => string {
  const taken = new Set(used);
  let n = 1;
  return () => {
    while (taken.has(`rId${n}`)) n++;
    taken.add(`rId${n}`);
    return `rId${n}`;
  };
}

/**
 * The worksheet with its <hyperlinks> as `links` has them, and the sheet's
 * relationships to match: links that were in the file are written as they
 * were, new ones to an address get a relationship of their own, and the
 * relationships of links that went are removed. <hyperlinks> is added in its
 * place among the worksheet's elements when the sheet had none, and removed
 * when it has none left.
 */
export function writeHyperlinks(pkg: OoxmlPackage, path: string, xml: string, p: string, links: readonly Hyperlink[]): string {
  const data = firstElement(xml, 'sheetData');
  if (!data) return xml;
  const after = siblings(xml, data.end);
  const box = after.find((e) => e.name === 'hyperlinks');
  const boxXml = box ? xml.slice(box.start, box.end) : '';
  const inner = box ? (elements(boxXml, 'hyperlinks', 0, boxXml.length, 1)[0]?.inner ?? '') : '';
  const originals = elements(inner, 'hyperlink');
  const relIdOf = (attrs: string) => {
    const a = parseAttrs(attrs);
    for (const k in a) if (localName(k) === 'id' && k.includes(':')) return a[k];
    return undefined;
  };

  // The prefix the relationships namespace has on the root element, declaring it if it hasn't one.
  const root = rootTag(xml);
  let r = root ? Object.keys(root.attrs).find((k) => k.startsWith('xmlns:') && root.attrs[k] === REL_NS)?.slice(6) : undefined;
  let declare = '';
  if (!r && links.some((l) => l.target && (l.id === undefined || !originals[l.id]))) {
    for (let n = 0; !r; n++) if (!root || root.attrs[`xmlns:${n ? `r${n}` : 'r'}`] === undefined) r = n ? `r${n}` : 'r';
    declare = ` xmlns:${r}="${REL_NS}"`;
  }

  const rels = pkg.rels(path);
  const nextId = freeIds(rels.map((x) => x.id));
  const newRels: string[] = [];
  const keptIds = new Set<string>();
  const out: string[] = [];
  const ordered = [...links].sort((a, b) => (a.id ?? Infinity) - (b.id ?? Infinity));
  for (const l of ordered) {
    const own = l.id !== undefined ? originals[l.id] : undefined;
    if (own) {
      const id = relIdOf(own.attrs);
      if (id) keptIds.add(id);
      out.push(inner.slice(own.start, own.end));
      continue;
    }
    let attrs = ` ref="${rangeToA1(l.rect)}"`;
    if (l.target) {
      const id = nextId();
      newRels.push(`<Relationship Id="${id}" Type="${HYPERLINK_REL}" Target="${attrValue(encodeTarget(l.target))}" TargetMode="External"/>`);
      attrs += ` ${r}:id="${id}"`;
    }
    if (l.location) attrs += ` location="${attrValue(l.location)}"`;
    if (l.tooltip) attrs += ` tooltip="${attrValue(l.tooltip)}"`;
    if (l.display) attrs += ` display="${attrValue(l.display)}"`;
    out.push(`<${p}hyperlink${attrs}/>`);
  }

  // Relationships of the file's links that went, unless a link still kept uses them.
  const gone = new Set(originals.map((o) => relIdOf(o.attrs)).filter((id): id is string => !!id && !keptIds.has(id)));
  const drop = new Set(rels.filter((x) => gone.has(x.id) && x.type === 'hyperlink').map((x) => x.id));
  if (drop.size || newRels.length) writeRels(pkg, path, drop, newRels);

  const el = out.length ? `<${p}hyperlinks>${out.join('')}</${p}hyperlinks>` : '';
  let next: string;
  if (box) next = xml.slice(0, box.start) + el + xml.slice(box.end);
  else {
    const before = after.find((e) => AFTER.has(e.name));
    const at = before ? before.start : (after.length ? after[after.length - 1].end : data.end);
    next = xml.slice(0, at) + el + xml.slice(at);
  }
  // The root's start tag comes before everything changed above, so where it ends hasn't moved.
  return root && declare ? next.slice(0, root.end) + declare + next.slice(root.end) : next;
}

/** Spaces and other characters an address can't hold as they are, escaped the way browsers send them, so every reader of the package takes it. */
function encodeTarget(target: string): string {
  return target.replace(/[\s"<>\\^`{|}]/g, (c) => encodeURIComponent(c));
}

/** Removes some of a part's relationships and adds others, creating its relationships part if it has none. */
function writeRels(pkg: OoxmlPackage, path: string, drop: Set<string>, add: string[]): void {
  const relsPath = relsPathFor(path);
  let xml = pkg.text(relsPath) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELS_NS}"></Relationships>`;
  if (drop.size) xml = replaceElements(xml, 'Relationship', (el, text) => (drop.has(parseAttrs(el.attrs).Id ?? '') ? '' : text));
  const root = rootTag(xml);
  if (add.length && root) {
    // Before </Relationships>, or opening a <Relationships/> that closes itself.
    xml = root.empty
      ? `${xml.slice(0, root.end - 1)}>${add.join('')}</${root.qname}>${xml.slice(root.end + 1)}`
      : `${xml.slice(0, xml.lastIndexOf('</'))}${add.join('')}${xml.slice(xml.lastIndexOf('</'))}`;
  }
  pkg.setText(relsPath, xml);
}
