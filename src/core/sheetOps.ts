import { parseA1Ref, stringifyA1Ref, stringifyTokens, tokenize } from '@borgar/fx';
import { msg, TextError } from '../i18n/message';
import { quoteSheetName } from './refs';
import { nearestVisible, sheetIndexMap, sheetOpProblem, type SheetOp } from './sheetRules';
import type { SheetInfo } from './types';
import type { SheetModel, WorkbookModel } from './workbook';
import { decodeEntities, elements, escapeXml, firstElement, localName, parseAttrs, replaceElements, type ElementSpan } from './xml';
import { relsPathFor, type OoxmlPackage } from './zip';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const WORKSHEET_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const WORKSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Elements that hold formulas: cells', validation's, conditional formatting's, sparklines', charts'. */
const FORMULA_TAGS = ['f', 'formula', 'formula1', 'formula2'];

/**
 * Adds, renames, deletes, moves, hides or unhides a sheet in a saved
 * workbook package. Everything that names the sheet follows a rename:
 * formulas on every sheet, defined names, charts, validation and conditional
 * formatting, table formulas, pivot sources, links within the workbook,
 * form controls and the document's list of its parts. Deleting makes those
 * references #REF!, as Excel does, and removes the parts only that sheet
 * used. Everything that counts sheets by position (local names, the active
 * tab) is renumbered.
 *
 * `model` must be read from `pkg` as it is now. Throws, changing nothing,
 * when Excel would refuse the change. Returns how many cell formulas now
 * show #REF!.
 */
export function applySheetOp(pkg: OoxmlPackage, model: WorkbookModel, op: SheetOp): { broken: number } {
  const problem = sheetOpProblem(op, model.sheets.map((m) => m.info), model.structureLocked);
  if (problem) throw new TextError(problem);
  const book = readBook(pkg, model.workbookPath);
  switch (op.type) {
    case 'addSheet':
      addSheet(pkg, model, book, op);
      return { broken: 0 };
    case 'renameSheet':
      renameSheet(pkg, model, book, op);
      return { broken: 0 };
    case 'deleteSheet':
      return deleteSheet(pkg, model, book, op.sheet);
    case 'moveSheet':
      moveSheet(pkg, model, book, op);
      return { broken: 0 };
    case 'hideSheet':
      hideSheet(pkg, model, book, op);
      return { broken: 0 };
  }
}

// ---------------------------------------------------------------------------
// The workbook part

interface Book {
  path: string;
  xml: string;
  /** Each <sheet> element, in order. */
  entries: { text: string; attrs: Record<string, string>; rid: string; sheetId: number }[];
}

function readBook(pkg: OoxmlPackage, path: string): Book {
  const xml = pkg.text(path) ?? '';
  const sheets = firstElement(xml, 'sheets');
  const entries = elements(sheets?.inner ?? '', 'sheet').map((el) => {
    const attrs = parseAttrs(el.attrs);
    const idKey = Object.keys(attrs).find((k) => k.includes(':') && localName(k) === 'id') ?? 'r:id';
    return { text: (sheets?.inner ?? '').slice(el.start, el.end), attrs, rid: attrs[idKey] ?? '', sheetId: Number(attrs.sheetId) || 0 };
  });
  return { path, xml, entries };
}

/** The element written again from its parsed attributes, with some set (or, for null, removed). */
function retag(el: ElementSpan, changes: Record<string, string | null>): string {
  const attrs = parseAttrs(el.attrs);
  for (const [k, v] of Object.entries(changes)) {
    if (v === null) delete attrs[k];
    else attrs[k] = v;
  }
  const open = `<${el.qname}${Object.entries(attrs).map(([k, v]) => ` ${k}="${escapeXml(v)}"`).join('')}`;
  return el.inner === null ? `${open}/>` : `${open}>${el.inner}</${el.qname}>`;
}

/**
 * Writes the workbook part with its <sheet> elements in the order given,
 * defined names passed through `name` (which returns the formula and the
 * sheet the name is local to, or null to drop it), and the first workbook
 * view's active and first tabs passed through `view`.
 */
function writeBook(
  pkg: OoxmlPackage,
  book: Book,
  sheets: string[],
  name: (formula: string, local: number | undefined) => { formula: string; local: number | undefined } | null,
  view: (activeTab: number, firstSheet: number) => [number, number],
): void {
  let xml = book.xml;
  const list = firstElement(xml, 'sheets');
  if (list) xml = xml.slice(0, list.start) + `<${list.qname}${list.attrs}>${sheets.join('')}</${list.qname}>` + xml.slice(list.end);
  xml = replaceElements(xml, 'definedName', (el, text) => {
    if (el.inner === null) return text;
    const a = parseAttrs(el.attrs);
    const local = a.localSheetId !== undefined ? Number(a.localSheetId) : undefined;
    const formula = decodeEntities(el.inner);
    const next = name(formula, local);
    if (!next) return '';
    if (next.formula === formula && next.local === local) return text;
    const open = retag({ ...el, inner: null }, { localSheetId: next.local === undefined ? null : String(next.local) }).slice(0, -2);
    return `${open}>${escapeXml(next.formula)}</${el.qname}>`;
  });
  // A <definedNames> left with none is dropped.
  xml = replaceElements(xml, 'definedNames', (el, text) => (el.inner !== null && !el.inner.trim() ? '' : text));
  xml = replaceElements(xml, 'workbookView', (el, text) => {
    const a = parseAttrs(el.attrs);
    const active = Number(a.activeTab ?? 0);
    const first = Number(a.firstSheet ?? 0);
    const [na, nf] = view(active, first);
    if (na === active && nf === first) return text;
    return retag(el, { activeTab: na || a.activeTab !== undefined ? String(na) : null, firstSheet: nf || a.firstSheet !== undefined ? String(nf) : null });
  });
  pkg.setText(book.path, xml);
}

/** The active tab of the workbook's first view. */
function activeTab(book: Book): number {
  const view = firstElement(book.xml, 'workbookView');
  return Number((view && parseAttrs(view.attrs).activeTab) ?? 0);
}

/** Marks a sheet's first view as selected, or not, as the sheet Excel opens on must be. */
function setTabSelected(pkg: OoxmlPackage, path: string | undefined, on: boolean): void {
  const xml = path ? pkg.textOnce(path) : undefined;
  if (!path || xml === undefined) return;
  const next = replaceElements(
    xml,
    'sheetView',
    (el, text) => {
      const v = parseAttrs(el.attrs).tabSelected;
      return (v === '1' || v === 'true') === on ? text : retag(el, { tabSelected: on ? '1' : null });
    },
    1,
  );
  if (next !== xml) pkg.setText(path, next);
}

// ---------------------------------------------------------------------------
// References by name

interface RefChange {
  /** The sheet renamed or deleted, in lower case. */
  sheet: string;
  /** Its new name; null when it is deleted. */
  to: string | null;
  /** Every sheet's name in order, for 3-D references that lose an end. */
  order: string[];
  /** Tables that went with a deleted sheet, in lower case. */
  tables: Set<string>;
}

/** Whether a formula may refer to the sheet or its tables: a quick test before tokenizing it. */
function mentions(text: string, c: RefChange): boolean {
  const lower = text.toLowerCase();
  if (lower.includes(c.sheet) || lower.includes(c.sheet.replace(/'/g, "''"))) return true;
  for (const t of c.tables) if (lower.includes(t)) return true;
  return false;
}

/**
 * A sheet reference's context ("Data", or "Jan:Mar" for a 3-D reference)
 * after the change: undefined when it doesn't involve the sheet, null when it
 * is lost. A 3-D reference that loses an end sheet ends at the next one in.
 */
function changeContext(context: string, c: RefChange): string | null | undefined {
  const ends = context.split(':');
  const hit = ends.map((e) => e.toLowerCase() === c.sheet);
  if (ends.length > 2 || !hit.some(Boolean)) return undefined;
  if (c.to !== null) return ends.map((e, i) => (hit[i] ? c.to : e)).join(':');
  if (ends.length === 1) return null;
  const [a, b] = ends.map((e) => c.order.findIndex((n) => n.toLowerCase() === e.toLowerCase()));
  const step = a < b ? 1 : -1;
  const from = hit[0] ? a + step : a;
  const to = hit[1] ? b - step : b;
  if (a < 0 || b < 0 || a === b || (to - from) * step < 0) return null;
  return from === to ? c.order[from] : `${c.order[from]}:${c.order[to]}`;
}

/** A formula (file form) after the change; the same string when nothing in it changes. */
function changeRefs(formula: string, c: RefChange): string {
  if (!mentions(formula, c)) return formula;
  let tokens;
  try {
    tokens = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return formula;
  }
  let changed = false;
  for (const t of tokens) {
    let next: string | undefined;
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_named') {
      const ref = parseA1Ref(t.value, { allowTernary: false });
      // References into other workbooks have the workbook in their context too.
      if (!ref || ref.context?.length !== 1) continue;
      const context = changeContext(ref.context[0], c);
      if (context === undefined) continue;
      if (context === null) {
        next = '#REF!';
      } else {
        ref.context = [context];
        next = stringifyA1Ref(ref);
      }
    } else if (t.type === 'structured') {
      const table = t.value.slice(0, t.value.indexOf('['));
      if (c.tables.has(table.slice(table.lastIndexOf('!') + 1).toLowerCase())) next = '#REF!';
    }
    if (next !== undefined && next !== t.value) {
      t.value = next;
      changed = true;
    }
  }
  return changed ? stringifyTokens(tokens) : formula;
}

/** Rewrites the text of formula elements, which hold escaped XML text. */
function rewriteFormulas(xml: string, tags: string[], fn: (f: string) => string): string {
  for (const tag of tags) {
    xml = replaceElements(xml, tag, (el, text) => {
      if (el.inner === null) return text;
      const f = decodeEntities(el.inner);
      const next = fn(f);
      return next === f ? text : `<${el.qname}${el.attrs}>${escapeXml(next)}</${el.qname}>`;
    });
  }
  return xml;
}

/** Rewrites attributes that hold a reference, such as a link's location or a control's linked cell. */
function rewriteRefAttrs(xml: string, tag: string, names: string[], fn: (f: string) => string): string {
  return replaceElements(xml, tag, (el, text) => {
    const a = parseAttrs(el.attrs);
    const changes: Record<string, string> = {};
    for (const n of names) {
      const next = a[n] !== undefined ? fn(a[n]) : undefined;
      if (next !== undefined && next !== a[n]) changes[n] = next;
    }
    return Object.keys(changes).length ? retag(el, changes) : text;
  });
}

function rewritePart(pkg: OoxmlPackage, path: string, fn: (xml: string) => string): void {
  const xml = pkg.textOnce(path);
  if (xml === undefined) return;
  const next = fn(xml);
  if (next !== xml) pkg.setText(path, next);
}

/**
 * Follows a rename or delete everywhere a sheet is named: cells, validation,
 * conditional formatting and sparklines on every sheet (and links, for a
 * rename), defined names (by `name`), charts, table formulas, pivot sources
 * and form controls. Returns how many cell formulas lost a reference.
 */
function followNames(pkg: OoxmlPackage, model: WorkbookModel, c: RefChange, skip: number): number {
  const fx = (f: string) => changeRefs(f, c);
  let broken = 0;
  model.sheets.forEach((m, s) => {
    if (s === skip || !m.path) return;
    const formulas = m.scan ? [...m.scan.grid.formulas.values()].filter((f) => mentions(f, c)) : [];
    for (const f of formulas) {
      const next = fx(f);
      if (next !== f && next.includes('#REF!') && !f.includes('#REF!')) broken++;
    }
    rewritePart(pkg, m.path, (xml) => rewriteSheet(xml, m, formulas.length > 0, fx, c.to !== null));
  });
  for (const path of pkg.list()) {
    if (/^xl\/charts\/chart\w*\.xml$/.test(path)) {
      rewritePart(pkg, path, (xml) => {
        const next = rewriteFormulas(xml, ['f'], fx);
        return c.to === null ? next : pivotChartSource(next, c);
      });
    } else if (/^xl\/tables\/table\w*\.xml$/.test(path)) {
      rewritePart(pkg, path, (xml) => rewriteFormulas(xml, ['calculatedColumnFormula', 'totalsRowFormula'], fx));
    } else if (/pivotCacheDefinition\w*\.xml$/.test(path)) {
      rewritePart(pkg, path, (xml) => pivotSource(xml, c));
    } else if (/^xl\/ctrlProps\/[^/]+\.xml$/.test(path)) {
      rewritePart(pkg, path, (xml) => rewriteRefAttrs(xml, 'formControlPr', ['fmlaLink', 'fmlaRange', 'fmlaTxbx'], fx));
    } else if (/\.vml$/i.test(path)) {
      rewritePart(pkg, path, (xml) => rewriteFormulas(xml, ['FmlaLink', 'FmlaRange', 'FmlaTxbx'], fx));
    } else if (/^xl\/drawings\/drawing\w*\.xml$/.test(path)) {
      // Shapes whose text comes from a cell.
      rewritePart(pkg, path, (xml) => rewriteRefAttrs(xml, 'sp', ['textlink'], fx));
    }
  }
  return broken;
}

/** A worksheet's formulas after the change; `cells` says whether any cell formula needs it, `links` whether links follow. */
function rewriteSheet(xml: string, m: SheetModel, cells: boolean, fx: (f: string) => string, links: boolean): string {
  const scan = m.scan;
  if (!scan) return rewriteFormulas(xml, FORMULA_TAGS, fx);
  const [a, b] = scan.sheetDataStart >= 0 ? [scan.sheetDataStart, scan.sheetDataEnd] : [0, 0];
  const body = cells ? rewriteFormulas(xml.slice(a, b), ['f'], fx) : undefined;
  const before = xml.slice(b);
  let tail = rewriteFormulas(before, FORMULA_TAGS, fx);
  // Excel itself leaves links to a renamed sheet broken; here they follow it.
  if (links) tail = rewriteRefAttrs(tail, 'hyperlink', ['location'], fx);
  if (body === undefined && tail === before) return xml;
  return xml.slice(0, a) + (body ?? xml.slice(a, b)) + tail;
}

/** A pivot cache's source after the change: renamed, or #REF! when its sheet or table is gone. */
function pivotSource(xml: string, c: RefChange): string {
  let gone = false;
  const next = replaceElements(xml, 'worksheetSource', (el, text) => {
    const a = parseAttrs(el.attrs);
    const onSheet = a.sheet !== undefined && a.sheet.toLowerCase() === c.sheet;
    if (onSheet && c.to !== null) return retag(el, { sheet: c.to });
    if (!onSheet && !(a.name !== undefined && c.tables.has(a.name.toLowerCase()))) return text;
    gone = true;
    return retag(el, { sheet: null, name: null, ref: '#REF!' });
  });
  // There is nothing left to refresh it from.
  return gone ? replaceElements(next, 'pivotCacheDefinition', (el) => retag(el, { refreshOnLoad: null }), 1) : next;
}

/** A pivot chart names its pivot table as "[Book.xlsx]Sheet!PivotTable1". */
function pivotChartSource(xml: string, c: RefChange): string {
  return replaceElements(xml, 'pivotSource', (src, text) => {
    const name = firstElement(src.inner ?? '', 'name');
    if (!name?.inner) return text;
    const full = decodeEntities(name.inner);
    const bang = full.lastIndexOf('!');
    const close = full.indexOf(']');
    let sheet = full.slice(close + 1, bang);
    if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    if (bang < 0 || sheet.toLowerCase() !== c.sheet) return text;
    const renamed = `${full.slice(0, close + 1)}${quoteSheetName(c.to!)}${full.slice(bang)}`;
    const inner = src.inner!;
    return `<${src.qname}${src.attrs}>${inner.slice(0, name.start)}<${name.qname}${name.attrs}>${escapeXml(renamed)}</${name.qname}>${inner.slice(name.end)}</${src.qname}>`;
  });
}

// ---------------------------------------------------------------------------
// The document's list of its parts (docProps/app.xml)

interface TitleGroup {
  name: string;
  items: string[];
}

/**
 * Edits the titles of parts the document properties list ("Worksheets: 3,
 * Named Ranges: 1" and the names themselves), keeping their counts right.
 * Left alone if it isn't laid out as Excel writes it.
 */
function editTitles(pkg: OoxmlPackage, fn: (groups: TitleGroup[]) => void): void {
  const path = pkg.rels('').find((r) => r.type === 'extended-properties')?.target ?? 'docProps/app.xml';
  const xml = pkg.text(path);
  const pairs = xml ? firstElement(xml, 'HeadingPairs') : undefined;
  const titles = xml ? firstElement(xml, 'TitlesOfParts') : undefined;
  const pv = pairs?.inner ? firstElement(pairs.inner, 'vector') : undefined;
  const tv = titles?.inner ? firstElement(titles.inner, 'vector') : undefined;
  if (!xml || !pairs || !titles || !pv?.inner || !tv?.inner || parseAttrs(tv.attrs).baseType !== 'lpstr') return;
  const names = elements(tv.inner, 'lpstr').map((e) => decodeEntities(e.inner ?? ''));
  const variants = elements(pv.inner, 'variant');
  const groups: TitleGroup[] = [];
  let at = 0;
  for (let i = 0; i + 1 < variants.length; i += 2) {
    const label = firstElement(variants[i].inner ?? '', 'lpstr');
    const count = firstElement(variants[i + 1].inner ?? '', 'i4');
    if (!label || !count) return;
    const n = Number(count.inner);
    groups.push({ name: decodeEntities(label.inner ?? ''), items: names.slice(at, at + n) });
    at += n;
  }
  if (at !== names.length || variants.length % 2) return;
  const before = JSON.stringify(groups);
  fn(groups);
  if (JSON.stringify(groups) === before) return;
  const kept = groups.filter((g) => g.items.length);
  const vt = pv.qname.slice(0, pv.qname.indexOf(':') + 1);
  const lpstr = (s: string) => `<${vt}lpstr>${escapeXml(s)}</${vt}lpstr>`;
  const pairsVector = `${retag({ ...pv, inner: null }, { size: String(kept.length * 2) }).slice(0, -2)}>${kept
    .map((g) => `<${vt}variant>${lpstr(g.name)}</${vt}variant><${vt}variant><${vt}i4>${g.items.length}</${vt}i4></${vt}variant>`)
    .join('')}</${pv.qname}>`;
  const items = kept.flatMap((g) => g.items);
  const titlesVector = `${retag({ ...tv, inner: null }, { size: String(items.length) }).slice(0, -2)}>${items.map(lpstr).join('')}</${tv.qname}>`;
  // Offsets of each vector in the whole part.
  const innerAt = (el: ElementSpan) => el.end - el.inner!.length - el.qname.length - 3;
  const edits = [
    { start: innerAt(pairs) + pv.start, end: innerAt(pairs) + pv.end, text: pairsVector },
    { start: innerAt(titles) + tv.start, end: innerAt(titles) + tv.end, text: titlesVector },
  ].sort((x, y) => y.start - x.start);
  let out = xml;
  for (const e of edits) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  pkg.setText(path, out);
}

/** A title names something on the sheet, such as "Data!Print_Area". */
function onSheet(title: string, sheet: string): boolean {
  const lower = title.toLowerCase();
  const name = sheet.toLowerCase();
  return lower.startsWith(`${name}!`) || lower.startsWith(`'${name.replace(/'/g, "''")}'!`);
}

/** Puts the sheets each group lists in the workbook's new order, leaving its other titles where they are. */
function orderTitles(groups: TitleGroup[], order: string[]): void {
  const pos = new Map(order.map((n, i) => [n, i]));
  for (const g of groups) {
    const slots = g.items.flatMap((x, i) => (pos.has(x) ? [i] : []));
    const sorted = slots.map((i) => g.items[i]).sort((a, b) => pos.get(a)! - pos.get(b)!);
    slots.forEach((slot, k) => (g.items[slot] = sorted[k]));
  }
}

// ---------------------------------------------------------------------------
// Parts and relationships

/** Removes a part, its relationships and its content type. */
function removeParts(pkg: OoxmlPackage, paths: Set<string>): void {
  for (const p of paths) {
    pkg.remove(p);
    pkg.remove(relsPathFor(p));
  }
  const names = new Set([...paths].map((p) => `/${p}`.toLowerCase()));
  rewritePart(pkg, '[Content_Types].xml', (xml) =>
    replaceElements(xml, 'Override', (el, text) => (names.has((parseAttrs(el.attrs).PartName ?? '').toLowerCase()) ? '' : text)),
  );
}

/** Every part reachable by relationships from `from` (the package itself when ''). */
function reachable(pkg: OoxmlPackage, from: string[]): Set<string> {
  const seen = new Set<string>(from.filter(Boolean));
  const queue = [...from];
  while (queue.length) {
    for (const r of pkg.rels(queue.pop()!)) {
      if (r.external || seen.has(r.target) || !pkg.has(r.target)) continue;
      seen.add(r.target);
      queue.push(r.target);
    }
  }
  return seen;
}

function removeRelationship(pkg: OoxmlPackage, source: string, id: string): void {
  rewritePart(pkg, relsPathFor(source), (xml) => replaceElements(xml, 'Relationship', (el, text) => (parseAttrs(el.attrs).Id === id ? '' : text)));
}

/** Adds a relationship from `source`, returning its id. */
function addRelationship(pkg: OoxmlPackage, source: string, type: string, target: string): string {
  const path = relsPathFor(source);
  const xml = pkg.text(path) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"></Relationships>`;
  const rels = elements(xml, 'Relationship');
  const ids = new Set(rels.map((r) => parseAttrs(r.attrs).Id));
  let n = rels.length + 1;
  while (ids.has(`rId${n}`)) n++;
  const close = xml.lastIndexOf('</');
  pkg.setText(path, `${xml.slice(0, close)}<${rels[0]?.qname ?? 'Relationship'} Id="rId${n}" Type="${type}" Target="${escapeXml(target)}"/>${xml.slice(close)}`);
  return `rId${n}`;
}

// ---------------------------------------------------------------------------
// The changes

function addSheet(pkg: OoxmlPackage, model: WorkbookModel, book: Book, op: Extract<SheetOp, { type: 'addSheet' }>): void {
  const wbDir = book.path.slice(0, book.path.lastIndexOf('/') + 1);
  const first = model.sheets.find((m) => m.info.kind === 'worksheet' && m.path);
  const dir = first ? first.path.slice(0, first.path.lastIndexOf('/') + 1) : `${wbDir}worksheets/`;
  const taken = new Set(pkg.list().map((p) => p.toLowerCase()));
  let k = 1;
  while (taken.has(`${dir}sheet${k}.xml`.toLowerCase())) k++;
  const path = `${dir}sheet${k}.xml`;

  // Rows as tall as the workbook's other sheets have them, in points.
  const height = ((first?.info.defaultRowHeightPx ?? 20) * 3) / 4;
  pkg.setText(
    path,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<worksheet xmlns="${MAIN_NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="${height}"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
  );
  const rid = addRelationship(pkg, book.path, WORKSHEET_REL, path.startsWith(wbDir) ? path.slice(wbDir.length) : `/${path}`);
  rewritePart(pkg, '[Content_Types].xml', (xml) => {
    const close = xml.lastIndexOf('</');
    const q = firstElement(xml, 'Override')?.qname ?? 'Override';
    return `${xml.slice(0, close)}<${q} PartName="/${path}" ContentType="${WORKSHEET_TYPE}"/>${xml.slice(close)}`;
  });

  const sample = book.entries[0];
  const idKey = Object.keys(sample?.attrs ?? {}).find((k) => k.includes(':') && localName(k) === 'id') ?? 'r:id';
  const qname = sample ? sample.text.slice(1, sample.text.search(/[\s/>]/)) : 'sheet';
  const sheetId = Math.max(0, ...book.entries.map((e) => e.sheetId)) + 1;
  const entries = book.entries.map((e) => e.text);
  entries.splice(op.at, 0, `<${qname} name="${escapeXml(op.name)}" sheetId="${sheetId}" ${idKey}="${rid}"/>`);
  const map = sheetIndexMap(op, model.sheets.length);
  writeBook(
    pkg,
    book,
    entries,
    (formula, local) => ({ formula, local: local === undefined ? undefined : (map[local] ?? local) }),
    (active, firstTab) => [map[active] ?? active, map[firstTab] ?? firstTab],
  );
  editTitles(pkg, (groups) => {
    const worksheets = groups.find((g) => model.sheets.some((m) => m.info.kind === 'worksheet' && g.items.includes(m.info.name)));
    if (!worksheets) return;
    worksheets.items.push(op.name);
    const order = model.sheets.map((m) => m.info.name);
    order.splice(op.at, 0, op.name);
    orderTitles(groups, order);
  });
}

function renameSheet(pkg: OoxmlPackage, model: WorkbookModel, book: Book, op: Extract<SheetOp, { type: 'renameSheet' }>): void {
  const old = model.sheets[op.sheet].info.name;
  if (old === op.name) return;
  const c: RefChange = { sheet: old.toLowerCase(), to: op.name, order: [], tables: new Set() };
  followNames(pkg, model, c, -1);
  const entries = book.entries.map((e) => e.text);
  entries[op.sheet] = retag(elements(entries[op.sheet], 'sheet')[0], { name: op.name });
  writeBook(
    pkg,
    book,
    entries,
    (formula, local) => ({ formula: changeRefs(formula, c), local }),
    (active, first) => [active, first],
  );
  editTitles(pkg, (groups) => {
    for (const g of groups) g.items = g.items.map((x) => (x === old ? op.name : onSheet(x, old) ? `${quoteSheetName(op.name)}${x.slice(x.indexOf('!', old.length))}` : x));
  });
}

function deleteSheet(pkg: OoxmlPackage, model: WorkbookModel, book: Book, d: number): { broken: number } {
  const m = model.sheets[d];
  const name = m.info.name;
  const wbRels = pkg.rels(book.path);
  const sheetRels = m.path ? pkg.rels(m.path) : [];
  // Slicers and timelines would be left pointing at pivot tables and tables that had gone.
  const slicers = wbRels.some((r) => r.type === 'slicerCache' || r.type === 'timelineCache');
  if (sheetRels.some((r) => r.type === 'slicer' || r.type === 'timeline' || (slicers && (r.type === 'pivotTable' || r.type === 'table')))) {
    throw new TextError(msg('sheets.refuse.slicers', { sheet: name }));
  }

  // The sheet's part, and everything only it used: drawings, charts, pictures, notes, tables, pivot tables.
  removeRelationship(pkg, book.path, book.entries[d].rid);
  if (m.path) {
    const kept = reachable(pkg, ['']);
    removeParts(pkg, new Set([...reachable(pkg, [m.path])].filter((p) => !kept.has(p))));
  }

  const order = model.sheets.map((s) => s.info.name);
  const c: RefChange = { sheet: name.toLowerCase(), to: null, order, tables: new Set(m.info.tables.map((t) => t.name.toLowerCase())) };
  const broken = followNames(pkg, model, c, d);

  const states = model.sheets.filter((_s, i) => i !== d).map((s) => s.info);
  const wasActive = activeTab(book) === d;
  const entries = book.entries.filter((_e, i) => i !== d).map((e) => e.text);
  // A sheet's position before, from its position after.
  const before = (i: number) => (i >= d ? i + 1 : i);
  let active = 0;
  writeBook(
    pkg,
    book,
    entries,
    (formula, local) => {
      // Names local to the sheet go with it, as Excel's print areas and titles do.
      if (local === d) return null;
      return { formula: changeRefs(formula, c), local: local !== undefined && local > d ? local - 1 : local };
    },
    (at, first) => {
      active = at === d ? nearestVisible(states, d) : at > d ? at - 1 : at;
      return [active, first > d ? first - 1 : Math.min(first, states.length - 1)];
    },
  );
  if (wasActive) setTabSelected(pkg, model.sheets[before(active)].path, true);
  rewritePart(pkg, book.path, (xml) =>
    replaceElements(xml, 'customWorkbookView', (el, text) =>
      Number(parseAttrs(el.attrs).activeSheetId) === book.entries[d].sheetId ? retag(el, { activeSheetId: String(book.entries[before(active)].sheetId) }) : text,
    ),
  );
  editTitles(pkg, (groups) => {
    for (const g of groups) g.items = g.items.filter((x) => x !== name && !onSheet(x, name));
  });
  return { broken };
}

function moveSheet(pkg: OoxmlPackage, model: WorkbookModel, book: Book, op: Extract<SheetOp, { type: 'moveSheet' }>): void {
  const map = sheetIndexMap(op, model.sheets.length);
  const entries: string[] = [];
  book.entries.forEach((e, i) => {
    entries[map[i]!] = e.text;
  });
  writeBook(
    pkg,
    book,
    entries,
    (formula, local) => ({ formula, local: local === undefined ? undefined : (map[local] ?? local) }),
    (active, first) => [map[active] ?? active, first],
  );
  const order: string[] = [];
  model.sheets.forEach((s, i) => {
    order[map[i]!] = s.info.name;
  });
  editTitles(pkg, (groups) => orderTitles(groups, order));
}

function hideSheet(pkg: OoxmlPackage, model: WorkbookModel, book: Book, op: Extract<SheetOp, { type: 'hideSheet' }>): void {
  const info = model.sheets[op.sheet].info;
  if (op.hidden === (info.state !== 'visible')) return;
  const entries = book.entries.map((e) => e.text);
  entries[op.sheet] = retag(elements(entries[op.sheet], 'sheet')[0], { state: op.hidden ? 'hidden' : null });
  const states = model.sheets.map((s, i): Pick<SheetInfo, 'state'> => (i === op.sheet ? { state: op.hidden ? 'hidden' : 'visible' } : s.info));
  const was = activeTab(book);
  let active = was;
  writeBook(
    pkg,
    book,
    entries,
    (formula, local) => ({ formula, local }),
    (at, first) => {
      // Excel opens on a visible sheet: the next one along from a sheet hidden.
      active = op.hidden && at === op.sheet ? nearestVisible(states, at) : at;
      return [active, first];
    },
  );
  if (op.hidden) setTabSelected(pkg, model.sheets[op.sheet].path, false);
  if (active !== was) setTabSelected(pkg, model.sheets[active]?.path, true);
}
