import type { Person } from './comments';
import { newGuid } from './commentEdits';
import { cellKey, MAX_COLS, MAX_ROWS, parseRange, toA1 } from '@ruled/formula/refs';
import type { CellNote, SheetInfo } from './types';
import type { WorkbookModel } from './workbook';
import { decodeEntities, decodeOoxmlText, elements, encodeOoxmlText, escapeXml, findTagEnd, firstElement, parseAttrs, replaceElements, type ElementSpan } from './xml';
import { relsPathFor, type OoxmlPackage, type Relationship } from './zip';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const THREADS_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const REVISION_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const EXCEL_NS = 'urn:schemas-microsoft-com:office:excel';
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const REL_TYPE = {
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  vmlDrawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
  threadedComment: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
  person: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
};
const CONTENT_TYPE = {
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  threads: 'application/vnd.ms-excel.threadedcomments+xml',
  persons: 'application/vnd.ms-excel.person+xml',
};
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * What Excel puts in the note it keeps for each threaded comment, which is
 * what versions of Excel from before threads show. Excel writes it in English.
 */
const PLACEHOLDER =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n';

/**
 * Writes changed notes and threaded comments into a workbook package. On each
 * sheet with changes, its comments part, the VML drawing that gives each note
 * its box, its threaded comments part and the workbook's list of people are
 * updated where they exist, and created with the relationships and content
 * types Excel needs where they don't. Parts left with nothing in them are
 * removed. Comments on cells that didn't change are kept as written.
 *
 * `edits` has, per sheet, each changed cell (by cellKey) and what it holds
 * now; `people` names everyone whose comments may be new to the file.
 */
export function writeComments(pkg: OoxmlPackage, model: WorkbookModel, edits: Map<number, CellNote | null>[], people: ReadonlyMap<string, Person>): void {
  const writers = new Set<string>();
  edits.forEach((cells, s) => {
    const m = model.sheets[s];
    if (!cells.size || !m?.path || m.info.kind !== 'worksheet' || pkg.text(m.path) === undefined) return;
    writeSheet(pkg, model.workbookPath, m.path, m.scan?.prefix ?? '', m.info, cells);
    for (const note of cells.values()) if (note?.kind === 'thread') for (const e of note.entries) if (e.personId) writers.add(e.personId);
  });
  const known = [...writers].flatMap((id) => people.get(id) ?? []);
  if (known.length) addPeople(pkg, model.workbookPath, known);
}

function writeSheet(pkg: OoxmlPackage, workbookPath: string, path: string, p: string, info: SheetInfo, cells: Map<number, CellNote | null>): void {
  const dir = workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1);
  const rels = pkg.rels(path).filter((r) => !r.external);
  const noted = [...cells].filter((e): e is [number, CellNote] => !!e[1]);

  const threadsRel = rels.find((r) => r.type === 'threadedComment');
  const threadsXml = threadsRel ? pkg.text(threadsRel.target) : undefined;
  if (threadsRel && threadsXml !== undefined) {
    const next = rewriteThreads(threadsXml, cells);
    if (next !== threadsXml) pkg.setText(threadsRel.target, next);
  } else if (noted.some(([, n]) => n.kind === 'thread')) {
    const part = freePart(pkg, (n) => `${dir}threadedComments/threadedComment${n}.xml`);
    pkg.setText(part, rewriteThreads(`${XML_HEAD}<ThreadedComments xmlns="${THREADS_NS}" xmlns:x="${MAIN_NS}"></ThreadedComments>`, cells));
    addRel(pkg, path, REL_TYPE.threadedComment, part);
    setOverride(pkg, part, CONTENT_TYPE.threads);
  }

  // Notes, and the note Excel keeps for each thread.
  const notesRel = rels.find((r) => r.type === 'comments');
  const notesXml = notesRel ? pkg.text(notesRel.target) : undefined;
  if (notesRel && notesXml !== undefined) {
    const next = rewriteNotes(notesXml, cells);
    if (next !== notesXml) pkg.setText(notesRel.target, next);
  } else if (noted.length) {
    const part = freePart(pkg, (n) => `${dir}comments${n}.xml`);
    const empty = `${XML_HEAD}<comments xmlns="${MAIN_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="xr" xmlns:xr="${REVISION_NS}"><authors></authors><commentList></commentList></comments>`;
    pkg.setText(part, rewriteNotes(empty, cells));
    addRel(pkg, path, REL_TYPE.comments, part);
    setOverride(pkg, part, CONTENT_TYPE.comments);
  }

  // Each note's box, in the drawing the sheet's <legacyDrawing> names.
  let sheetXml = pkg.text(path)!;
  const from = tailStart(sheetXml, p);
  const legacy = firstElement(sheetXml, 'legacyDrawing', from);
  const legacyId = legacy ? attrNamed(parseAttrs(legacy.attrs), 'id') : undefined;
  const vmlRel = legacyId !== undefined ? rels.find((r) => r.id === legacyId) : undefined;
  const vmlXml = vmlRel ? pkg.text(vmlRel.target) : undefined;
  if (legacy && vmlRel && vmlXml !== undefined) {
    const next = rewriteVml(vmlXml, cells, info, pkg);
    if (next !== vmlXml) pkg.setText(vmlRel.target, next);
  } else if (noted.length) {
    // A <legacyDrawing> whose drawing is missing is replaced.
    if (legacy) sheetXml = sheetXml.slice(0, legacy.start) + sheetXml.slice(legacy.end);
    const part = freePart(pkg, (n) => `${dir}drawings/vmlDrawing${n}.vml`);
    const block = freeBlock(pkg);
    const empty = `<xml xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}" xmlns:x="${EXCEL_NS}"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${block}"/></o:shapelayout></xml>`;
    pkg.setText(part, rewriteVml(empty, cells, info, pkg));
    const id = addRel(pkg, path, REL_TYPE.vmlDrawing, part);
    if (!hasDefault(pkg, 'vml')) addContentType(pkg, `<Default Extension="vml" ContentType="${CONTENT_TYPE.vml}"/>`);
    pkg.setText(path, addLegacyDrawing(sheetXml, p, id));
  }
  removeEmptyNoteParts(pkg, path, p);
}

/**
 * Removes a sheet's notes and threads parts once there is nothing left in
 * them, and then the drawing that gives notes their boxes, and its
 * <legacyDrawing>, once nothing else (no form control) is drawn there
 * either, as Excel does. `p` is the sheet's prefix for its elements.
 */
export function removeEmptyNoteParts(pkg: OoxmlPackage, path: string, p: string): void {
  const rels = pkg.rels(path).filter((r) => !r.external);
  const threads = rels.find((r) => r.type === 'threadedComment');
  const threadsXml = threads && pkg.text(threads.target);
  if (threads && threadsXml !== undefined && !firstElement(threadsXml, 'threadedComment')) removePart(pkg, path, threads);
  const notes = rels.find((r) => r.type === 'comments');
  const notesXml = notes && pkg.text(notes.target);
  if (notes && notesXml !== undefined && !firstElement(notesXml, 'comment')) removePart(pkg, path, notes);
  else if (notesXml !== undefined || !rels.some((r) => r.type === 'vmlDrawing')) return;
  const sheetXml = pkg.text(path);
  const legacy = sheetXml === undefined ? undefined : firstElement(sheetXml, 'legacyDrawing', tailStart(sheetXml, p));
  const id = legacy && attrNamed(parseAttrs(legacy.attrs), 'id');
  const vml = id === undefined ? undefined : rels.find((r) => r.id === id);
  const vmlXml = vml && pkg.text(vml.target);
  if (!sheetXml || !legacy || !vml || vmlXml === undefined || elements(vmlXml, 'shape').length) return;
  removePart(pkg, path, vml);
  pkg.setText(path, sheetXml.slice(0, legacy.start) + sheetXml.slice(legacy.end));
}

// ---------------------------------------------------------------------------
// Threaded comments

/**
 * The threads part with each changed cell's thread written in place of the
 * one it had, or added at the end. A comment that is unchanged keeps its
 * element as written; one whose text changed keeps its attributes.
 */
function rewriteThreads(xml: string, cells: Map<number, CellNote | null>): string {
  const p = rootOf(xml).prefix;
  const found = elements(xml, 'threadedComment');
  const byId = new Map<string, ElementSpan>();
  for (const el of found) {
    const id = parseAttrs(el.attrs).id;
    if (id) byId.set(id, el);
  }
  let out = '';
  let pos = 0;
  const written = new Set<number>();
  for (const el of found) {
    const key = keyOfRef(parseAttrs(el.attrs).ref);
    if (key === undefined || !cells.has(key)) continue;
    out += xml.slice(pos, el.start);
    pos = el.end;
    if (written.has(key)) continue;
    written.add(key);
    out += threadXml(cells.get(key) ?? null, key, xml, byId, p);
  }
  out += xml.slice(pos);
  const added = [...cells]
    .filter(([key]) => !written.has(key))
    .map(([key, note]) => threadXml(note, key, xml, byId, p))
    .join('');
  return added ? appendToRoot(out, added) : out;
}

function threadXml(note: CellNote | null, key: number, xml: string, byId: Map<string, ElementSpan>, p: string): string {
  if (note?.kind !== 'thread' || !note.entries.length) return '';
  const ref = toA1(Math.floor(key / MAX_COLS), key % MAX_COLS);
  const rootId = note.entries[0].id ?? newGuid();
  return note.entries
    .map((e, k) => {
      const id = e.id ?? (k ? newGuid() : rootId);
      const el = e.id ? byId.get(e.id) : undefined;
      const was = el ? parseAttrs(el.attrs) : {};
      // Excel's order: ref, dT, personId, id, parentId, done.
      const want: [string, string | undefined][] = [
        ['ref', ref],
        ['dT', e.time ?? was.dT],
        ['personId', e.personId ?? was.personId],
        ['id', id],
        ['parentId', k ? rootId : undefined],
        ['done', !k && note.resolved ? '1' : undefined],
      ];
      const inner = el?.inner ?? '';
      const textEl = firstElement(inner, 'text');
      const sameText = !!el && decodeOoxmlText(decodeEntities(textEl?.inner ?? '')) === e.text;
      const sameAttrs = want.every(([name, value]) => (name === 'done' ? (was.done === '1' || was.done === 'true') === !!value : was[name] === value));
      if (el && sameText && sameAttrs) return xml.slice(el.start, el.end);
      // Attributes of the element as written, in its order, then any it lacked.
      const attrs: [string, string][] = [];
      for (const [name, value] of Object.entries(was)) {
        const w = want.find(([n]) => n === name);
        if (!w) attrs.push([name, value]);
        else if (w[1] !== undefined) attrs.push([name, w[1]]);
      }
      for (const [name, value] of want) if (value !== undefined && !(name in was)) attrs.push([name, value]);
      const qname = el?.qname ?? `${p}threadedComment`;
      const tag = `<${qname}${attrs.map(([n, v]) => ` ${n}="${escapeXml(v)}"`).join('')}>`;
      // Mentions point into the text by position, so they go when it changes; anything else stays.
      const body = sameText ? inner : `<${textEl?.qname ?? `${p}text`}>${escapeXml(encodeOoxmlText(e.text))}</${textEl?.qname ?? `${p}text`}>${elementsText(inner, 'extLst')}`;
      return `${tag}${body}</${qname}>`;
    })
    .join('');
}

function addPeople(pkg: OoxmlPackage, workbookPath: string, people: Person[]): void {
  const rel = pkg.rels(workbookPath).find((r) => r.type === 'person' && !r.external);
  const xml = rel ? pkg.text(rel.target) : undefined;
  const have = new Set(xml ? elements(xml, 'person').map((el) => parseAttrs(el.attrs).id) : []);
  const add = people.filter((x) => !have.has(x.id));
  if (!add.length) return;
  const p = xml ? rootOf(xml).prefix : '';
  const text = add
    .map((x) => `<${p}person displayName="${escapeXml(x.displayName)}" id="${x.id}" userId="${escapeXml(x.userId ?? x.displayName)}" providerId="${escapeXml(x.providerId ?? 'None')}"/>`)
    .join('');
  if (rel && xml !== undefined) {
    pkg.setText(rel.target, appendToRoot(xml, text));
    return;
  }
  const dir = workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1);
  const part = freePart(pkg, (n) => `${dir}persons/person${n === 1 ? '' : n}.xml`);
  pkg.setText(part, `${XML_HEAD}<personList xmlns="${THREADS_NS}" xmlns:x="${MAIN_NS}">${text}</personList>`);
  addRel(pkg, workbookPath, REL_TYPE.person, part);
  setOverride(pkg, part, CONTENT_TYPE.persons);
}

// ---------------------------------------------------------------------------
// Notes

/**
 * The comments part with each changed cell's note written in place of the
 * one it had, removed, or added among the others in reading order. Threads
 * get the note Excel writes for them, which lists the thread's comments.
 */
function rewriteNotes(xml: string, cells: Map<number, CellNote | null>): string {
  const root = rootOf(xml);
  const p = root.prefix;
  const xr = prefixFor(root.attrs, REVISION_NS);
  const authorsEl = firstElement(xml, 'authors');
  const authors = authorsEl?.inner ? elements(authorsEl.inner, 'author').map((a) => decodeEntities(a.inner ?? '')) : [];
  const added: string[] = [];
  const authorId = (name: string) => {
    let i = authors.indexOf(name);
    if (i < 0) {
      i = authors.push(name) - 1;
      added.push(name);
    }
    return i;
  };

  const listEl = firstElement(xml, 'commentList');
  const inner = listEl?.inner ?? '';
  const list: { key: number; text: string }[] = [];
  const seen = new Set<number>();
  for (const el of elements(inner, 'comment')) {
    const key = keyOfRef(parseAttrs(el.attrs).ref) ?? -1;
    if (!cells.has(key)) list.push({ key, text: inner.slice(el.start, el.end) });
    else if (!seen.has(key)) {
      seen.add(key);
      const note = cells.get(key);
      if (note) list.push({ key, text: noteXml(note, key, p, xr, authorId) });
    }
  }
  for (const [key, note] of [...cells].sort((a, b) => a[0] - b[0])) {
    if (seen.has(key) || !note) continue;
    const text = noteXml(note, key, p, xr, authorId);
    const at = list.findIndex((c) => c.key > key);
    list.splice(at < 0 ? list.length : at, 0, { key, text });
  }

  const listText = `<${listEl?.qname ?? `${p}commentList`}${listEl?.attrs ?? ''}>${list.map((c) => c.text).join('')}</${listEl?.qname ?? `${p}commentList`}>`;
  const authorsText = added.length
    ? `<${authorsEl?.qname ?? `${p}authors`}${authorsEl?.attrs ?? ''}>${authorsEl?.inner ?? ''}${added.map((a) => `<${p}author>${escapeXml(a)}</${p}author>`).join('')}</${authorsEl?.qname ?? `${p}authors`}>`
    : undefined;
  // Both are replaced from the end, so the earlier offsets stay right.
  let out = xml;
  if (listEl) out = out.slice(0, listEl.start) + listText + out.slice(listEl.end);
  else if (authorsEl) out = out.slice(0, authorsEl.end) + listText + out.slice(authorsEl.end);
  else out = appendToRoot(out, listText);
  if (authorsText) {
    out = authorsEl ? out.slice(0, authorsEl.start) + authorsText + out.slice(authorsEl.end) : insertAfterRootTag(out, authorsText);
  }
  return out;
}

function noteXml(note: CellNote, key: number, p: string, xr: string | undefined, authorId: (name: string) => number): string {
  const ref = toA1(Math.floor(key / MAX_COLS), key % MAX_COLS);
  const t = (text: string) => `<${p}t${/^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : ''}>${escapeXml(encodeOoxmlText(text))}</${p}t>`;
  if (note.kind === 'thread') {
    const [first, ...replies] = note.entries;
    const text = `${PLACEHOLDER}Comment:\n    ${first?.text ?? ''}${replies.map((r) => `\nReply:\n    ${r.text}`).join('')}`;
    const uid = xr && first?.id ? ` ${xr}:uid="${escapeXml(first.id)}"` : '';
    return `<${p}comment ref="${ref}" authorId="${authorId(`tc=${first?.id ?? ''}`)}" shapeId="0"${uid}><${p}text>${t(text)}</${p}text></${p}comment>`;
  }
  const { author, text } = note.entries[0] ?? { text: '' };
  // As Excel writes a note: its author's name in bold, then the text on the next line.
  const run = (bold: boolean, s: string) =>
    `<${p}r><${p}rPr>${bold ? `<${p}b/>` : ''}<${p}sz val="9"/><${p}color indexed="81"/><${p}rFont val="Tahoma"/><${p}family val="2"/></${p}rPr>${t(s)}</${p}r>`;
  const body = author ? run(true, `${author}:`) + run(false, `\n${text}`) : run(false, text);
  return `<${p}comment ref="${ref}" authorId="${authorId(author ?? '')}" shapeId="0"><${p}text>${body}</${p}text></${p}comment>`;
}

// ---------------------------------------------------------------------------
// Note boxes (VML)

/** Excel's shape type for a note's box. */
const NOTE_SHAPETYPE = (v: Prefix, o: Prefix) =>
  `<${v.prefix}shapetype id="_x0000_t202"${v.declare}${o.declare} coordsize="21600,21600" ${o.prefix}spt="202" path="m,l,21600r21600,l21600,xe"><${v.prefix}stroke joinstyle="miter"/><${v.prefix}path gradientshapeok="t" ${o.prefix}connecttype="rect"/></${v.prefix}shapetype>`;

/** The drawing with the boxes of notes on changed cells that are gone removed, and boxes added for new ones, hidden as Excel adds them. */
function rewriteVml(xml: string, cells: Map<number, CellNote | null>, info: SheetInfo, pkg: OoxmlPackage): string {
  const root = rootOf(xml);
  const declared = (uri: string, fallback: string) => {
    const prefix = prefixFor(root.attrs, uri);
    return { prefix: `${prefix ?? fallback}:`, declare: prefix === undefined ? ` xmlns:${fallback}="${uri}"` : '' };
  };
  const v = declared(VML_NS, 'v');
  const o = declared(OFFICE_NS, 'o');
  const x = declared(EXCEL_NS, 'x');

  const boxed = new Set<number>();
  let top = 0;
  let out = replaceElements(xml, 'shape', (el, text) => {
    const n = /^_x0000_s(\d+)$/.exec(parseAttrs(el.attrs).id ?? '');
    if (n) top = Math.max(top, Number(n[1]));
    const key = noteBoxKey(el.inner ?? '');
    if (key === undefined || !cells.has(key)) return text;
    if (cells.get(key) && !boxed.has(key)) {
      boxed.add(key);
      return text;
    }
    return '';
  });

  const wanted = [...cells].filter(([key, note]) => note && !boxed.has(key)).map(([key]) => key);
  if (!wanted.length) return out;
  const idmap = firstElement(out, 'idmap');
  const blocks = (idmap ? (parseAttrs(idmap.attrs).data ?? '') : '')
    .split(',')
    .map((b) => Number(b.trim()))
    .filter((b) => Number.isInteger(b) && b > 0);
  if (!blocks.length) blocks.push(freeBlock(pkg));
  let id = Math.max(top, blocks[0] * 1024);
  let z = elements(out, 'shape').length;
  const shapes = wanted.map((key) => {
    id++;
    const block = Math.floor(id / 1024);
    if (!blocks.includes(block)) blocks.push(block);
    return noteBox(key, id, ++z, info, v, o, x);
  });
  const layout = `<${o.prefix}shapelayout ${v.prefix}ext="edit"${o.declare}${v.declare}><${o.prefix}idmap ${v.prefix}ext="edit" data="${blocks.join(',')}"/></${o.prefix}shapelayout>`;
  if (idmap) {
    const attrs = { ...parseAttrs(idmap.attrs), data: blocks.join(',') };
    out = out.slice(0, idmap.start) + `<${idmap.qname}${Object.entries(attrs).map(([n, value]) => ` ${n}="${escapeXml(value)}"`).join('')}/>` + out.slice(idmap.end);
  } else {
    out = insertAfterRootTag(out, layout);
  }
  const typeEl = elements(out, 'shapetype').find((el) => parseAttrs(el.attrs).id === '_x0000_t202');
  return appendToRoot(out, (typeEl ? '' : NOTE_SHAPETYPE(v, o)) + shapes.join(''));
}

/** The cell a VML shape is the box of, if it is a note's. */
function noteBoxKey(inner: string): number | undefined {
  const data = firstElement(inner, 'ClientData');
  if (!data?.inner || parseAttrs(data.attrs).ObjectType !== 'Note') return undefined;
  const row = Number(firstElement(data.inner, 'Row')?.inner?.trim() ?? NaN);
  const col = Number(firstElement(data.inner, 'Column')?.inner?.trim() ?? NaN);
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0 ? cellKey(row, col) : undefined;
}

type Prefix = { prefix: string; declare: string };

/** A note's box as Excel adds one: to the right of its cell, two columns wide and four rows high, and hidden until the cell is pointed at. */
function noteBox(key: number, id: number, z: number, info: SheetInfo, v: Prefix, o: Prefix, x: Prefix): string {
  const row = Math.floor(key / MAX_COLS);
  const col = key % MAX_COLS;
  const top = Math.max(0, row - 1);
  const anchor = [Math.min(col + 1, MAX_COLS - 1), 15, top, 2, Math.min(col + 3, MAX_COLS - 1), 15, Math.min(top + 4, MAX_ROWS - 1), 16];
  // Where that is in points, which Excel reads when the anchor can't be used.
  const left = (offset(col + 1, info.colWidths, info.defaultColWidthPx) + 15) * 0.75;
  const down = (offset(top, info.rowHeights, info.defaultRowHeightPx) + 2) * 0.75;
  const [vp, op, xp] = [v.prefix, o.prefix, x.prefix];
  return (
    `<${vp}shape id="_x0000_s${id}" type="#_x0000_t202"${v.declare}${o.declare}${x.declare} style="position:absolute;margin-left:${left}pt;margin-top:${down}pt;width:108pt;height:59.25pt;z-index:${z};visibility:hidden" fillcolor="#ffffe1" ${op}insetmode="auto">` +
    `<${vp}fill color2="#ffffe1"/><${vp}shadow on="t" color="black" obscured="t"/><${vp}path ${op}connecttype="none"/>` +
    `<${vp}textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></${vp}textbox>` +
    `<${xp}ClientData ObjectType="Note"><${xp}MoveWithCells/><${xp}SizeWithCells/><${xp}Anchor>${anchor.join(', ')}</${xp}Anchor><${xp}AutoFill>False</${xp}AutoFill><${xp}Row>${row}</${xp}Row><${xp}Column>${col}</${xp}Column></${xp}ClientData></${vp}shape>`
  );
}

/** Pixels from the sheet's edge to the start of row or column n, given the custom sizes. */
function offset(n: number, sizes: Record<number, number>, size: number): number {
  let px = n * size;
  for (const [k, v] of Object.entries(sizes)) if (Number(k) < n) px += v - size;
  return px;
}

/** A block of shape ids (o:idmap) that no VML drawing in the package uses: Excel numbers note boxes 1024 to a block. */
function freeBlock(pkg: OoxmlPackage): number {
  let max = 0;
  for (const path of pkg.list()) {
    if (!/\.vml$/i.test(path)) continue;
    const idmap = firstElement(pkg.text(path) ?? '', 'idmap');
    for (const b of (idmap ? (parseAttrs(idmap.attrs).data ?? '') : '').split(',')) if (Number(b) > max) max = Number(b);
  }
  return max + 1;
}

// ---------------------------------------------------------------------------
// The worksheet

/** Where the elements after <sheetData> start, found without reading the cells. */
function tailStart(xml: string, p: string): number {
  const close = xml.lastIndexOf(`</${p}sheetData>`);
  if (close >= 0) return close;
  const empty = xml.lastIndexOf(`<${p}sheetData/>`);
  return empty >= 0 ? empty : 0;
}

/** Adds <legacyDrawing r:id> where CT_Worksheet has it: after <drawing>, before the elements that follow it. */
function addLegacyDrawing(xml: string, p: string, id: string): string {
  const root = rootOf(xml);
  let r = prefixFor(root.attrs, REL_NS);
  let out = xml;
  if (r === undefined) {
    r = 'r';
    out = out.slice(0, root.tagEnd) + ` xmlns:r="${REL_NS}"` + out.slice(root.tagEnd);
  }
  const from = tailStart(out, p);
  const after = ['legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst', 'AlternateContent']
    .map((name) => firstElement(out, name, from)?.start ?? -1)
    .filter((at) => at >= 0);
  const at = after.length ? Math.min(...after) : out.lastIndexOf('</');
  return out.slice(0, at) + `<${p}legacyDrawing ${r}:id="${id}"/>` + out.slice(at);
}

// ---------------------------------------------------------------------------
// Parts, relationships and content types

interface Root {
  qname: string;
  /** "x:" for <x:comments>, else "". */
  prefix: string;
  attrs: Record<string, string>;
  /** Offset of the ">" (or "/>") that ends the start tag. */
  tagEnd: number;
  selfClosing: boolean;
}

/** The document element's start tag. */
function rootOf(xml: string): Root {
  let lt = xml.indexOf('<');
  // Past the declaration, comments and processing instructions.
  while (lt !== -1 && (xml.charCodeAt(lt + 1) === 63 || xml.charCodeAt(lt + 1) === 33)) {
    const end = xml.startsWith('<!--', lt) ? xml.indexOf('-->', lt) : xml.indexOf('>', lt);
    lt = end < 0 ? -1 : xml.indexOf('<', end);
  }
  if (lt < 0) return { qname: '', prefix: '', attrs: {}, tagEnd: xml.length, selfClosing: true };
  const gt = findTagEnd(xml, lt);
  const selfClosing = xml.charCodeAt(gt - 1) === 47;
  const inner = xml.slice(lt + 1, selfClosing ? gt - 1 : gt);
  const space = inner.search(/\s/);
  const qname = space < 0 ? inner : inner.slice(0, space);
  const colon = qname.indexOf(':');
  return {
    qname,
    prefix: colon < 0 ? '' : qname.slice(0, colon + 1),
    attrs: space < 0 ? {} : parseAttrs(inner.slice(space)),
    tagEnd: selfClosing ? gt - 1 : gt,
    selfClosing,
  };
}

/** The prefix the document element declares for a namespace. */
function prefixFor(attrs: Record<string, string>, uri: string): string | undefined {
  for (const [name, value] of Object.entries(attrs)) if (value === uri && name.startsWith('xmlns:')) return name.slice(6);
  return undefined;
}

/** An attribute by its local name, whatever its prefix (r:id). */
function attrNamed(attrs: Record<string, string>, name: string): string | undefined {
  for (const [k, value] of Object.entries(attrs)) if (k === name || k.endsWith(`:${name}`)) return value;
  return undefined;
}

function keyOfRef(ref: string | undefined): number | undefined {
  const at = parseRange(ref ?? '');
  return at ? cellKey(at.r1, at.c1) : undefined;
}

/** The text of every element called `name` in xml, one after another. */
function elementsText(xml: string, name: string): string {
  return elements(xml, name)
    .map((el) => xml.slice(el.start, el.end))
    .join('');
}

/** Puts text at the end of the document element. */
function appendToRoot(xml: string, text: string): string {
  const root = rootOf(xml);
  if (root.selfClosing) return `${xml.slice(0, root.tagEnd)}>${text}</${root.qname}>${xml.slice(root.tagEnd + 2)}`;
  const close = xml.lastIndexOf('</');
  return xml.slice(0, close) + text + xml.slice(close);
}

/** Puts text first in the document element. */
function insertAfterRootTag(xml: string, text: string): string {
  const root = rootOf(xml);
  if (root.selfClosing) return appendToRoot(xml, text);
  return xml.slice(0, root.tagEnd + 1) + text + xml.slice(root.tagEnd + 1);
}

/** The first name that `make` gives for 1, 2, 3… that no part has, ignoring case as packages do. */
function freePart(pkg: OoxmlPackage, make: (n: number) => string): string {
  const taken = new Set(pkg.list().map((x) => x.toLowerCase()));
  let n = 1;
  while (taken.has(make(n).toLowerCase())) n++;
  return make(n);
}

/** `to` as a relationship from `from` writes it, relative to the folder `from` is in. */
function relativeTarget(from: string, to: string): string {
  const a = from.split('/').slice(0, -1);
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...a.slice(i).map(() => '..'), ...b.slice(i)].join('/');
}

/** Adds a relationship from one part to another, returning its new id. */
export function addRel(pkg: OoxmlPackage, from: string, type: string, to: string): string {
  const path = relsPathFor(from);
  const xml = pkg.text(path) ?? `${XML_HEAD}<Relationships xmlns="${PACKAGE_RELS_NS}"></Relationships>`;
  const ids = new Set(elements(xml, 'Relationship').map((el) => parseAttrs(el.attrs).Id));
  let n = ids.size + 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const p = rootOf(xml).prefix;
  pkg.setText(path, appendToRoot(xml, `<${p}Relationship Id="${id}" Type="${type}" Target="${escapeXml(relativeTarget(from, to))}"/>`));
  return id;
}

/** Removes a part, the relationship to it from `from`, its own relationships and its content type. */
export function removePart(pkg: OoxmlPackage, from: string, rel: Relationship): void {
  const path = relsPathFor(from);
  const xml = pkg.text(path);
  if (xml !== undefined) pkg.setText(path, replaceElements(xml, 'Relationship', (el, text) => (parseAttrs(el.attrs).Id === rel.id ? '' : text)));
  pkg.remove(rel.target);
  pkg.remove(relsPathFor(rel.target));
  setOverride(pkg, rel.target, undefined);
}

/** Sets a part's content type, or (undefined) removes its override. */
export function setOverride(pkg: OoxmlPackage, part: string, type: string | undefined): void {
  const xml = pkg.text('[Content_Types].xml');
  if (xml === undefined) return;
  const name = `/${part}`.toLowerCase();
  const without = replaceElements(xml, 'Override', (el, text) => ((parseAttrs(el.attrs).PartName ?? '').toLowerCase() === name ? '' : text));
  pkg.setText('[Content_Types].xml', without);
  if (type) addContentType(pkg, `<Override PartName="/${escapeXml(part)}" ContentType="${type}"/>`);
}

function hasDefault(pkg: OoxmlPackage, extension: string): boolean {
  const xml = pkg.text('[Content_Types].xml') ?? '';
  return elements(xml, 'Default').some((el) => (parseAttrs(el.attrs).Extension ?? '').toLowerCase() === extension);
}

function addContentType(pkg: OoxmlPackage, element: string): void {
  const xml = pkg.text('[Content_Types].xml');
  if (xml === undefined) return;
  const p = rootOf(xml).prefix;
  pkg.setText('[Content_Types].xml', appendToRoot(xml, p ? element.replace(/^<(\w+)/, `<${p}$1`) : element));
}
