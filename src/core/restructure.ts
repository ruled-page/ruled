import { msg, TextError } from '../i18n/message';
import { inRect, MAX_COLS, MAX_ROWS, parseRange, rangeToA1, toA1 } from './refs';
import { refMapFor, splits, transformFormula, transformSqref, type FormulaContext, type LineOp, type RefMap, type StructuralOp } from './refTransform';
import type { Rect } from './types';
import type { SheetModel, WorkbookModel } from './workbook';
import { countElements, decodeEntities, elements, escapeXml, findTagEnd, replaceElements } from './xml';
import type { OoxmlPackage } from './zip';

/**
 * Applies a structural change — inserting or deleting rows or columns, or
 * moving cells — to a saved workbook package, rewriting every reference that
 * points at the cells that moved: formulas on every sheet, defined names,
 * merged cells, validation, conditional formatting, hyperlinks, filters,
 * tables, pivot sources, charts, drawings and notes.
 *
 * `model` must be read from `pkg` as it is now. Throws, changing nothing,
 * when the change would split something Excel can't split either. Returns
 * how many cell formulas now show #REF! because cells they used were deleted.
 */
export function restructure(pkg: OoxmlPackage, model: WorkbookModel, op: StructuralOp): { broken: number } {
  if (op.type === 'growTable') {
    growTable(pkg, model, op);
    return { broken: 0 };
  }
  check(model, op);
  const map = refMapFor(op);
  const names = model.sheets.map((m) => m.info.name);
  const byName = new Map(names.map((n, i) => [n.toLowerCase(), i]));
  const ctx: FormulaContext = { names, indexOf: (n) => byName.get(n.toLowerCase()) };
  const fx = (f: string, from: number, to = from) => transformFormula(f, from, to, map, ctx);
  let broken = 0;
  // For cell formulas, each rewritten once: count the ones that lose a reference.
  const cellFx = (f: string, from: number, to = from) => {
    const next = fx(f, from, to);
    if (next !== f && next.includes('#REF!') && !f.includes('#REF!')) broken++;
    return next;
  };
  // On sheets whose cells stay put, only formulas naming a sheet that changes can change.
  const moving = [...map.sheets].map((s) => names[s].toLowerCase());
  const mentions = (f: string) => {
    const lower = f.toLowerCase();
    return moving.some((n) => lower.includes(n));
  };

  const extra = tableChanges(pkg, model, op, map, fx);
  const incoming = new Map<number, CellOut[]>();
  // Moving cells to another sheet: take them from the source first.
  if (op.type === 'move' && op.toSheet !== op.sheet) {
    const src = model.sheets[op.sheet];
    incoming.set(op.toSheet, collect(src).cells.filter((c) => inRect(op.src, c.row, c.col)).map((c) => place(c, op.sheet, op.toSheet, map, cellFx, src)!));
  }
  model.sheets.forEach((m, s) => {
    if (m.xml === undefined || !m.scan) return;
    let xml = rewriteCells(m, s, op, map, cellFx, mentions, extra.get(s) ?? [], incoming.get(s) ?? []);
    xml = rewriteSheetMeta(xml, s, op, map, fx);
    if (xml !== m.xml) pkg.setText(m.path, xml);
  });

  const wb = pkg.text(model.workbookPath);
  if (wb) {
    pkg.setText(
      model.workbookPath,
      replaceElements(wb, 'definedName', ({ qname, attrs, inner }, all) => {
        if (inner === null) return all;
        const local = /\blocalSheetId="(\d+)"/.exec(attrs);
        const next = fx(decodeEntities(inner), local ? Number(local[1]) : -1);
        return next === decodeEntities(inner) ? all : `<${qname}${attrs}>${escapeXml(next)}</${qname}>`;
      }),
    );
  }

  for (const path of pkg.list()) {
    if (/^xl\/charts\/chart\w*\.xml$/.test(path)) rewriteText(pkg, path, (xml) => rewriteFormulaElements(xml, ['f'], (f) => fx(f, -1)));
    if (/pivotCacheDefinition\w*\.xml$/.test(path)) rewriteText(pkg, path, (xml) => pivotSource(xml, map, ctx));
  }
  model.sheets.forEach((m, s) => {
    if (!map.sheets.has(s) || !m.path) return;
    for (const r of pkg.rels(m.path)) {
      if (r.external) continue;
      if (r.type === 'pivotTable') rewriteText(pkg, r.target, (xml) => rewriteRefAttrs(xml, ['location'], s, map));
      if (r.type === 'comments') rewriteText(pkg, r.target, (xml) => comments(xml, s, map));
      if (r.type === 'threadedComment') rewriteText(pkg, r.target, (xml) => threads(xml, s, map));
      if (r.type === 'vmlDrawing') rewriteText(pkg, r.target, (xml) => vmlNotes(xml, s, map));
      if (r.type === 'drawing' && op.type !== 'move') rewriteText(pkg, r.target, (xml) => anchors(xml, op));
    }
  });
  return { broken };
}

function rewriteText(pkg: OoxmlPackage, path: string, fn: (xml: string) => string): void {
  const xml = pkg.text(path);
  if (xml === undefined) return;
  const next = fn(xml);
  if (next !== xml) pkg.setText(path, next);
}

// ---------------------------------------------------------------------------
// Refusals

/**
 * An Excel table takes in the rows just below it: its range, and its
 * filter's, reach that much further. Formulas that use the table see the
 * new rows when the workbook is read again.
 */
function growTable(pkg: OoxmlPackage, model: WorkbookModel, op: Extract<StructuralOp, { type: 'growTable' }>): void {
  const info = model.sheets[op.sheet]?.info;
  const table = info?.tables.find((t) => t.name === op.table);
  if (!info || info.kind !== 'worksheet' || !table?.path) throw new TextError(msg('core.structure.notWorksheet'));
  if (info.protected) throw new TextError(msg('core.reason.protected'));
  if (table.totalsRows || table.rect.r2 + op.rows >= MAX_ROWS) throw new TextError(msg('core.structure.noRoomBelow'));
  const xml = pkg.text(table.path);
  if (!xml) return;
  const grow = (_all: string, a: string, ref: string, b: string) => {
    const r = parseRange(ref);
    return r ? `${a}${rangeToA1({ ...r, r2: r.r2 + op.rows })}${b}` : _all;
  };
  pkg.setText(
    table.path,
    xml.replace(/(<(?:\w+:)?table\b[^>]*?\sref=")([^"]*)(")/, grow).replace(/(<(?:\w+:)?autoFilter\b[^>]*?\sref=")([^"]*)(")/, grow),
  );
}

function check(model: WorkbookModel, op: StructuralOp): void {
  const m = model.sheets[op.sheet];
  if (!m || m.info.kind !== 'worksheet') throw new TextError(msg('core.structure.notWorksheet'));
  if (m.info.protected) throw new TextError(msg('core.reason.protected'));
  const sheets = op.type === 'move' ? [op.sheet, op.toSheet] : [op.sheet];
  for (const s of sheets) {
    const info = model.sheets[s]?.info;
    if (!info || info.kind !== 'worksheet') throw new TextError(msg('core.structure.moveNotWorksheet'));
    if (info.protected) throw new TextError(msg('core.structure.sheetProtected', { sheet: info.name }));
  }
  const scan = m.scan;
  if (op.type === 'insertRows' && scan && scan.maxRow + op.count >= MAX_ROWS) {
    throw new TextError(msg('core.structure.noRoomBelow'));
  }
  if (op.type === 'insertCols' && scan && scan.maxCol + op.count >= MAX_COLS) {
    throw new TextError(msg('core.structure.noRoomRight'));
  }
  if (op.type === 'move') {
    const h = op.src.r2 - op.src.r1;
    const w = op.src.c2 - op.src.c1;
    if (op.row + h >= MAX_ROWS || op.col + w >= MAX_COLS || op.row < 0 || op.col < 0) throw new TextError(msg('core.structure.doesNotFit'));
  }
  const blockedAt = (s: number, rect: Rect) => {
    for (const b of model.sheets[s].info.blocked) {
      if (b.kind === 'tableHeader') continue;
      if (b.r1 <= rect.r2 && b.r2 >= rect.r1 && b.c1 <= rect.c2 && b.c2 >= rect.c1) return b;
    }
    return undefined;
  };
  for (const a of scan?.arrayRanges ?? []) {
    if (splits(op, a)) throw new TextError(msg('core.structure.splitsArray', { range: rangeToA1(a) }));
  }
  for (const b of m.info.blocked) {
    if (b.kind === 'pivot' && splits(op, b)) throw new TextError(msg('core.structure.splitsPivot', { range: rangeToA1(b) }));
  }
  for (const t of m.info.tables) {
    const header = t.headerRows > 0 ? t.rect.r1 : -1;
    if (op.type === 'deleteRows' && header >= op.at && header < op.at + op.count && !(t.rect.r1 >= op.at && t.rect.r2 < op.at + op.count)) {
      throw new TextError(msg('core.structure.deletesTableHeader', { table: t.name }));
    }
    if ((op.type === 'deleteRows' && t.rect.r1 >= op.at && t.rect.r2 < op.at + op.count) || (op.type === 'deleteCols' && t.rect.c1 >= op.at && t.rect.c2 < op.at + op.count)) {
      throw new TextError(msg('core.structure.deletesTable', { table: t.name }));
    }
    if (op.type === 'move' && splits(op, t.rect)) throw new TextError(msg('core.structure.splitsTable', { table: t.name }));
  }
  if (op.type === 'move') {
    for (const merge of m.info.merges) {
      if (splits(op, merge)) throw new TextError(msg('core.structure.splitsMerge', { range: rangeToA1(merge) }));
    }
    if (op.toSheet !== op.sheet) {
      // Merges and tables belong to their sheet's own parts, so they can't follow cells to another sheet.
      const inside = (r: Rect) => r.r1 >= op.src.r1 && r.r2 <= op.src.r2 && r.c1 >= op.src.c1 && r.c2 <= op.src.c2;
      if (m.info.merges.some(inside)) throw new TextError(msg('core.structure.mergeToOtherSheet'));
      if (m.info.tables.some((t) => inside(t.rect))) throw new TextError(msg('core.structure.tableToOtherSheet'));
    }
    const dst = { r1: op.row, c1: op.col, r2: op.row + op.src.r2 - op.src.r1, c2: op.col + op.src.c2 - op.src.c1 };
    const b = blockedAt(op.toSheet, dst) ?? blockedAt(op.sheet, op.src);
    if (b) throw new TextError(b.reason);
  }
}

// ---------------------------------------------------------------------------
// Cells

interface CellOut {
  row: number;
  col: number;
  xml: string;
}

interface Collected {
  p: string;
  cells: (CellOut & { idx: number })[];
  /** Row index → the row element's attributes, without r and spans. */
  rows: Map<number, string>;
}

function collect(m: SheetModel): Collected {
  const scan = m.scan!;
  const xml = m.xml!;
  const g = scan.grid;
  const cells: Collected['cells'] = [];
  for (let row = 0; row < g.rowCount; row++) {
    for (let i = g.rowPtr[row]; i < g.rowPtr[row + 1]; i++) {
      cells.push({ row, col: g.cols[i], idx: i, xml: xml.slice(scan.cellStart[i], scan.cellEnd[i]) });
    }
  }
  const rows = new Map<number, string>();
  for (let i = 0; i < scan.rowRec.row.length; i++) {
    const tag = xml.slice(scan.rowRec.start[i], scan.rowRec.tagEnd[i] + 1);
    const attrs = tag
      .replace(/^<[\w:]+/, '')
      .replace(/\/?>$/, '')
      .replace(/\s(r|spans)="[^"]*"/g, '');
    if (attrs.trim()) rows.set(scan.rowRec.row[i], attrs);
  }
  return { p: scan.prefix, cells, rows };
}

/** A cell's XML at a new position, with its formula written out in full for that position. */
function relocate(xml: string, p: string, row: number, col: number, formula?: { text: string; arrayRef?: string }): string {
  const end = findTagEnd(xml, 0);
  const selfClosing = xml.charCodeAt(end - 1) === 47;
  let tag = xml.slice(0, selfClosing ? end - 1 : end);
  tag = /\sr="[^"]*"/.test(tag) ? tag.replace(/\sr="[^"]*"/, ` r="${toA1(row, col)}"`) : `${tag} r="${toA1(row, col)}"`;
  let inner = selfClosing ? '' : xml.slice(end + 1, xml.lastIndexOf('<'));
  if (formula) {
    inner = replaceElements(inner, 'f', () => '', 1);
    const attrs = formula.arrayRef ? ` t="array" ref="${formula.arrayRef}"` : '';
    inner = `<${p}f${attrs}>${escapeXml(formula.text)}</${p}f>${inner}`;
  }
  return inner ? `${tag}>${inner}</${p}c>` : `${tag}/>`;
}

function place(
  c: CellOut & { idx: number },
  from: number,
  to: number,
  map: RefMap,
  fx: (f: string, from: number, to?: number) => string,
  m: SheetModel,
): CellOut | null {
  const at = map.range(from, { r1: c.row, c1: c.col, r2: c.row, c2: c.col });
  if (!at || at.sheet !== to) return null;
  const { r1: row, c1: col } = at.rect;
  const scan = m.scan!;
  const f = scan.grid.formulas.get(c.idx);
  const meta = scan.fmeta.get(c.idx);
  if (f === undefined || meta?.kind === 'dataTable') return row === c.row && col === c.col ? c : { row, col, xml: relocate(c.xml, scan.prefix, row, col) };
  let arrayRef: string | undefined;
  if (meta?.kind === 'array' && meta.ref) {
    const r = map.range(from, meta.ref);
    arrayRef = r ? rangeToA1(r.rect) : undefined;
  }
  return { row, col, xml: relocate(c.xml, scan.prefix, row, col, { text: fx(f, from, to), arrayRef }) };
}

/**
 * The sheet's <sheetData> after the change. Sheets whose cells move are
 * rebuilt cell by cell; on other sheets only formulas that point at moved
 * cells are rewritten (with every member of a shared formula written out, as
 * its members no longer follow one pattern).
 */
function rewriteCells(
  m: SheetModel,
  s: number,
  op: StructuralOp,
  map: RefMap,
  fx: (f: string, from: number, to?: number) => string,
  mentions: (f: string) => boolean,
  extra: CellOut[],
  incoming: CellOut[],
): string {
  const scan = m.scan!;
  const xml = m.xml!;
  const moves = map.sheets.has(s);
  const g = scan.grid;
  if (!moves) {
    // Only formulas can change here.
    const rewritten = new Map<number, string>();
    for (const [idx, f] of g.formulas) {
      if (!mentions(f)) continue;
      const next = fx(f, s);
      if (next !== f) rewritten.set(idx, next);
    }
    const changed = new Set(rewritten.keys());
    if (!changed.size) return xml;
    for (const group of scan.sharedGroups.values()) {
      if (group.members.some((i) => changed.has(i)) || changed.has(group.master)) {
        changed.add(group.master);
        for (const i of group.members) changed.add(i);
      }
    }
    const parts: string[] = [];
    let pos = 0;
    const order = [...changed].sort((a, b) => scan.cellStart[a] - scan.cellStart[b]);
    for (const idx of order) {
      const row = rowOf(g.rowPtr, idx);
      const meta = scan.fmeta.get(idx);
      if (meta?.kind === 'dataTable') continue;
      parts.push(xml.slice(pos, scan.cellStart[idx]));
      const cell = xml.slice(scan.cellStart[idx], scan.cellEnd[idx]);
      const arrayRef = meta?.kind === 'array' && meta.ref ? rangeToA1(meta.ref) : undefined;
      parts.push(relocate(cell, scan.prefix, row, g.cols[idx], { text: rewritten.get(idx) ?? g.formulas.get(idx)!, arrayRef }));
      pos = scan.cellEnd[idx];
    }
    parts.push(xml.slice(pos));
    return parts.join('');
  }

  const { p, cells, rows } = collect(m);
  const out: CellOut[] = [];
  const dst = op.type === 'move' ? { r1: op.row, c1: op.col, r2: op.row + op.src.r2 - op.src.r1, c2: op.col + op.src.c2 - op.src.c1 } : undefined;
  for (const c of cells) {
    if (op.type === 'move') {
      const fromSrc = s === op.sheet && inRect(op.src, c.row, c.col);
      // Cells moving to another sheet were collected there; the ones pasted over are replaced.
      if (fromSrc && op.toSheet !== s) continue;
      if (!fromSrc && s === op.toSheet && dst && inRect(dst, c.row, c.col)) continue;
    }
    const placed = place(c, s, s, map, fx, m);
    if (placed) out.push(placed);
  }
  out.push(...incoming);
  // New cells (such as names for inserted table columns) replace whatever is there.
  const taken = new Set(extra.map((c) => `${c.row}:${c.col}`));
  const cellsOut = [...out.filter((c) => !taken.has(`${c.row}:${c.col}`)), ...extra];

  // Row formats move with their rows; other changes leave them where they are.
  const rowsOut = new Map<number, string>();
  for (const [row, attrs] of rows) {
    if (op.type === 'insertRows' || op.type === 'deleteRows') {
      const at = map.range(s, { r1: row, r2: row, c1: 0, c2: 0 });
      if (at) rowsOut.set(at.rect.r1, attrs);
    } else {
      rowsOut.set(row, attrs);
    }
  }
  // Rows inserted inside a group join it, as Excel's do.
  if (op.type === 'insertRows') {
    const join = Math.min(outlineLevelOf(rows.get(op.at - 1)), outlineLevelOf(rows.get(op.at)));
    for (let row = op.at; join && row < op.at + op.count; row++) rowsOut.set(row, ` outlineLevel="${join}"`);
  }

  const byRow = new Map<number, CellOut[]>();
  for (const c of cellsOut) {
    const list = byRow.get(c.row);
    if (list) list.push(c);
    else byRow.set(c.row, [c]);
  }
  const allRows = [...new Set([...rowsOut.keys(), ...byRow.keys()])].sort((a, b) => a - b);
  const body = allRows
    .map((row) => {
      const attrs = rowsOut.get(row) ?? '';
      const list = (byRow.get(row) ?? []).sort((a, b) => a.col - b.col);
      return list.length ? `<${p}row r="${row + 1}"${attrs}>${list.map((c) => c.xml).join('')}</${p}row>` : `<${p}row r="${row + 1}"${attrs}/>`;
    })
    .join('');
  const sheetData = `<${p}sheetData>${body}</${p}sheetData>`;
  let next = scan.sheetDataStart >= 0 ? xml.slice(0, scan.sheetDataStart) + sheetData + xml.slice(scan.sheetDataEnd) : insertSheetData(xml, sheetData);

  // The used range.
  let r1 = Infinity;
  let c1 = Infinity;
  let r2 = -1;
  let c2 = -1;
  for (const c of cellsOut) {
    r1 = Math.min(r1, c.row);
    r2 = Math.max(r2, c.row);
    c1 = Math.min(c1, c.col);
    c2 = Math.max(c2, c.col);
  }
  const dim = r2 < 0 ? 'A1' : rangeToA1({ r1, c1, r2, c2 });
  next = next.replace(/(<(?:\w+:)?dimension\b[^>]*\bref=")[^"]*"/, `$1${dim}"`);
  return next;
}

/** Adds <sheetData> to a sheet that has none, after the elements that come before it in a worksheet. */
function insertSheetData(xml: string, sheetData: string): string {
  const before = /^[\s\S]*(?:<\/(?:\w+:)?(?:sheetPr|sheetViews|sheetFormatPr|cols)>|<(?:\w+:)?(?:sheetPr|dimension|sheetFormatPr)\b[^>]*\/>)/.exec(xml);
  const at = before ? before[0].length : findTagEnd(xml, xml.search(/<(?:\w+:)?worksheet\b/)) + 1;
  return xml.slice(0, at) + sheetData + xml.slice(at);
}

function rowOf(rowPtr: Int32Array, idx: number): number {
  let lo = 0;
  let hi = rowPtr.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (rowPtr[mid] <= idx) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

// ---------------------------------------------------------------------------
// Everything in a worksheet outside its cells

/** Where <sheetData> starts and ends; both at the end of the document if there is none. */
function sheetDataSpan(xml: string): [number, number] {
  const open = /<((?:\w+:)?sheetData)\b/.exec(xml);
  if (!open) return [xml.length, xml.length];
  const end = findTagEnd(xml, open.index);
  if (xml.charCodeAt(end - 1) === 47) return [open.index, end + 1];
  const close = xml.indexOf(`</${open[1]}>`, end);
  return [open.index, close < 0 ? xml.length : close + open[1].length + 3];
}

/** Calls fn on each element named `tag` (any prefix); fn returns its replacement, '' to remove it. */
function eachElement(xml: string, tag: string, fn: (el: string) => string): string {
  return replaceElements(xml, tag, (_el, text) => fn(text));
}

/** Rewrites the text of formula elements (<formula>, <formula1>, <xm:f>…), which hold escaped XML text. */
function rewriteFormulaElements(xml: string, tags: string[], fn: (f: string) => string): string {
  const re = new RegExp(`(<(?:\\w+:)?(?:${tags.join('|')})(?:\\s[^>]*)?>)([^<]*)(</(?:\\w+:)?(?:${tags.join('|')})>)`, 'g');
  return xml.replace(re, (all, open: string, text: string, close: string) => {
    const f = decodeEntities(text);
    const next = fn(f);
    return next === f ? all : `${open}${escapeXml(next)}${close}`;
  });
}

function rewriteRefAttrs(xml: string, tags: string[], sheet: number, map: RefMap): string {
  const re = new RegExp(`(<(?:\\w+:)?(?:${tags.join('|')})\\b[^>]*?\\sref=")([^"]*)(")`, 'g');
  return xml.replace(re, (all, a: string, ref: string, b: string) => {
    const next = transformSqref(ref, sheet, map);
    return next ? `${a}${next}${b}` : all;
  });
}

function rewriteSheetMeta(xml: string, s: number, op: StructuralOp, map: RefMap, fx: (f: string, from: number) => string): string {
  const cut = sheetDataSpan(xml);
  let head = xml.slice(0, cut[0]);
  let tail = xml.slice(cut[1]);
  const moves = map.sheets.has(s);

  // Formulas in validation, conditional formatting and sparklines, which may point anywhere.
  tail = rewriteFormulaElements(tail, ['formula', 'formula1', 'formula2', 'f'], (f) => fx(f, s));
  if (!moves) return head + xml.slice(cut[0], cut[1]) + tail;

  const sq = (v: string) => transformSqref(v, s, map);
  // Ranges on this sheet. Rules and links whose ranges are all gone go too.
  for (const tag of ['dataValidation', 'conditionalFormatting', 'protectedRange', 'ignoredError']) {
    tail = eachElement(tail, tag, (el) => {
      const own = /^<[^>]*?\ssqref="([^"]*)"/.exec(el);
      const x14 = /<(?:\w+:)?sqref>([^<]*)<\/(?:\w+:)?sqref>/.exec(el);
      if (own) {
        const next = sq(own[1]);
        return next ? el.replace(/\ssqref="[^"]*"/, ` sqref="${next}"`) : '';
      }
      if (x14) {
        const next = sq(x14[1]);
        return next ? el.replace(x14[0], x14[0].replace(x14[1], next)) : '';
      }
      return el;
    });
  }
  // Sparklines' cells. Only these: the <sqref> of the Excel 2010 elements above has already moved.
  tail = eachElement(tail, 'sparkline', (el) =>
    el.replace(/(<(?:\w+:)?sqref>)([^<]*)(<\/(?:\w+:)?sqref>)/, (all, a: string, v: string, b: string) => {
      const next = sq(v);
      return next ? a + next + b : all;
    }),
  );
  for (const tag of ['mergeCell', 'hyperlink']) {
    tail = eachElement(tail, tag, (el) => {
      const ref = /\sref="([^"]*)"/.exec(el);
      if (!ref) return el;
      const next = sq(ref[1]);
      return next ? el.replace(ref[0], ` ref="${next}"`) : '';
    });
  }
  // Containers that lost every item go, and the rest keep their counts right.
  for (const [box, item] of [['mergeCells', 'mergeCell'], ['dataValidations', 'dataValidation'], ['hyperlinks', 'hyperlink'], ['protectedRanges', 'protectedRange'], ['ignoredErrors', 'ignoredError']]) {
    tail = replaceElements(tail, box, ({ qname, attrs, inner }) => {
      const n = inner ? countElements(inner, item) : 0;
      if (!n) return '';
      return `<${qname}${attrs.replace(/\scount="\d+"/, ` count="${n}"`)}>${inner}</${qname}>`;
    });
  }
  tail = rewriteRefAttrs(tail, ['autoFilter', 'sortState', 'sortCondition'], s, map);

  // The selected cell, which may have gone.
  head = head.replace(/<((?:\w+:)?selection)\b([^>]*?)\/?>/g, (all, tag: string, attrs: string) => {
    const active = /\sactiveCell="([^"]*)"/.exec(attrs);
    const next = active ? sq(active[1]) : '';
    if (!active) return all;
    if (!next) return `<${tag}${attrs.replace(/\sactiveCell="[^"]*"/, ' activeCell="A1"').replace(/\ssqref="[^"]*"/, ' sqref="A1"')}/>`;
    const sqref = /\ssqref="([^"]*)"/.exec(attrs);
    return `<${tag}${attrs.replace(active[0], ` activeCell="${next}"`).replace(/\ssqref="[^"]*"/, ` sqref="${(sqref && sq(sqref[1])) || next}"`)}/>`;
  });

  if (op.type === 'insertCols' || op.type === 'deleteCols') head = shiftCols(head, op);
  if (op.type === 'deleteRows') head = lowerOutlineLevel(head, 'outlineLevelRow', 'row', xml.slice(cut[0], cut[1]));
  if (op.type === 'deleteCols') head = lowerOutlineLevel(head, 'outlineLevelCol', 'col', head);
  if (op.type === 'insertRows' || op.type === 'deleteRows') tail = shiftBreaks(tail, 'rowBreaks', op);
  if (op.type === 'insertCols' || op.type === 'deleteCols') tail = shiftBreaks(tail, 'colBreaks', op);
  return head + xml.slice(cut[0], cut[1]) + tail;
}

/** An element's outline level, from its attributes. */
const outlineLevelOf = (attrs: string | undefined) => Number(/\soutlineLevel="(\d+)"/.exec(attrs ?? '')?.[1] ?? 0);

/** <sheetFormatPr>'s deepest outline level, lowered to that of the rows (or columns) left once the deepest are deleted. */
function lowerOutlineLevel(head: string, name: string, tag: string, lines: string): string {
  let deepest = 0;
  for (const m of lines.matchAll(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*?\\soutlineLevel="(\\d+)"`, 'g'))) deepest = Math.max(deepest, Number(m[1]));
  return head.replace(new RegExp(`(<(?:\\w+:)?sheetFormatPr\\b[^>]*?\\s${name}=")(\\d+)"`), (all, open: string, v: string) => (Number(v) > deepest ? `${open}${deepest}"` : all));
}

/** Column widths, formats and outline levels (<cols>) follow inserted and deleted columns. */
function shiftCols(head: string, op: LineOp): string {
  const COL = /<((?:\w+:)?col)\b([^>]*?)\/>/g;
  const span = (attrs: string) => [Number(/\smin="(\d+)"/.exec(attrs)?.[1]) - 1, Number(/\smax="(\d+)"/.exec(attrs)?.[1]) - 1];
  // Columns inserted inside a group join it, as Excel's do: a <col> spanning where they go
  // takes them in; between two that end and start there, they need one of their own, like
  // the column left of them, as Excel's take its width and format.
  let join = 0;
  if (op.type === 'insertCols') {
    const cols = [...head.matchAll(COL)].map((m) => ({ span: span(m[2]), level: outlineLevelOf(m[2]) }));
    const levelAt = (c: number) => cols.find((x) => x.span[0] <= c && c <= x.span[1])?.level ?? 0;
    if (!cols.some((x) => x.span[0] < op.at && op.at <= x.span[1])) join = Math.min(levelAt(op.at - 1), levelAt(op.at));
  }
  return head.replace(COL, (all, tag: string, attrs: string) => {
    const [min, max] = span(attrs);
    if (Number.isNaN(min) || Number.isNaN(max)) return all;
    if (join && max === op.at - 1) {
      const like = attrs.replace(/\s(min|max|outlineLevel|collapsed)="[^"]*"/g, '');
      return `${all}<${tag} min="${op.at + 1}" max="${op.at + op.count}"${like} outlineLevel="${join}"/>`;
    }
    let a = min;
    let b = max;
    if (op.type === 'insertCols') {
      if (a >= op.at) {
        a += op.count;
        b += op.count;
      } else if (b >= op.at) {
        b += op.count;
      }
    } else {
      const end = op.at + op.count;
      a = a < op.at ? a : a >= end ? a - op.count : op.at;
      b = b < op.at ? b : b >= end ? b - op.count : op.at - 1;
      if (b < a) return '';
    }
    a = Math.min(a, MAX_COLS - 1);
    b = Math.min(b, MAX_COLS - 1);
    return `<${tag}${attrs.replace(/\smin="\d+"/, ` min="${a + 1}"`).replace(/\smax="\d+"/, ` max="${b + 1}"`)}/>`;
  });
}

/** Page breaks (<brk id> is the 1-based row or column the break follows). */
function shiftBreaks(tail: string, tag: string, op: LineOp): string {
  return replaceElements(tail, tag, ({ qname, attrs, inner }, all) => {
    if (inner === null) return all;
    const open = `<${qname}${attrs}>`;
    const body = inner;
    const close = `</${qname}>`;
    const insert = op.type.startsWith('insert');
    const next = body.replace(/<((?:\w+:)?brk)\b([^>]*?)\/>/g, (el, t: string, attrs: string) => {
      const id = Number(/\sid="(\d+)"/.exec(attrs)?.[1]);
      if (Number.isNaN(id)) return el;
      if (insert) return id >= op.at ? `<${t}${attrs.replace(/\sid="\d+"/, ` id="${id + op.count}"`)}/>` : el;
      if (id >= op.at && id < op.at + op.count) return '';
      return id >= op.at + op.count ? `<${t}${attrs.replace(/\sid="\d+"/, ` id="${id - op.count}"`)}/>` : el;
    });
    const n = countElements(next, 'brk');
    return `${open.replace(/\scount="\d+"/, ` count="${n}"`).replace(/\smanualBreakCount="\d+"/, ` manualBreakCount="${n}"`)}${next}${close}`;
  }, 1);
}

// ---------------------------------------------------------------------------
// Other parts

/** Tables follow their cells; columns inserted or deleted inside a table are added to or removed from it. */
function tableChanges(
  pkg: OoxmlPackage,
  model: WorkbookModel,
  op: StructuralOp,
  map: RefMap,
  fx: (f: string, from: number) => string,
): Map<number, CellOut[]> {
  const extra = new Map<number, CellOut[]>();
  model.sheets.forEach((m, s) => {
    if (!map.sheets.has(s) || !m.path) return;
    for (const rel of pkg.rels(m.path)) {
      if (rel.type !== 'table' || rel.external) continue;
      const xml = pkg.text(rel.target);
      if (!xml) continue;
      const refMatch = /<(?:\w+:)?table\b[^>]*?\sref="([^"]*)"/.exec(xml);
      const rect = refMatch ? parseRange(refMatch[1]) : null;
      if (!rect) continue;
      const moved = map.range(s, rect);
      if (!moved || moved.sheet !== s) continue;
      let next = xml.replace(/(<(?:\w+:)?table\b[^>]*?\sref=")[^"]*"/, `$1${rangeToA1(moved.rect)}"`);
      next = rewriteRefAttrs(next, ['autoFilter', 'sortState', 'sortCondition'], s, map);
      next = rewriteFormulaElements(next, ['calculatedColumnFormula', 'totalsRowFormula'], (f) => fx(f, s));
      if (op.type === 'insertCols' && rect.c1 < op.at && op.at <= rect.c2) {
        // New columns need names, both in the table and in its header row.
        const ids = [...next.matchAll(/<(?:\w+:)?tableColumn\b[^>]*?\sid="(\d+)"/g)].map((x) => Number(x[1]));
        const used = new Set([...next.matchAll(/<(?:\w+:)?tableColumn\b[^>]*?\sname="([^"]*)"/g)].map((x) => decodeEntities(x[1]).toLowerCase()));
        let id = Math.max(0, ...ids);
        const added: string[] = [];
        const p = /<(\w+:)?tableColumns\b/.exec(next)?.[1] ?? '';
        const header = /\sheaderRowCount="0"/.test(next) ? -1 : rect.r1;
        for (let k = 0, n = 1; k < op.count; k++) {
          while (used.has(`column${n}`)) n++;
          const name = `Column${n}`;
          used.add(name.toLowerCase());
          added.push(`<${p}tableColumn id="${++id}" name="${name}"/>`);
          if (header >= 0) {
            const col = op.at + k;
            const list = extra.get(s) ?? [];
            list.push({ row: header, col, xml: `<${m.scan?.prefix ?? ''}c r="${toA1(header, col)}" t="inlineStr"><${m.scan?.prefix ?? ''}is><${m.scan?.prefix ?? ''}t>${name}</${m.scan?.prefix ?? ''}t></${m.scan?.prefix ?? ''}is></${m.scan?.prefix ?? ''}c>` });
            extra.set(s, list);
          }
        }
        next = spliceColumns(next, op.at - rect.c1, 0, added);
      }
      if (op.type === 'deleteCols') {
        const from = Math.max(op.at, rect.c1);
        const to = Math.min(op.at + op.count - 1, rect.c2);
        if (from <= to) next = spliceColumns(next, from - rect.c1, to - from + 1, []);
      }
      if (next !== xml) pkg.setText(rel.target, next);
    }
  });
  return extra;
}

/** Removes `remove` table columns at position `at` and inserts `add` there, keeping the count right. */
function spliceColumns(xml: string, at: number, remove: number, add: string[]): string {
  return replaceElements(
    xml,
    'tableColumns',
    ({ qname, attrs, inner }, all) => {
      if (inner === null) return all;
      const cols = elements(inner, 'tableColumn').map((c) => inner.slice(c.start, c.end));
      cols.splice(at, remove, ...add);
      return `<${qname}${attrs.replace(/\scount="\d+"/, ` count="${cols.length}"`)}>${cols.join('')}</${qname}>`;
    },
    1,
  );
}

function pivotSource(xml: string, map: RefMap, ctx: FormulaContext): string {
  return xml.replace(/<((?:\w+:)?worksheetSource)\b([^>]*?)\/>/, (all, tag: string, attrs: string) => {
    const sheetName = /\ssheet="([^"]*)"/.exec(attrs)?.[1];
    const ref = /\sref="([^"]*)"/.exec(attrs)?.[1];
    const sheet = sheetName !== undefined ? ctx.indexOf(decodeEntities(sheetName)) : undefined;
    if (sheet === undefined || !ref || !map.sheets.has(sheet)) return all;
    const next = transformSqref(ref, sheet, map);
    return next ? `<${tag}${attrs.replace(/\sref="[^"]*"/, ` ref="${next}"`)}/>` : all;
  });
}

function comments(xml: string, s: number, map: RefMap): string {
  return eachElement(xml, 'comment', (el) => {
    const ref = /\sref="([^"]*)"/.exec(el);
    if (!ref) return el;
    const next = transformSqref(ref[1], s, map);
    return next ? el.replace(ref[0], ` ref="${next}"`) : '';
  });
}

function threads(xml: string, s: number, map: RefMap): string {
  return eachElement(xml, 'threadedComment', (el) => {
    const ref = /\sref="([^"]*)"/.exec(el);
    if (!ref) return el;
    const next = transformSqref(ref[1], s, map);
    return next ? el.replace(ref[0], ` ref="${next}"`) : '';
  });
}

/** The legacy drawing that positions each note's box. */
function vmlNotes(xml: string, s: number, map: RefMap): string {
  return replaceElements(xml, 'shape', ({ qname }, shape) => {
    if (qname !== 'v:shape') return shape;
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(shape);
    const col = /<x:Column>(\d+)<\/x:Column>/.exec(shape);
    if (!row || !col) return shape;
    const r = Number(row[1]);
    const c = Number(col[1]);
    const at = map.range(s, { r1: r, r2: r, c1: c, c2: c });
    if (!at || at.sheet !== s) return '';
    const dr = at.rect.r1 - r;
    const dc = at.rect.c1 - c;
    if (!dr && !dc) return shape;
    return shape
      .replace(row[0], `<x:Row>${at.rect.r1}</x:Row>`)
      .replace(col[0], `<x:Column>${at.rect.c1}</x:Column>`)
      .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (_a, v: string) => {
        const n = v.split(',').map((x) => Number(x.trim()));
        if (n.length !== 8) return _a;
        n[0] = Math.max(0, n[0] + dc);
        n[4] = Math.max(0, n[4] + dc);
        n[2] = Math.max(0, n[2] + dr);
        n[6] = Math.max(0, n[6] + dr);
        return `<x:Anchor>${n.join(', ')}</x:Anchor>`;
      });
  });
}

/** Pictures, charts and shapes are anchored to cells; they move, or stretch, with them. */
function anchors(xml: string, op: LineOp): string {
  const rows = op.type === 'insertRows' || op.type === 'deleteRows';
  const tag = rows ? 'row' : 'col';
  const insert = op.type.startsWith('insert');
  const shift = (i: number) => {
    if (insert) return i >= op.at ? i + op.count : i;
    if (i < op.at) return i;
    return i >= op.at + op.count ? i - op.count : op.at;
  };
  const move = (_el: unknown, anchor: string) =>
    anchor.replace(new RegExp(`(<(?:\\w+:)?${tag}>)(\\d+)(</(?:\\w+:)?${tag}>)`), (_a, a: string, v: string, b: string) => `${a}${shift(Number(v))}${b}`);
  return replaceElements(replaceElements(xml, 'from', move), 'to', move);
}
