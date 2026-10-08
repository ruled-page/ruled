import type { EngineValue } from './engineValue';
import { parseA1Ref } from '@borgar/fx';
import { writeDefinedNames } from './definedNames';
import { isDynamicArrayFormula } from './formula';
import { writeHyperlinks, type SheetLinks } from './hyperlinks';
import { rowAttributes, writeLayout, writeTableFilter, type SheetLayout } from './layout';
import { cellKey, MAX_COLS, MAX_ROWS, rangeToA1, toA1 } from './refs';
import type { CellState } from './session';
import { findCell, type ScanResult } from './sheetScanner';
import { CellFlag, VT, type CellNote, type DefinedName, type FilterState, type Rect, type SheetGrid } from './types';
import type { Person } from './comments';
import { writeComments } from './commentWriter';
import type { SheetModel, WorkbookModel } from './workbook';
import { countElements, encodeOoxmlText, escapeXml, findTagEnd, firstElement, parseAttrs, replaceElements, setAttributes } from './xml';
import { setCustomProperties } from './properties';
import { relsPathFor, type OoxmlPackage } from './zip';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export interface SaveInput {
  pkg: OoxmlPackage;
  model: WorkbookModel;
  /** Per sheet: cells whose content or style differs from the file. */
  edits: Map<number, CellState>[];
  /** Engine value to store as the saved result of a formula cell, if trusted. */
  cachedValue: (sheet: number, row: number, col: number) => EngineValue | undefined;
  /** Per sheet: untouched formula cells whose result changed. */
  recalculated: Set<number>[];
  /** Remove the VBA project, making the file a plain .xlsx. */
  stripMacros?: boolean;
  /** Custom document properties to set, e.g. the schema the file was checked against. */
  properties?: Record<string, string>;
  /** Whether the engine's result for a formula cell is an array, which Excel would spill. */
  isArray?: (sheet: number, row: number, col: number) => boolean;
  /** Per sheet: where each dynamic array formula's results spill now, by its cell's key; a blocked one covers only its own cell. */
  spills?: Map<number, Rect>[];
  /** The value a formula's results spill into a cell (null where none does now), and the cell's format. */
  spilled?: (sheet: number, row: number, col: number) => { value: EngineValue; style: number };
  /** Whether anything in a sheet's range (the whole sheet when null) differs from the file as opened. */
  changedIn?: (sheet: number, rect: Rect | null) => boolean;
  /** Per sheet: its layout, where it has been edited. */
  layout?: (SheetLayout | undefined)[];
  /** Per sheet: cells whose note or thread differs from the file, with what they hold now (null: neither). */
  notes?: Map<number, CellNote | null>[];
  /** Everyone who has written a threaded comment, by person id, for those new to the file. */
  people?: ReadonlyMap<string, Person>;
  /** The defined names as the file had them and as they are now, when they differ. */
  names?: { file: DefinedName[]; now: DefinedName[] };
  /** Per sheet: its hyperlinks, where they have been edited. */
  links?: (SheetLinks | undefined)[];
}

type Patch =
  /** `ref`: where a dynamic array formula's results spill. */
  | { kind: 'edit'; state: CellState; value?: EngineValue; dynamic?: boolean; ref?: Rect }
  | { kind: 'value'; value: EngineValue; ref?: Rect }
  | { kind: 'expand'; formula: string; value: EngineValue | undefined }
  /** A value spilled from a formula in another cell, written without a formula as Excel writes it; null where a spill no longer reaches. */
  | { kind: 'spilled'; value: EngineValue; style: number };

/** Writes the edited workbook, changing only the parts that must change. */
export function writeWorkbook(input: SaveInput): Uint8Array {
  return buildPackage(input).toBytes();
}

/** The edited workbook as a package, before it is zipped. */
export function buildPackage(input: SaveInput): OoxmlPackage {
  const { pkg, model } = input;

  // The metadata entry that marks dynamic array formulas, added when the first one is written.
  let dynamicArray: number | undefined;
  const cm = () => (dynamicArray ??= dynamicArrayMetadata(pkg, model.workbookPath));
  model.sheets.forEach((m, s) => {
    if (!m.scan || m.xml === undefined) return;
    const patches = sheetPatches(m, s, input);
    const layout = input.layout?.[s];
    const rows = layout ? rowAttributes(layout) : new Map<number, Record<string, string | null>>();
    let xml = patches.size || rows.size ? rewriteSheet(m.xml, m.scan, patches, cm, rows) : m.xml;
    if (layout?.changed) xml = writeLayout(xml, m.scan.prefix, layout);
    const links = input.links?.[s];
    if (links?.changed) xml = writeHyperlinks(pkg, m.path, xml, m.scan.prefix, links.list);
    if (xml !== m.xml) pkg.setText(m.path, xml);
    // A table's filter is kept in the table's own part: the new one set, a table's that went cleared.
    if (layout?.filter !== undefined) {
      const now = layout.filter;
      const was = layout.fileFilter.table;
      const patch = (name: string, filter: FilterState | null) => {
        const t = m.info.tables.find((x) => x.name === name);
        const txml = t?.path && pkg.text(t.path);
        if (t && txml) pkg.setText(t.path!, writeTableFilter(txml, filter, { ...t.rect, r2: t.rect.r2 - t.totalsRows }));
      };
      if (was && was !== now?.table) patch(was, null);
      if (now?.table) patch(now.table, now);
    }
  });
  if (input.notes?.some((m) => m.size)) writeComments(pkg, model, input.notes, input.people ?? new Map());

  // Excel checks calcChain.xml against the cells that hold formulas and
  // reports the file as damaged if they disagree. It rebuilds the chain
  // itself when the part is missing.
  removeCalcChain(pkg, model.workbookPath);

  if (input.names) {
    const xml = pkg.text(model.workbookPath);
    if (xml) pkg.setText(model.workbookPath, writeDefinedNames(xml, input.names.file, input.names.now));
  }

  const wbXml = pkg.text(model.workbookPath);
  if (wbXml) pkg.setText(model.workbookPath, forceFullCalcOnLoad(wbXml));

  // A pivot table keeps its own copy of its source data; Excel refreshes it on
  // opening only when asked to, so ask when that data changed.
  for (const path of pkg.list()) {
    if (!/pivotCacheDefinition\d*\.xml$/.test(path)) continue;
    const xml = pkg.text(path);
    if (xml && pivotSourceChanged(xml, model, input)) pkg.setText(path, setRootAttr(xml, 'refreshOnLoad', '1'));
  }

  if (input.stripMacros) removeMacros(pkg, model.workbookPath);
  if (input.properties) setCustomProperties(pkg, input.properties);

  if (model.styles.additions.length || Object.values(model.styles.addedParts).some((a) => a.length)) {
    const stylesRel = pkg.rels(model.workbookPath).find((r) => r.type === 'styles');
    const xml = stylesRel ? pkg.text(stylesRel.target) : undefined;
    if (stylesRel && xml) pkg.setText(stylesRel.target, appendStyles(xml, model));
  }

  return pkg;
}

function sheetPatches(m: SheetModel, s: number, input: SaveInput): Map<number, Patch> {
  const patches = new Map<number, Patch>();
  const grid = m.scan!.grid;
  for (const [key, state] of input.edits[s]) {
    const row = Math.floor(key / MAX_COLS);
    const col = key % MAX_COLS;
    const formula = state.content.kind === 'formula' ? state.content.formula : undefined;
    const value = formula !== undefined ? input.cachedValue(s, row, col) : undefined;
    const dynamic = formula !== undefined && (isDynamicArrayFormula(formula) || !!input.isArray?.(s, row, col));
    patches.set(key, { kind: 'edit', state, value, dynamic });
  }
  for (const key of input.recalculated[s]) {
    if (patches.has(key)) continue;
    const v = input.cachedValue(s, Math.floor(key / MAX_COLS), key % MAX_COLS);
    if (v !== undefined) patches.set(key, { kind: 'value', value: v });
  }
  // Shared formulas: followers carry no formula text of their own, only a
  // pointer to the master cell. If the master was edited, every follower
  // gets its formula written out in full.
  for (const group of m.scan!.sharedGroups.values()) {
    if (group.master < 0) continue;
    const masterKey = keyOf(grid, group.master);
    if (!input.edits[s].has(masterKey)) continue;
    for (const idx of group.members) {
      const key = keyOf(grid, idx);
      const existing = patches.get(key);
      if (existing?.kind === 'edit') continue;
      const formula = grid.formulas.get(idx);
      if (formula === undefined) continue;
      const value =
        existing?.kind === 'value' ? existing.value : grid.flags[idx] & CellFlag.NoCache ? undefined : gridValueOf(grid, idx);
      patches.set(key, { kind: 'expand', formula, value });
    }
  }
  if (input.spills && input.spilled) spillPatches(m, s, input, patches);
  return patches;
}

/**
 * Dynamic arrays as Excel writes them: the formula's cell with its spill
 * range as the formula's ref, and the values it spills into the cells of
 * that range, without formulas. Cells holding saved values of a spill that
 * no longer reaches them are cleared. Values that are as saved stay as they
 * were written.
 */
function spillPatches(m: SheetModel, s: number, input: SaveInput, patches: Map<number, Patch>): void {
  const scan = m.scan!;
  const grid = scan.grid;
  const covered = new Set<number>();
  for (const [key, rect] of input.spills![s]) {
    const row = Math.floor(key / MAX_COLS);
    const col = key % MAX_COLS;
    const idx = findCell(grid, row, col);
    const own = patches.get(key);
    if (own?.kind === 'edit' || own?.kind === 'value') own.ref = rect;
    else if (!own && !sameRect(scan.fmeta.get(idx)?.ref, rect)) {
      const value = input.cachedValue(s, row, col) ?? (idx >= 0 && !(grid.flags[idx] & CellFlag.NoCache) ? gridValueOf(grid, idx) : null);
      patches.set(key, { kind: 'value', value, ref: rect });
    }
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        const k = cellKey(r, c);
        if (k === key) continue;
        covered.add(k);
        const spilled = input.spilled!(s, r, c);
        const i = findCell(grid, r, c);
        if (!input.edits[s].has(k) && i >= 0 && grid.flags[i] & CellFlag.Spilled && sameValue(gridValueOf(grid, i), spilled.value)) continue;
        patches.set(k, { kind: 'spilled', ...spilled });
      }
    }
  }
  for (let row = 0; row < grid.rowCount; row++) {
    for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) {
      const k = cellKey(row, grid.cols[i]);
      if (grid.flags[i] & CellFlag.Spilled && !covered.has(k) && !patches.has(k)) patches.set(k, { kind: 'spilled', value: null, style: grid.styles[i] });
    }
  }
}

function sameRect(a: Rect | undefined, b: Rect): boolean {
  return !!a && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

function sameValue(a: EngineValue, b: EngineValue): boolean {
  if (a && b && typeof a === 'object' && typeof b === 'object') return a.error === b.error;
  return a === b;
}

function keyOf(grid: SheetGrid, idx: number): number {
  let lo = 0;
  let hi = grid.rowCount - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (grid.rowPtr[mid] <= idx) lo = mid;
    else hi = mid - 1;
  }
  return cellKey(lo, grid.cols[idx]);
}

function gridValueOf(grid: SheetGrid, idx: number): EngineValue {
  switch (grid.types[idx]) {
    case VT.Number:
      return grid.nums[idx];
    case VT.String:
      return grid.strs[idx] ?? '';
    case VT.Bool:
      return grid.nums[idx] === 1;
    case VT.Error:
      return { error: grid.strs[idx] ?? '#VALUE!' };
    default:
      return null;
  }
}

function rewriteSheet(
  xml: string,
  scan: ScanResult,
  patches: Map<number, Patch>,
  cm: () => number,
  rowAttrs: Map<number, Record<string, string | null>>,
): string {
  const p = scan.prefix;
  const grid = scan.grid;
  const rec = scan.rowRec;

  const byRow = new Map<number, number[]>();
  for (const key of patches.keys()) {
    const row = Math.floor(key / MAX_COLS);
    let cols = byRow.get(row);
    if (!cols) byRow.set(row, (cols = []));
    cols.push(key % MAX_COLS);
  }
  const recByRow = new Map<number, number>();
  for (let i = 0; i < rec.row.length; i++) recByRow.set(rec.row[i], i);
  const rows = [...new Set([...rec.row, ...byRow.keys(), ...rowAttrs.keys()])].sort((a, b) => a - b);

  const parts: string[] = [];
  for (const row of rows) {
    const ri = recByRow.get(row);
    const dirtyCols = byRow.get(row);
    const attrs = rowAttrs.get(row);
    if (!dirtyCols) {
      // A row whose height or visibility changed keeps its cells as they were.
      if (!attrs) parts.push(xml.slice(rec.start[ri!], rec.end[ri!]));
      else if (ri === undefined) parts.push(setAttributes(`<${p}row r="${row + 1}"/>`, attrs));
      else parts.push(setAttributes(xml.slice(rec.start[ri], rec.tagEnd[ri] + 1), attrs) + xml.slice(rec.tagEnd[ri] + 1, rec.end[ri]));
      continue;
    }
    let open: string;
    if (ri !== undefined) {
      open = xml.slice(rec.start[ri], rec.tagEnd[ri] + 1).replace(/\s+spans="[^"]*"/, '');
      if (open.endsWith('/>')) open = open.slice(0, -2) + '>';
    } else {
      open = `<${p}row r="${row + 1}">`;
    }
    if (attrs) open = setAttributes(open, attrs);
    const cells: string[] = [];
    const existing = new Map<number, number>();
    if (row < grid.rowCount) for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) existing.set(grid.cols[i], i);
    const cols = [...new Set([...existing.keys(), ...dirtyCols])].sort((a, b) => a - b);
    for (const col of cols) {
      const idx = existing.get(col);
      const raw = idx !== undefined ? xml.slice(scan.cellStart[idx], scan.cellEnd[idx]) : undefined;
      const patch = patches.get(cellKey(row, col));
      if (!patch) cells.push(raw!);
      else {
        const out = patchCell(p, row, col, raw, patch, cm);
        if (out) cells.push(out);
      }
    }
    parts.push(open + cells.join('') + `</${p}row>`);
  }

  const sheetData = `<${p}sheetData>${parts.join('')}</${p}sheetData>`;
  let head = xml.slice(0, scan.sheetDataStart);
  const tail = xml.slice(scan.sheetDataEnd);
  head = updateDimension(head, p, scan, patches);
  return head + sheetData + tail;
}

/** Attributes of the original <c> element other than the ones we rewrite. */
function keptAttrs(raw: string | undefined, drop: string[]): string {
  if (!raw) return '';
  const tagEnd = findTagEnd(raw, 0);
  const selfClosing = raw.charCodeAt(tagEnd - 1) === 47;
  const attrs = raw.slice(raw.indexOf(' '), selfClosing ? tagEnd - 1 : tagEnd);
  const out: string[] = [];
  const re = /([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(attrs))) if (!drop.includes(m[1])) out.push(`${m[1]}=${m[2]}`);
  return out.length ? ' ' + out.join(' ') : '';
}

function valueXml(p: string, v: EngineValue): { t: string; v: string } {
  if (v === null || v === undefined) return { t: '', v: '' };
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return { t: ' t="e"', v: `<${p}v>#NUM!</${p}v>` };
    return { t: '', v: `<${p}v>${v}</${p}v>` };
  }
  if (typeof v === 'boolean') return { t: ' t="b"', v: `<${p}v>${v ? 1 : 0}</${p}v>` };
  if (typeof v === 'string') return { t: ' t="str"', v: `<${p}v>${escapeXml(encodeOoxmlText(v))}</${p}v>` };
  return { t: ' t="e"', v: `<${p}v>${escapeXml(v.error)}</${p}v>` };
}

/** A formula element with its ref attribute set to `ref`. */
function withRef(f: string, ref: string): string {
  if (!f) return f;
  const end = findTagEnd(f, 0);
  const selfClosing = f.charCodeAt(end - 1) === 47;
  const tag = f.slice(0, selfClosing ? end - 1 : end);
  const next = /\sref="[^"]*"/.test(tag) ? tag.replace(/\sref="[^"]*"/, ` ref="${ref}"`) : `${tag} ref="${ref}"`;
  return next + f.slice(selfClosing ? end - 1 : end);
}

function rawFormulaElement(raw: string, p: string): string {
  const f = firstElement(raw, 'f');
  return f && f.qname === `${p}f` ? raw.slice(f.start, f.end) : '';
}

function patchCell(p: string, row: number, col: number, raw: string | undefined, patch: Patch, cm: () => number): string | null {
  const r = toA1(row, col);
  if (patch.kind === 'value') {
    if (!raw) return null;
    const attrs = keptAttrs(raw, ['r', 't']);
    const { t, v } = valueXml(p, patch.value);
    const f = rawFormulaElement(raw, p);
    return `<${p}c r="${r}"${attrs}${t}>${patch.ref ? withRef(f, rangeToA1(patch.ref)) : f}${v}</${p}c>`;
  }
  if (patch.kind === 'spilled') {
    const attrs = keptAttrs(raw, ['r', 's', 't', 'cm', 'vm']);
    const s = patch.style ? ` s="${patch.style}"` : '';
    const { t, v } = valueXml(p, patch.value);
    if (!v) return s || attrs ? `<${p}c r="${r}"${s}${attrs}/>` : null;
    return `<${p}c r="${r}"${s}${attrs}${t}>${v}</${p}c>`;
  }
  if (patch.kind === 'expand') {
    const attrs = keptAttrs(raw, ['r', 't']);
    const { t, v } = valueXml(p, patch.value ?? null);
    return `<${p}c r="${r}"${attrs}${t}><${p}f>${escapeXml(patch.formula)}</${p}f>${v}</${p}c>`;
  }

  const { content, style } = patch.state;
  // cm/vm point at array and rich-value metadata that no longer applies.
  const attrs = keptAttrs(raw, ['r', 's', 't', 'cm', 'vm']);
  const s = style ? ` s="${style}"` : '';
  const open = `<${p}c r="${r}"${s}${attrs}`;
  switch (content.kind) {
    case 'empty':
      return style || attrs ? `${open}/>` : null;
    case 'number':
      return Number.isFinite(content.value)
        ? `${open}><${p}v>${content.value}</${p}v></${p}c>`
        : `${open} t="e"><${p}v>#NUM!</${p}v></${p}c>`;
    case 'string': {
      const text = escapeXml(encodeOoxmlText(content.value));
      const space = /^\s|\s$|\n/.test(content.value) ? ' xml:space="preserve"' : '';
      return `${open} t="inlineStr"><${p}is><${p}t${space}>${text}</${p}t></${p}is></${p}c>`;
    }
    case 'bool':
      return `${open} t="b"><${p}v>${content.value ? 1 : 0}</${p}v></${p}c>`;
    case 'error':
      return `${open} t="e"><${p}v>${escapeXml(content.value)}</${p}v></${p}c>`;
    case 'formula': {
      const { t, v } = valueXml(p, patch.value ?? null);
      if (patch.dynamic) {
        const ref = patch.ref ? rangeToA1(patch.ref) : r;
        return `${open} cm="${cm()}"${t}><${p}f t="array" ref="${ref}">${escapeXml(content.formula)}</${p}f>${v}</${p}c>`;
      }
      return `${open}${t}><${p}f>${escapeXml(content.formula)}</${p}f>${v}</${p}c>`;
    }
  }
}

function updateDimension(head: string, p: string, scan: ScanResult, patches: Map<number, Patch>): string {
  const re = new RegExp(`(<${p}dimension\\b[^>]*\\bref=")([^"]*)(")`);
  if (!re.test(head)) return head;
  let r: Rect | null = scan.maxRow >= 0 ? { r1: firstRow(scan), c1: minCol(scan.grid), r2: scan.maxRow, c2: scan.maxCol } : null;
  for (const [key, patch] of patches) {
    if ((patch.kind === 'edit' && patch.state.content.kind === 'empty') || (patch.kind === 'spilled' && patch.value === null)) continue;
    const row = Math.floor(key / MAX_COLS);
    const col = key % MAX_COLS;
    r = r
      ? { r1: Math.min(r.r1, row), c1: Math.min(r.c1, col), r2: Math.max(r.r2, row), c2: Math.max(r.c2, col) }
      : { r1: row, c1: col, r2: row, c2: col };
  }
  return head.replace(re, (_, a: string, _ref: string, b: string) => `${a}${r ? rangeToA1(r) : 'A1'}${b}`);
}

function firstRow(scan: ScanResult): number {
  const g = scan.grid;
  for (let r = 0; r < g.rowCount; r++) if (g.rowPtr[r + 1] > g.rowPtr[r]) return r;
  return 0;
}

function minCol(g: SheetGrid): number {
  let min = Infinity;
  for (let i = 0; i < g.cols.length; i++) if (g.cols[i] < min) min = g.cols[i];
  return Number.isFinite(min) ? min : 0;
}

/**
 * Turns a macro-enabled workbook into a plain one: drops the VBA project
 * (and its signatures) and switches the workbook's content type to .xlsx.
 */
/**
 * Removes a workbook's macros: the VBA project, and the contents of Excel 4.0
 * macro sheets, which become empty worksheets so that sheet positions (and
 * everything that refers to them) stay as they were. Names that run macros
 * when the file opens or closes go too.
 */
export function removeMacros(pkg: OoxmlPackage, workbookPath: string): void {
  const macroSheets = pkg.rels(workbookPath).filter((r) => r.type === 'xlMacrosheet' || r.type === 'xlIntlMacrosheet');
  for (const r of macroSheets) {
    for (const sub of pkg.rels(r.target)) if (!sub.external) pkg.remove(sub.target);
    pkg.remove(relsPathFor(r.target));
    pkg.setText(r.target, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${MAIN_NS}"><sheetData/></worksheet>`);
  }
  if (macroSheets.length) {
    const relsPath = relsPathFor(workbookPath);
    const relsXml = pkg.text(relsPath);
    if (relsXml) {
      pkg.setText(relsPath, relsXml.replace(/(<(?:\w+:)?Relationship\b[^>]*Type=")[^"]*\/(?:xlMacrosheet|xlIntlMacrosheet)"/g, '$1http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'));
    }
    const ct = pkg.text('[Content_Types].xml');
    if (ct) pkg.setText('[Content_Types].xml', ct.replace(/application\/vnd\.ms-excel\.(?:intl)?macrosheet\+xml/g, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'));
    const wb = pkg.text(workbookPath);
    if (wb) pkg.setText(workbookPath, replaceElements(wb, 'definedName', ({ attrs }, all) => (/\sname="(?:_xlnm\.)?Auto_(?:Open|Close|Activate|Deactivate)\b/i.test(attrs) ? '' : all)));
  }

  const rels = pkg.rels(workbookPath).filter((r) => r.type === 'vbaProject');
  for (const r of rels) {
    for (const sub of pkg.rels(r.target)) if (!sub.external) pkg.remove(sub.target);
    pkg.remove(relsPathFor(r.target));
    pkg.remove(r.target);
  }
  const relsPath = relsPathFor(workbookPath);
  const relsXml = pkg.text(relsPath);
  if (relsXml) pkg.setText(relsPath, relsXml.replace(/<(\w+:)?Relationship\b[^>]*\/vbaProject"[^>]*\/>/g, ''));
  const ct = pkg.text('[Content_Types].xml');
  if (ct) {
    pkg.setText(
      '[Content_Types].xml',
      ct
        .replace(/application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/g, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
        .replace(/<(\w+:)?Override\b[^>]*ContentType="application\/vnd\.ms-office\.vbaProject(Signature\w*)?"[^>]*\/>/g, '')
        .replace(/<(\w+:)?Default\b[^>]*ContentType="application\/vnd\.ms-office\.vbaProject"[^>]*\/>/g, ''),
    );
  }
}

function removeCalcChain(pkg: OoxmlPackage, workbookPath: string): void {
  const chain = pkg.rels(workbookPath).find((r) => r.type === 'calcChain');
  if (!chain) return;
  pkg.remove(chain.target);
  const relsPath = relsPathFor(workbookPath);
  const rels = pkg.text(relsPath);
  if (rels) {
    pkg.setText(relsPath, rels.replace(/<(\w+:)?Relationship\b[^>]*\/calcChain"[^>]*\/>/g, ''));
  }
  const ct = pkg.text('[Content_Types].xml');
  if (ct) {
    pkg.setText(
      '[Content_Types].xml',
      ct.replace(new RegExp(`<(\\w+:)?Override\\b[^>]*PartName="/${chain.target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), ''),
    );
  }
}

/** Asks Excel to recalculate every formula when the file is next opened. */
function forceFullCalcOnLoad(xml: string): string {
  const p = /<(\w+:)?workbook\b/.exec(xml)?.[1] ?? '';
  const calcPr = new RegExp(`<${p}calcPr\\b[^>]*?/?>`);
  const m = calcPr.exec(xml);
  if (m) {
    const tag = /\bfullCalcOnLoad="[^"]*"/.test(m[0])
      ? m[0].replace(/\bfullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"')
      : m[0].replace(new RegExp(`^<${p}calcPr`), `<${p}calcPr fullCalcOnLoad="1"`);
    return xml.replace(m[0], () => tag);
  }
  // calcPr must follow these elements when they are present (CT_Workbook order).
  const tag = `<${p}calcPr fullCalcOnLoad="1"/>`;
  for (const after of ['definedNames', 'externalReferences', 'functionGroups', 'sheets']) {
    const close = new RegExp(`</${p}${after}>|<${p}${after}\\b[^>]*/>`).exec(xml);
    if (close) return xml.slice(0, close.index + close[0].length) + tag + xml.slice(close.index + close[0].length);
  }
  return xml;
}

/**
 * Whether a pivot table's source data may have changed: its range, table or
 * name on a sheet with edits in it. A source this can't find, such as
 * another workbook, counts as changed.
 */
function pivotSourceChanged(xml: string, model: WorkbookModel, input: SaveInput): boolean {
  if (!input.changedIn) return true;
  const source = firstElement(xml, 'worksheetSource');
  if (!source) return true;
  const a = parseAttrs(source.attrs);
  // A source whose sheet was deleted can't be refreshed from.
  if (a.ref === '#REF!') return false;
  let sheet = a.sheet !== undefined ? model.sheets.findIndex((m) => m.info.name === a.sheet) : -1;
  let rect: Rect | null = null;
  const area = (ref: string) => {
    const parsed = parseA1Ref(ref, { allowTernary: false });
    const r = parsed && 'range' in parsed ? parsed.range : undefined;
    if (!r) return null;
    return { r1: r.top ?? 0, c1: r.left ?? 0, r2: r.bottom ?? MAX_ROWS - 1, c2: r.right ?? MAX_COLS - 1 };
  };
  if (a.ref) rect = area(a.ref);
  else if (a.name) {
    const table = model.sheets.findIndex((m) => m.info.tables.some((t) => t.name === a.name));
    const name = model.names.find((n) => n.name === a.name);
    if (table >= 0) {
      sheet = table;
      rect = model.sheets[table].info.tables.find((t) => t.name === a.name)!.rect;
    } else if (name) {
      const parsed = parseA1Ref(name.formula, { allowTernary: false });
      const context = parsed && 'context' in parsed ? parsed.context?.[0] : undefined;
      sheet = context !== undefined ? model.sheets.findIndex((m) => m.info.name === context) : -1;
      rect = area(name.formula.replace(/^.*!/, ''));
    }
  }
  return sheet < 0 || input.changedIn(sheet, rect);
}

const METADATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const METADATA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const DYNAMIC_ARRAY_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const XLDAPR_TYPE =
  '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>';
const XLDAPR_BLOCK = `<bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties xmlns:xda="${DYNAMIC_ARRAY_NS}" fDynamic="1" fCollapsed="0"/></ext></extLst></bk>`;

/**
 * The cell metadata index (a cell's cm attribute) that marks a dynamic array
 * formula, as Excel writes it: an XLDAPR entry in the workbook's metadata
 * part. The part, or the entry, is added when the workbook has none.
 */
function dynamicArrayMetadata(pkg: OoxmlPackage, workbookPath: string): number {
  const rel = pkg.rels(workbookPath).find((r) => r.type === 'sheetMetadata');
  const existing = rel ? pkg.text(rel.target) : undefined;
  if (!rel || !existing) {
    const dir = workbookPath.includes('/') ? workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1) : '';
    let path = `${dir}metadata.xml`;
    for (let n = 2; pkg.has(path); n++) path = `${dir}metadata${n}.xml`;
    pkg.setText(
      path,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<metadata xmlns="${MAIN_NS}"><metadataTypes count="1">${XLDAPR_TYPE}</metadataTypes>` +
        `<futureMetadata name="XLDAPR" count="1">${XLDAPR_BLOCK}</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`,
    );
    const relsPath = relsPathFor(workbookPath);
    const rels = pkg.text(relsPath);
    if (rels) {
      let id = 'rIdWbeMetadata';
      for (let n = 2; rels.includes(`Id="${id}"`); n++) id = `rIdWbeMetadata${n}`;
      pkg.setText(relsPath, rels.replace(/<\/(\w+:)?Relationships>/, (close) => `<Relationship Id="${id}" Type="${METADATA_REL}" Target="${path.slice(dir.length)}"/>${close}`));
    }
    const ct = pkg.text('[Content_Types].xml');
    if (ct) pkg.setText('[Content_Types].xml', ct.replace(/<\/(\w+:)?Types>/, (close) => `<Override PartName="/${path}" ContentType="${METADATA_TYPE}"/>${close}`));
    return 1;
  }

  // The workbook has metadata already: use its XLDAPR entry, or add one.
  let xml = existing;
  const p = /<(\w+:)?metadata\b/.exec(xml)?.[1] ?? '';
  const typesEl = firstElement(xml, 'metadataTypes');
  const types = typesEl?.inner ? [...typesEl.inner.matchAll(/<(?:\w+:)?metadataType\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1]) : [];
  let type = types.indexOf('XLDAPR') + 1;
  if (!type) {
    type = types.length + 1;
    const tag = XLDAPR_TYPE.replace('<metadataType', `<${p}metadataType`);
    xml = typesEl
      ? xml.slice(0, typesEl.start) + `<${typesEl.qname}${typesEl.attrs.replace(/\scount="\d+"/, '')} count="${type}">${typesEl.inner ?? ''}${tag}</${typesEl.qname}>` + xml.slice(typesEl.end)
      : xml.replace(/<(\w+:)?metadata\b[^>]*>/, (open) => `${open}<${p}metadataTypes count="1">${tag}</${p}metadataTypes>`);
  }
  // The block in the XLDAPR future metadata that says "dynamic, not collapsed".
  const future = [...xml.matchAll(/<(?:\w+:)?futureMetadata\b[^>]*\bname="XLDAPR"[^>]*>/g)][0];
  let block = -1;
  if (future) {
    const el = firstElement(xml, 'futureMetadata', future.index);
    block = (el?.inner?.split(/<(?:\w+:)?bk\b/).slice(1) ?? []).findIndex((bk) => /fDynamic="1"/.test(bk) && !/fCollapsed="1"/.test(bk));
    if (el && block < 0) {
      block = countElements(el.inner ?? '', 'bk');
      xml = xml.slice(0, el.start) + `<${el.qname}${el.attrs.replace(/\scount="\d+"/, '')} count="${block + 1}">${el.inner ?? ''}${XLDAPR_BLOCK}</${el.qname}>` + xml.slice(el.end);
    }
  } else {
    block = 0;
    const at = firstElement(xml, 'cellMetadata')?.start ?? firstElement(xml, 'valueMetadata')?.start ?? xml.lastIndexOf('</');
    xml = xml.slice(0, at) + `<${p}futureMetadata name="XLDAPR" count="1">${XLDAPR_BLOCK}</${p}futureMetadata>` + xml.slice(at);
  }
  // The cell metadata entry pointing at that block; its position is the cm value.
  const cells = firstElement(xml, 'cellMetadata');
  const entries = cells?.inner?.split(/<(?:\w+:)?bk\b/).slice(1) ?? [];
  let cm = entries.findIndex((bk) => new RegExp(`\\bt="${type}"`).test(bk) && new RegExp(`\\bv="${block}"`).test(bk)) + 1;
  if (!cm) {
    cm = entries.length + 1;
    const entry = `<${p}bk><${p}rc t="${type}" v="${block}"/></${p}bk>`;
    xml = cells
      ? xml.slice(0, cells.start) + `<${cells.qname}${cells.attrs.replace(/\scount="\d+"/, '')} count="${cm}">${cells.inner ?? ''}${entry}</${cells.qname}>` + xml.slice(cells.end)
      : (() => {
          const at = firstElement(xml, 'valueMetadata')?.start ?? firstElement(xml, 'extLst')?.start ?? xml.lastIndexOf('</');
          return xml.slice(0, at) + `<${p}cellMetadata count="1">${entry}</${p}cellMetadata>` + xml.slice(at);
        })();
  }
  pkg.setText(rel.target, xml);
  return cm;
}

function setRootAttr(xml: string, name: string, value: string): string {
  const m = /<(?!\?)[^>]+>/.exec(xml);
  if (!m) return xml;
  const tag = m[0];
  const re = new RegExp(`\\b${name}="[^"]*"`);
  const next = re.test(tag) ? tag.replace(re, `${name}="${value}"`) : tag.replace(/^<([^\s>/]+)/, `<$1 ${name}="${value}"`);
  return xml.replace(tag, () => next);
}

function appendStyles(xml: string, model: WorkbookModel): string {
  const p = /<(\w+:)?styleSheet\b/.exec(xml)?.[1] ?? '';
  const adds = model.styles.additions;
  let out = xml;

  const numFmts = adds.filter((a) => a.numFmt).map((a) => `<${p}numFmt numFmtId="${a.numFmt!.id}" formatCode="${escapeXml(a.numFmt!.code)}"/>`);
  if (numFmts.length) {
    const block = firstElement(out, 'numFmts');
    if (block) {
      const body = block.inner ?? '';
      const count = countElements(body, 'numFmt') + numFmts.length;
      const openTag = `<${block.qname}${block.attrs.replace(/\s*count="\d+"/, '')} count="${count}">`;
      out = out.slice(0, block.start) + openTag + body + numFmts.join('') + `</${block.qname}>` + out.slice(block.end);
    } else {
      const root = new RegExp(`<${p}styleSheet\\b[^>]*>`).exec(out)!;
      const at = root.index + root[0].length;
      out = out.slice(0, at) + `<${p}numFmts count="${numFmts.length}">${numFmts.join('')}</${p}numFmts>` + out.slice(at);
    }
  }

  // Fonts, fills and borders new formats use, after the file's, whose ids they follow on from.
  for (const kind of ['font', 'fill', 'border'] as const) {
    const added = model.styles.addedParts[kind];
    if (!added.length) continue;
    const block = firstElement(out, `${kind}s`);
    if (block?.inner == null) continue;
    const count = countElements(block.inner, kind) + added.length;
    const openTag = `<${block.qname}${block.attrs.replace(/\s*count="\d+"/, '')} count="${count}">`;
    out = out.slice(0, block.start) + openTag + block.inner + added.join('') + `</${block.qname}>` + out.slice(block.end);
  }

  const xfs = adds.map((a) => a.xfRaw);
  const close = new RegExp(`</${p}cellXfs>`).exec(out);
  if (close) {
    out = out.slice(0, close.index) + xfs.join('') + out.slice(close.index);
    out = out.replace(new RegExp(`(<${p}cellXfs\\b[^>]*\\bcount=")(\\d+)(")`), (_, a: string, _n: string, b: string) => `${a}${model.styles.display.length}${b}`);
  }
  return out;
}

