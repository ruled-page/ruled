import type { EngineValue } from './calc';
import { cellKey, MAX_COLS, rangeToA1, toA1 } from './refs';
import type { CellState } from './session';
import type { ScanResult } from './sheetScanner';
import { CellFlag, VT, type Rect, type SheetGrid } from './types';
import type { SheetModel, WorkbookModel } from './workbook';
import { encodeOoxmlText, escapeXml, findTagEnd } from './xml';
import { relsPathFor, type OoxmlPackage } from './zip';

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
}

type Patch =
  | { kind: 'edit'; state: CellState; value?: EngineValue }
  | { kind: 'value'; value: EngineValue }
  | { kind: 'expand'; formula: string; value: EngineValue | undefined };

/** Writes the edited workbook, changing only the parts that must change. */
export function writeWorkbook(input: SaveInput): Uint8Array {
  const { pkg, model } = input;

  model.sheets.forEach((m, s) => {
    if (!m.scan || m.xml === undefined) return;
    const patches = sheetPatches(m, s, input);
    if (patches.size) pkg.setText(m.path, rewriteSheet(m.xml, m.scan, patches));
  });

  // Excel checks calcChain.xml against the cells that hold formulas and
  // reports the file as damaged if they disagree. It rebuilds the chain
  // itself when the part is missing.
  removeCalcChain(pkg, model.workbookPath);

  const wbXml = pkg.text(model.workbookPath);
  if (wbXml) pkg.setText(model.workbookPath, forceFullCalcOnLoad(wbXml));

  for (const path of pkg.list()) {
    if (/pivotCacheDefinition\d*\.xml$/.test(path)) {
      const xml = pkg.text(path);
      if (xml) pkg.setText(path, setRootAttr(xml, 'refreshOnLoad', '1'));
    }
  }

  if (input.stripMacros) removeMacros(pkg, model.workbookPath);

  if (model.styles.additions.length) {
    const stylesRel = pkg.rels(model.workbookPath).find((r) => r.type === 'styles');
    const xml = stylesRel ? pkg.text(stylesRel.target) : undefined;
    if (stylesRel && xml) pkg.setText(stylesRel.target, appendStyles(xml, model));
  }

  return pkg.toBytes();
}

function sheetPatches(m: SheetModel, s: number, input: SaveInput): Map<number, Patch> {
  const patches = new Map<number, Patch>();
  const grid = m.scan!.grid;
  for (const [key, state] of input.edits[s]) {
    const value = state.content.kind === 'formula' ? input.cachedValue(s, Math.floor(key / MAX_COLS), key % MAX_COLS) : undefined;
    patches.set(key, { kind: 'edit', state, value });
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
  return patches;
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

function rewriteSheet(xml: string, scan: ScanResult, patches: Map<number, Patch>): string {
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
  const rows = [...new Set([...rec.row, ...byRow.keys()])].sort((a, b) => a - b);

  const parts: string[] = [];
  for (const row of rows) {
    const ri = recByRow.get(row);
    const dirtyCols = byRow.get(row);
    if (!dirtyCols) {
      parts.push(xml.slice(rec.start[ri!], rec.end[ri!]));
      continue;
    }
    let open: string;
    if (ri !== undefined) {
      open = xml.slice(rec.start[ri], rec.tagEnd[ri] + 1).replace(/\s+spans="[^"]*"/, '');
      if (open.endsWith('/>')) open = open.slice(0, -2) + '>';
    } else {
      open = `<${p}row r="${row + 1}">`;
    }
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
        const out = patchCell(p, row, col, raw, patch);
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

function rawFormulaElement(raw: string, p: string): string {
  const m = new RegExp(`<${p}f\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}f>)`).exec(raw);
  return m ? m[0] : '';
}

function patchCell(p: string, row: number, col: number, raw: string | undefined, patch: Patch): string | null {
  const r = toA1(row, col);
  if (patch.kind === 'value') {
    if (!raw) return null;
    const attrs = keptAttrs(raw, ['r', 't']);
    const { t, v } = valueXml(p, patch.value);
    return `<${p}c r="${r}"${attrs}${t}>${rawFormulaElement(raw, p)}${v}</${p}c>`;
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
      return `${open}${t}><${p}f>${escapeXml(content.formula)}</${p}f>${v}</${p}c>`;
    }
  }
}

function updateDimension(head: string, p: string, scan: ScanResult, patches: Map<number, Patch>): string {
  const re = new RegExp(`(<${p}dimension\\b[^>]*\\bref=")([^"]*)(")`);
  if (!re.test(head)) return head;
  let r: Rect | null = scan.maxRow >= 0 ? { r1: firstRow(scan), c1: minCol(scan.grid), r2: scan.maxRow, c2: scan.maxCol } : null;
  for (const [key, patch] of patches) {
    if (patch.kind === 'edit' && patch.state.content.kind === 'empty') continue;
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
export function removeMacros(pkg: OoxmlPackage, workbookPath: string): void {
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
    const block = new RegExp(`(<${p}numFmts\\b[^>]*?)(?:count="(\\d+)")?([^>]*>)([\\s\\S]*?)(</${p}numFmts>)`);
    const m = block.exec(out);
    if (m) {
      const count = (m[0].match(new RegExp(`<${p}numFmt\\b`, 'g')) ?? []).length + numFmts.length;
      const openTag = m[0].slice(0, m[0].indexOf('>') + 1).replace(/\s*count="\d+"/, '').replace(/>$/, ` count="${count}">`);
      out = out.replace(m[0], () => openTag + m[4] + numFmts.join('') + m[5]);
    } else {
      const root = new RegExp(`<${p}styleSheet\\b[^>]*>`).exec(out)!;
      const at = root.index + root[0].length;
      out = out.slice(0, at) + `<${p}numFmts count="${numFmts.length}">${numFmts.join('')}</${p}numFmts>` + out.slice(at);
    }
  }

  const xfs = adds.map((a) => a.xfRaw);
  const close = new RegExp(`</${p}cellXfs>`).exec(out);
  if (close) {
    out = out.slice(0, close.index) + xfs.join('') + out.slice(close.index);
    out = out.replace(new RegExp(`(<${p}cellXfs\\b[^>]*\\bcount=")(\\d+)(")`), (_, a: string, _n: string, b: string) => `${a}${model.styles.display.length}${b}`);
  }
  return out;
}

