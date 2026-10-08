import { makeShifter } from './formula';
import { MAX_COLS, MAX_ROWS, parseRange } from './refs';
import { CellFlag, VT, type Rect, type SheetGrid } from './types';
import { damaged, decodeEntities, decodeOoxmlText, elements, findTagEnd, parseAttrs } from './xml';

export interface FormulaMeta {
  kind: 'normal' | 'shared-master' | 'shared-follower' | 'array' | 'dataTable';
  si?: number;
  ref?: Rect;
}

/**
 * Everything needed to display a worksheet and later rewrite only the cells
 * that changed: parsed values plus the character offsets of every row and
 * cell element in the original XML.
 */
export interface ScanResult {
  prefix: string;
  grid: SheetGrid;
  cellStart: Int32Array;
  cellEnd: Int32Array;
  /** Row elements in document order. */
  rowRec: { row: Int32Array; start: Int32Array; tagEnd: Int32Array; end: Int32Array };
  fmeta: Map<number, FormulaMeta>;
  /** Shared formula group id → master cell index and follower cell indices. */
  sharedGroups: Map<number, { master: number; members: number[] }>;
  /** Multi-cell array formulas, spill ranges and data tables. */
  arrayRanges: (Rect & { master: number })[];
  /** Offsets of the whole <sheetData> element, or -1 when it is missing. */
  sheetDataStart: number;
  sheetDataEnd: number;
  rowStyles: Map<number, number>;
  rowHeightsPt: Map<number, number>;
  hiddenRows: number[];
  maxRow: number;
  maxCol: number;
  richTextCells: number;
}

function isBoundary(code: number): boolean {
  return code === 32 || code === 62 || code === 47 || code === 9 || code === 10 || code === 13;
}

function indexOfTag(xml: string, tag: string, from: number, limit: number): number {
  let i = from;
  for (;;) {
    i = xml.indexOf(tag, i);
    if (i === -1 || i >= limit) return -1;
    if (isBoundary(xml.charCodeAt(i + tag.length))) return i;
    i++;
  }
}

interface Element {
  attrs: string;
  inner: string | null;
}

/** First element called `tag` (with prefix) inside xml[from, to). */
function findElement(xml: string, tag: string, from: number, to: number): Element | null {
  const s = indexOfTag(xml, `<${tag}`, from, to);
  if (s === -1) return null;
  const gt = findTagEnd(xml, s);
  if (xml.charCodeAt(gt - 1) === 47) return { attrs: xml.slice(s + tag.length + 1, gt - 1), inner: null };
  const close = xml.indexOf(`</${tag}>`, gt);
  return { attrs: xml.slice(s + tag.length + 1, gt), inner: xml.slice(gt + 1, close === -1 ? to : close) };
}

/** Plain text of a shared or inline string item, ignoring phonetic runs. */
export function stringItemText(inner: string): string {
  let out = '';
  // Phonetic runs (<rPh>) hold their own <t>, which isn't part of the text.
  const skip = inner.indexOf('rPh') === -1 ? [] : elements(inner, 'rPh');
  let k = 0;
  for (const t of elements(inner, 't')) {
    while (k < skip.length && skip[k].end <= t.start) k++;
    if (k < skip.length && skip[k].start < t.start) continue;
    out += t.inner ?? '';
  }
  return decodeOoxmlText(decodeEntities(out));
}

export function parseSharedStrings(xml: string | undefined): { strings: string[]; rich: Set<number> } {
  const strings: string[] = [];
  const rich = new Set<number>();
  if (!xml) return { strings, rich };
  const p = /<(\w+:)?sst\b/.exec(xml)?.[1] ?? '';
  for (const si of elements(xml, 'si')) {
    const inner = si.inner ?? '';
    if (inner.indexOf(`<${p}r>`) !== -1) rich.add(strings.length);
    strings.push(stringItemText(inner));
  }
  return { strings, rich };
}

/** Matches the scanning cursor's `{r, s, t}` attributes without allocating. */
const cellAttrs = { r: '', s: '', t: '' };
function readCellAttrs(s: string): void {
  cellAttrs.r = cellAttrs.s = cellAttrs.t = '';
  let i = 0;
  const n = s.length;
  while (i < n) {
    while (i < n && s.charCodeAt(i) <= 32) i++;
    if (i >= n) break;
    const eq = s.indexOf('=', i);
    if (eq === -1) break;
    const name = s.slice(i, eq).trim();
    let q = eq + 1;
    while (s.charCodeAt(q) <= 32) q++;
    const end = s.indexOf(s[q], q + 1);
    if (end === -1) break;
    if (name === 'r') cellAttrs.r = s.slice(q + 1, end);
    else if (name === 's') cellAttrs.s = s.slice(q + 1, end);
    else if (name === 't') cellAttrs.t = s.slice(q + 1, end);
    i = end + 1;
  }
}

function colOfRef(ref: string): number {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

function isoToSerial(iso: string): number {
  const ms = Date.parse(/[zZ]|[+-]\d\d:\d\d$/.test(iso) || iso.length <= 10 ? iso : iso + 'Z');
  return ms / 86_400_000 + 25_569;
}

export function scanSheet(xml: string, sst: string[], richStrings: Set<number> = new Set()): ScanResult {
  const p = /<(\w+:)?worksheet\b/.exec(xml)?.[1] ?? '';
  const rowsA: number[] = [];
  const colsA: number[] = [];
  const typesA: number[] = [];
  const numsA: number[] = [];
  const strsA: (string | undefined)[] = [];
  const stylesA: number[] = [];
  const flagsA: number[] = [];
  const cStart: number[] = [];
  const cEnd: number[] = [];
  const formulaText = new Map<number, string>();
  const fmeta = new Map<number, FormulaMeta>();
  const rowRow: number[] = [];
  const rowStart: number[] = [];
  const rowTagEnd: number[] = [];
  const rowEnd: number[] = [];
  const rowStyles = new Map<number, number>();
  const rowHeightsPt = new Map<number, number>();
  const hiddenRows: number[] = [];
  let richTextCells = 0;

  const sdOpen = indexOfTag(xml, `<${p}sheetData`, 0, xml.length);
  let sheetDataStart = -1;
  let sheetDataEnd = -1;
  let needsSort = false;
  let lastKey = -1;

  if (sdOpen !== -1) {
    sheetDataStart = sdOpen;
    const sdTagEnd = findTagEnd(xml, sdOpen);
    if (xml.charCodeAt(sdTagEnd) !== 62) throw damaged();
    let limit: number;
    if (xml.charCodeAt(sdTagEnd - 1) === 47) {
      sheetDataEnd = sdTagEnd + 1;
      limit = sdTagEnd;
    } else {
      const close = xml.indexOf(`</${p}sheetData>`, sdTagEnd);
      if (close === -1) throw damaged();
      limit = close;
      sheetDataEnd = close + `</${p}sheetData>`.length;
    }

    const rowTag = `<${p}row`;
    const rowClose = `</${p}row>`;
    const cTag = `<${p}c`;
    const cClose = `</${p}c>`;
    const vTag = `${p}v`;
    const fTag = `${p}f`;
    const isTag = `${p}is`;
    let pos = sdTagEnd + 1;
    let lastRow = -1;

    while (pos < limit) {
      const rs = indexOfTag(xml, rowTag, pos, limit);
      if (rs === -1) break;
      const tagEnd = findTagEnd(xml, rs);
      if (xml.charCodeAt(tagEnd) !== 62 || tagEnd >= limit) throw damaged();
      const selfClosing = xml.charCodeAt(tagEnd - 1) === 47;
      const ra = parseAttrs(xml.slice(rs + rowTag.length, selfClosing ? tagEnd - 1 : tagEnd));
      const row = ra.r ? parseInt(ra.r, 10) - 1 : lastRow + 1;
      // Excel's sheets end at row 1,048,576; a larger number would size the grid by it.
      if (!(row >= 0 && row < MAX_ROWS)) throw damaged();
      lastRow = row;
      if (ra.customFormat === '1' && ra.s) rowStyles.set(row, Number(ra.s));
      if (ra.ht && (ra.customHeight === '1' || ra.customHeight === 'true')) rowHeightsPt.set(row, Number(ra.ht));
      if (ra.hidden === '1' || ra.hidden === 'true') hiddenRows.push(row);

      let end: number;
      if (selfClosing) {
        end = tagEnd + 1;
      } else {
        const rc = xml.indexOf(rowClose, tagEnd);
        if (rc === -1 || rc > limit) throw damaged();
        end = rc + rowClose.length;
        let cpos = tagEnd + 1;
        let lastCol = -1;
        while (cpos < rc) {
          const cs = indexOfTag(xml, cTag, cpos, rc);
          if (cs === -1) break;
          const cTagEnd = findTagEnd(xml, cs);
          if (xml.charCodeAt(cTagEnd) !== 62 || cTagEnd >= rc) throw damaged();
          const cSelf = xml.charCodeAt(cTagEnd - 1) === 47;
          readCellAttrs(xml.slice(cs + cTag.length, cSelf ? cTagEnd - 1 : cTagEnd));
          const cClosePos = cSelf ? cTagEnd : xml.indexOf(cClose, cTagEnd);
          if (cClosePos === -1 || cClosePos > rc) throw damaged();
          const cellEndPos = cSelf ? cTagEnd + 1 : cClosePos + cClose.length;
          const col = cellAttrs.r ? colOfRef(cellAttrs.r) : lastCol + 1;
          if (!(col >= 0 && col < MAX_COLS)) throw damaged();
          lastCol = col;
          const idx = rowsA.length;
          const key = row * 16384 + col;
          if (key <= lastKey) needsSort = true;
          lastKey = key;

          let type: number = VT.None;
          let num = 0;
          let str: string | undefined;
          let flags = 0;
          if (!cSelf) {
            const innerEnd = cellEndPos - cClose.length;
            const t = cellAttrs.t;
            const v = findElement(xml, vTag, cTagEnd + 1, innerEnd);
            const vText = v?.inner ?? undefined;
            if (t === 's') {
              if (vText !== undefined) {
                const i = parseInt(vText, 10);
                type = VT.String;
                str = sst[i] ?? '';
                if (richStrings.has(i)) richTextCells++;
              }
            } else if (t === 'inlineStr') {
              const is = findElement(xml, isTag, cTagEnd + 1, innerEnd);
              type = VT.String;
              str = is?.inner ? stringItemText(is.inner) : '';
            } else if (vText !== undefined) {
              if (t === 'str') {
                type = VT.String;
                str = decodeOoxmlText(decodeEntities(vText));
              } else if (t === 'b') {
                type = VT.Bool;
                num = vText.trim() === '1' || vText.trim() === 'true' ? 1 : 0;
              } else if (t === 'e') {
                type = VT.Error;
                str = decodeEntities(vText).trim();
              } else if (t === 'd') {
                type = VT.Number;
                num = isoToSerial(vText.trim());
              } else {
                const n = Number(vText);
                if (!Number.isNaN(n) && vText.trim() !== '') {
                  type = VT.Number;
                  num = n;
                }
              }
            }

            const f = findElement(xml, fTag, cTagEnd + 1, innerEnd);
            if (f) {
              const fa = parseAttrs(f.attrs);
              const text = f.inner !== null ? decodeEntities(f.inner) : '';
              const ft = fa.t ?? 'normal';
              flags |= CellFlag.Formula;
              if (vText === undefined) flags |= CellFlag.NoCache;
              if (ft === 'shared') {
                const si = Number(fa.si);
                if (text) {
                  fmeta.set(idx, { kind: 'shared-master', si, ref: fa.ref ? parseRange(fa.ref) ?? undefined : undefined });
                  formulaText.set(idx, text);
                } else {
                  fmeta.set(idx, { kind: 'shared-follower', si });
                }
              } else if (ft === 'array' || ft === 'dataTable') {
                fmeta.set(idx, { kind: ft, ref: fa.ref ? parseRange(fa.ref) ?? undefined : undefined });
                formulaText.set(idx, text);
              } else {
                fmeta.set(idx, { kind: 'normal' });
                formulaText.set(idx, text);
              }
            }
          }

          rowsA.push(row);
          colsA.push(col);
          typesA.push(type);
          numsA.push(num);
          strsA.push(str);
          stylesA.push(cellAttrs.s ? Number(cellAttrs.s) : 0);
          flagsA.push(flags);
          cStart.push(cs);
          cEnd.push(cellEndPos);
          cpos = cellEndPos;
        }
      }
      rowRow.push(row);
      rowStart.push(rs);
      rowTagEnd.push(tagEnd);
      rowEnd.push(end);
      pos = end;
    }
  }

  // Cells must be in row/column order for lookups. Excel always writes them
  // that way, but other producers may not; sort and drop duplicates if so.
  let order: number[] | null = null;
  if (needsSort) {
    const idxs = rowsA.map((_, i) => i);
    idxs.sort((a, b) => rowsA[a] * 16384 + colsA[a] - (rowsA[b] * 16384 + colsA[b]) || a - b);
    order = [];
    for (let i = 0; i < idxs.length; i++) {
      const next = idxs[i + 1];
      if (next !== undefined && rowsA[next] === rowsA[idxs[i]] && colsA[next] === colsA[idxs[i]]) continue;
      order.push(idxs[i]);
    }
  }
  const pick = <T,>(arr: T[]): T[] => (order ? order.map((i) => arr[i]) : arr);
  const remap = new Map<number, number>();
  if (order) order.forEach((oldIdx, newIdx) => remap.set(oldIdx, newIdx));
  const newIndex = (i: number) => (order ? remap.get(i) : i);

  const rows = pick(rowsA);
  const n = rows.length;
  const cols = Int32Array.from(pick(colsA));
  const flags = Uint8Array.from(pick(flagsA));
  const formulas = new Map<number, string>();
  const meta = new Map<number, FormulaMeta>();
  for (const [i, text] of formulaText) {
    const j = newIndex(i);
    if (j !== undefined) formulas.set(j, text);
  }
  for (const [i, m] of fmeta) {
    const j = newIndex(i);
    if (j !== undefined) meta.set(j, m);
  }

  let maxRow = -1;
  let maxCol = -1;
  for (let i = 0; i < n; i++) {
    if (rows[i] > maxRow) maxRow = rows[i];
    if (cols[i] > maxCol) maxCol = cols[i];
  }
  const rowCount = maxRow + 1;
  const rowPtr = new Int32Array(rowCount + 1);
  for (let i = 0; i < n; i++) rowPtr[rows[i] + 1]++;
  for (let r = 0; r < rowCount; r++) rowPtr[r + 1] += rowPtr[r];

  const grid: SheetGrid = {
    rowCount,
    rowPtr,
    cols,
    types: Uint8Array.from(pick(typesA)),
    nums: Float64Array.from(pick(numsA)),
    strs: pick(strsA),
    styles: Int32Array.from(pick(stylesA)),
    flags,
    formulas,
  };

  // Expand shared formulas so every follower has its own formula text.
  const sharedGroups = new Map<number, { master: number; members: number[] }>();
  for (const [i, m] of meta) {
    if (m.si === undefined) continue;
    let g = sharedGroups.get(m.si);
    if (!g) sharedGroups.set(m.si, (g = { master: -1, members: [] }));
    if (m.kind === 'shared-master') g.master = i;
    else g.members.push(i);
  }
  for (const g of sharedGroups.values()) {
    const masterText = g.master >= 0 ? formulas.get(g.master) : undefined;
    let shift: ((r: number, c: number) => string) | undefined;
    if (masterText !== undefined) {
      try {
        shift = makeShifter(masterText, rows[g.master], cols[g.master]);
      } catch {
        shift = undefined;
      }
    }
    for (const i of g.members) {
      if (shift) formulas.set(i, shift(rows[i], cols[i]));
      else flags[i] |= CellFlag.Unsupported;
    }
  }

  // Multi-cell array formulas, spills and data tables are kept as values.
  const arrayRanges: (Rect & { master: number })[] = [];
  for (const [i, m] of meta) {
    if ((m.kind !== 'array' && m.kind !== 'dataTable') || !m.ref) continue;
    const single = m.ref.r1 === m.ref.r2 && m.ref.c1 === m.ref.c2;
    if (single && m.kind === 'array') continue;
    arrayRanges.push({ ...m.ref, master: i });
    flags[i] |= CellFlag.Unsupported | CellFlag.ArrayMember;
    for (let r = m.ref.r1; r <= Math.min(m.ref.r2, rowCount - 1); r++) {
      for (let k = rowPtr[r]; k < rowPtr[r + 1]; k++) {
        if (cols[k] >= m.ref.c1 && cols[k] <= m.ref.c2) flags[k] |= CellFlag.ArrayMember;
      }
    }
  }

  return {
    prefix: p,
    grid,
    cellStart: Int32Array.from(pick(cStart)),
    cellEnd: Int32Array.from(pick(cEnd)),
    rowRec: {
      row: Int32Array.from(rowRow),
      start: Int32Array.from(rowStart),
      tagEnd: Int32Array.from(rowTagEnd),
      end: Int32Array.from(rowEnd),
    },
    fmeta: meta,
    sharedGroups,
    arrayRanges,
    sheetDataStart,
    sheetDataEnd,
    rowStyles,
    rowHeightsPt,
    hiddenRows,
    maxRow,
    maxCol,
    richTextCells,
  };
}

/** Index of the cell at (row, col) in the grid, or -1. */
export function findCell(grid: SheetGrid, row: number, col: number): number {
  if (row < 0 || row >= grid.rowCount) return -1;
  let lo = grid.rowPtr[row];
  let hi = grid.rowPtr[row + 1] - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const c = grid.cols[mid];
    if (c === col) return mid;
    if (c < col) lo = mid + 1;
    else hi = mid - 1;
  }
  return -1;
}


