import { parseA1Ref, tokenize } from '@borgar/fx';
import { DetailedCellError, HyperFormula, type RawCellContent, type SerializedNamedExpression } from 'hyperformula';
import { engineFormula, stripFunctionPrefixes } from './formula';
import type { CellContent } from './types';

let registered: Set<string> | undefined;
let probe: HyperFormula | undefined;

/**
 * Why the engine cannot calculate this formula, or undefined when it can.
 * Checked up front so a formula is judged on its own content, not on errors
 * it inherits from the cells it references.
 */
export function unsupportedReason(fileFormula: string): string | undefined {
  registered ??= new Set(HyperFormula.getRegisteredFunctionNames('enGB'));
  let tokens;
  try {
    tokens = tokenize(fileFormula, { mergeRefs: true, allowTernary: false });
  } catch {
    return 'Formula syntax';
  }
  for (const t of tokens) {
    if (t.type === 'structured') return 'Table references (Table[Column])';
    if (t.type === 'func') {
      const name = stripFunctionPrefixes(t.value).toUpperCase();
      if (!registered.has(name)) return `${name.replace(/^_XLUDF\./, '')} function`;
    }
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_ternary' || t.type === 'unknown') {
      if (t.value.includes('[')) return 'Links to other workbooks';
      const ref = parseA1Ref(t.value, { allowTernary: false });
      if (ref && 'context' in ref && (ref.context?.length ?? 0) > 1) return '3-D references (Sheet1:Sheet3!A1)';
    }
  }
  probe ??= HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  if (!probe.validateFormula(engineFormula(fileFormula))) return 'Formula syntax';
  return undefined;
}

export type EngineValue = number | string | boolean | { error: string; message?: string } | null;

export interface EngineSheet {
  name: string;
  rows: RawCellContent[][];
}

export interface EngineName {
  name: string;
  /** File-form formula text, e.g. "Data!$A$1:$A$10". */
  formula: string;
  localSheet?: number;
}

export interface EngineWrite {
  sheet: number;
  row: number;
  col: number;
  raw: RawCellContent;
}

export interface EngineChange {
  sheet: number;
  row: number;
  col: number;
  value: EngineValue;
}

/** Errors the engine reports when it cannot interpret a formula at all. */
export function isUnsupportedError(v: EngineValue): v is { error: string; message?: string } {
  if (!v || typeof v !== 'object') return false;
  if (v.error === '#ERROR!') return true;
  return v.error === '#NAME?' && /not recognized|not supported/i.test(v.message ?? '');
}

export function toRaw(content: CellContent): RawCellContent {
  switch (content.kind) {
    case 'empty':
      return null;
    case 'number':
      return content.value;
    case 'bool':
      return content.value;
    // A leading apostrophe stops the engine from reading "007" or "1/2" as numbers.
    case 'string':
      return "'" + content.value;
    case 'error':
      return content.value;
    case 'formula':
      return engineFormula(content.formula);
  }
}

function convert(v: unknown): EngineValue {
  if (v instanceof DetailedCellError) return { error: v.value, message: v.message };
  if (v === undefined || v === null) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  return null;
}

/**
 * HyperFormula configured to behave like Excel: 1900 date system with the
 * Lotus leap-year bug, empty cells as zero in arithmetic, full sheet size.
 */
export class CalcEngine {
  private readonly hf: HyperFormula;
  /** Our sheet index → engine sheet id (-1 when the sheet is not a worksheet). */
  private readonly ids: number[];
  private readonly sheetOf = new Map<number, number>();

  constructor(sheets: (EngineSheet | null)[], names: EngineName[], date1904: boolean) {
    const config = {
      licenseKey: 'gpl-v3',
      maxRows: 1_048_576,
      maxColumns: 16_384,
      // The 1899-12-30 base date already absorbs Excel's phantom 29 Feb 1900.
      leapYear1900: false,
      nullDate: date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 30 },
      evaluateNullToZero: true,
      smartRounding: true,
      useArrayArithmetic: true,
      useWildcards: true,
      useRegularExpressions: false,
      matchWholeCell: true,
      caseSensitive: false,
      undoLimit: 0,
    };

    const data: Record<string, RawCellContent[][]> = {};
    for (const s of sheets) if (s) data[s.name] = s.rows;

    // Probe names against an empty copy of the workbook so one bad name
    // cannot stop the whole engine from building.
    const probe = HyperFormula.buildFromSheets(
      Object.fromEntries(Object.keys(data).map((n) => [n, []])),
      config,
    );
    const valid: SerializedNamedExpression[] = [];
    for (const n of names) {
      if (n.name.startsWith('_xlnm.') || /#REF!/.test(n.formula)) continue;
      const scopeName = n.localSheet !== undefined ? sheets[n.localSheet]?.name : undefined;
      if (n.localSheet !== undefined && !scopeName) continue;
      const scope = scopeName !== undefined ? probe.getSheetId(scopeName) : undefined;
      const expression = engineFormula(n.formula);
      try {
        if (probe.isItPossibleToAddNamedExpression(n.name, expression, scope)) {
          probe.addNamedExpression(n.name, expression, scope);
          valid.push({ name: n.name, expression, scope });
        }
      } catch {
        // Names the engine cannot represent are skipped; formulas using them
        // will be reported as not calculated here.
      }
    }
    probe.destroy();

    this.hf = HyperFormula.buildFromSheets(data, config, valid);
    this.ids = sheets.map((s) => (s ? this.hf.getSheetId(s.name) ?? -1 : -1));
    this.ids.forEach((id, i) => {
      if (id >= 0) this.sheetOf.set(id, i);
    });
  }

  value(sheet: number, row: number, col: number): EngineValue {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return null;
    return convert(this.hf.getCellValue({ sheet: id, row, col }));
  }

  /** Applies writes in one batch and returns every cell whose value changed. */
  write(writes: EngineWrite[]): EngineChange[] {
    if (!writes.length) return [];
    const changes = this.hf.batch(() => {
      for (const w of writes) {
        const id = this.ids[w.sheet];
        if (id === undefined || id < 0) continue;
        try {
          this.hf.setCellContents({ sheet: id, row: w.row, col: w.col }, w.raw);
        } catch {
          // Content the engine rejects (e.g. an unparseable formula) is stored as text.
          this.hf.setCellContents({ sheet: id, row: w.row, col: w.col }, "'" + String(w.raw));
        }
      }
    });
    const out: EngineChange[] = [];
    for (const c of changes) {
      if (!('address' in c) || !c.address) continue;
      const sheet = this.sheetOf.get(c.address.sheet);
      if (sheet === undefined) continue;
      out.push({ sheet, row: c.address.row, col: c.address.col, value: convert(c.newValue) });
    }
    return out;
  }

  /**
   * Formula cells that depend on any of the given cells, directly or
   * indirectly. Used to mark results the engine cannot vouch for.
   */
  dependentsOf(cells: { sheet: number; row: number; col: number }[], limit = 200_000): { sheet: number; row: number; col: number }[] {
    type Node = { sheet: number; row: number; col: number } | { start: { sheet: number; row: number; col: number }; end: { sheet: number; row: number; col: number } };
    const seen = new Set<string>();
    const out: { sheet: number; row: number; col: number }[] = [];
    const queue: Node[] = [];
    for (const c of cells) {
      const id = this.ids[c.sheet];
      if (id !== undefined && id >= 0) queue.push({ sheet: id, row: c.row, col: c.col });
    }
    for (let qi = 0; qi < queue.length && out.length < limit; qi++) {
      const node = queue[qi];
      let deps: Node[];
      try {
        deps = this.hf.getCellDependents(node as never) as Node[];
      } catch {
        continue;
      }
      for (const d of deps) {
        const id = 'start' in d ? `r${d.start.sheet}:${d.start.row}:${d.start.col}:${d.end.row}:${d.end.col}` : `c${d.sheet}:${d.row}:${d.col}`;
        if (seen.has(id)) continue;
        seen.add(id);
        queue.push(d);
        if (!('start' in d)) {
          const sheet = this.sheetOf.get(d.sheet);
          if (sheet !== undefined) out.push({ sheet, row: d.row, col: d.col });
        }
      }
    }
    return out;
  }

  destroy(): void {
    this.hf.destroy();
  }
}

export function valuesMatch(excel: EngineValue, engine: EngineValue): boolean {
  if (excel === null || excel === '') return engine === null || engine === '' || engine === 0;
  if (typeof excel === 'number' && typeof engine === 'number') {
    const scale = Math.max(1, Math.abs(excel), Math.abs(engine));
    return Math.abs(excel - engine) <= 1e-9 * scale;
  }
  if (typeof excel === 'object' && typeof engine === 'object' && engine) return excel.error === engine.error;
  return excel === engine;
}

export function describeValue(v: EngineValue): string {
  if (v === null) return '(empty)';
  if (typeof v === 'object') return v.error;
  if (typeof v === 'string') return JSON.stringify(v);
  return String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'FALSE' ? String(v).toUpperCase() : String(v);
}
