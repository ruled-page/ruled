import { english, errorText } from '../i18n';
import { spawnSync } from 'node:child_process';
import { strFromU8, unzipSync } from 'fflate';
import { MAX_ROWS, toA1 } from '../core/refs';
import { gridValue, WorkbookSession } from '../core/session';
import type { ScanResult } from '../core/sheetScanner';
import { CellFlag, VT, type SheetGrid } from '../core/types';

export type CheckStatus = 'pass' | 'fail' | 'skip' | 'info';

export interface Check {
  name: string;
  status: CheckStatus;
  detail: string;
}

export interface VerifyReport {
  file: string;
  ok: boolean;
  checks: Check[];
  /** The edits made to the copy, as sheet!cell → value typed. */
  probes: { sheet: string; ref: string; text: string }[];
  /** How the engine compared with Excel's saved values. */
  calc?: { formulas: number; live: number; unsupported: number; mismatched: number; samples: { where: string; formula: string; excel?: string; engine?: string; reason: string }[] };
  /** The edited copy (and the copy with a row inserted, if asked for). */
  outputs: { edited?: Uint8Array; restructured?: Uint8Array };
}

export interface VerifyOptions {
  /** Also insert a row at the top of the first worksheet and check the result. */
  structure?: boolean;
  /** A Python with openpyxl, to have an independent reader check the edits; found automatically when omitted. */
  python?: string | false;
}

interface Probe {
  sheet: number;
  row: number;
  col: number;
  text: string;
  expect: string | number;
}

function cellsOf(grid: SheetGrid): { row: number; idx: number }[] {
  const out: { row: number; idx: number }[] = [];
  for (let row = 0; row < grid.rowCount; row++) for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) out.push({ row, idx: i });
  return out;
}

/** A few cells per sheet to change: numbers go up by one, a text gets a suffix, and a new cell goes below the data. */
function chooseProbes(s: WorkbookSession): Probe[] {
  const probes: Probe[] = [];
  s.model.sheets.forEach((m, sheet) => {
    const grid = m.scan?.grid;
    if (!grid || m.info.kind !== 'worksheet' || m.info.protected) return;
    const free = (row: number, col: number) =>
      !m.info.blocked.some((b) => row >= b.r1 && row <= b.r2 && col >= b.c1 && col <= b.c2) &&
      !m.info.merges.some((g) => row >= g.r1 && row <= g.r2 && col >= g.c1 && col <= g.c2 && (row !== g.r1 || col !== g.c1));
    let numbers = 0;
    let text = 0;
    for (const { row, idx } of cellsOf(grid)) {
      const col = grid.cols[idx];
      if (grid.flags[idx] & (CellFlag.Formula | CellFlag.Spilled) || !free(row, col)) continue;
      if (grid.types[idx] === VT.Number && numbers < 2 && Number.isInteger(grid.nums[idx])) {
        probes.push({ sheet, row, col, text: String(grid.nums[idx] + 1), expect: grid.nums[idx] + 1 });
        numbers++;
      } else if (grid.types[idx] === VT.String && text < 1 && row > 0 && (grid.strs[idx] ?? '').length < 60) {
        const value = `${grid.strs[idx]} (edited)`;
        probes.push({ sheet, row, col, text: `'${value}`, expect: value });
        text++;
      }
      if (numbers >= 2 && text >= 1) break;
    }
    const below = Math.max(m.info.maxRow, 0) + 2;
    if (below < MAX_ROWS && free(below, 0)) probes.push({ sheet, row: below, col: 0, text: 'verify', expect: 'verify' });
  });
  return probes;
}

/** The cells holding an array formula's values other than its own cell, dynamic or entered with Ctrl+Shift+Enter, as "row:col": their values follow its formula. */
function spillCells(scan: ScanResult | undefined): Set<string> {
  const out = new Set<string>();
  for (const a of scan?.arrayRanges ?? []) {
    if (scan!.fmeta.get(a.master)?.kind !== 'array') continue;
    for (let row = a.r1; row <= a.r2; row++) for (let col = a.c1; col <= a.c2; col++) if (row !== a.r1 || col !== a.c1) out.add(`${row}:${col}`);
  }
  return out;
}

function contentKey(grid: SheetGrid, idx: number): string {
  const f = grid.formulas.get(idx);
  if (f !== undefined && grid.flags[idx] & CellFlag.Formula) return `=${f}`;
  const v = gridValue(grid, idx);
  return v !== null && typeof v === 'object' ? v.error : String(v);
}

/** Finds a Python that can import openpyxl, for an independent reading of the saved file. */
function findPython(given?: string | false): string | undefined {
  if (given === false) return undefined;
  for (const candidate of given ? [given] : ['.venv/bin/python', 'python3', 'python']) {
    const r = spawnSync(candidate, ['-c', 'import openpyxl'], { encoding: 'utf8' });
    if (r.status === 0) return candidate;
  }
  return undefined;
}

const OPENPYXL_READ = `
import datetime, io, json, sys, openpyxl
from openpyxl.utils.datetime import to_excel
wb = openpyxl.load_workbook(io.BytesIO(sys.stdin.buffer.read()), data_only=False)
cells = json.loads(sys.argv[1])
out = []
for sheet, ref in cells:
    v = wb[sheet][ref].value
    if isinstance(v, (datetime.datetime, datetime.date)):
        v = to_excel(v)
    out.append(v if isinstance(v, (int, float, str, bool)) or v is None else str(v))
print(json.dumps(out))
`;

/**
 * Round-trips a workbook through the editor and checks nothing was lost.
 * Only reads; the caller decides where the edited copies go.
 */
export function verifyWorkbook(bytes: Uint8Array, file: string, options: VerifyOptions = {}): VerifyReport {
  const checks: Check[] = [];
  const add = (name: string, status: CheckStatus, detail: string) => checks.push({ name, status, detail });
  const report: VerifyReport = { file, ok: false, checks, probes: [], outputs: {} };

  let original: WorkbookSession;
  try {
    original = WorkbookSession.open(bytes, file, 'en-US', { recalc: 'on' });
  } catch (e) {
    add('Opens', 'fail', english.text(errorText(e)));
    return report;
  }
  add('Opens', 'pass', `${original.model.sheets.length} sheet${original.model.sheets.length === 1 ? '' : 's'}.`);

  const unchanged = original.save();
  const same = unchanged.length === bytes.length && unchanged.every((b, i) => b === bytes[i]);
  add('Saving without edits gives back the same file', same ? 'pass' : 'fail', same ? 'Byte for byte.' : 'The bytes differ.');

  // How the engine compares with Excel on this file's own formulas.
  const calc = original.ensureEngine();
  report.calc = {
    formulas: calc.formulaCells,
    live: calc.liveCells,
    unsupported: calc.unsupported,
    mismatched: calc.mismatched,
    samples: calc.samples.slice(0, 20).map((x) => ({ where: `${x.sheet}!${x.ref}`, formula: x.formula, excel: x.excel === undefined ? undefined : english.text(x.excel), engine: x.engine === undefined ? undefined : english.text(x.engine), reason: english.text(x.reason) })),
  };
  const exact = calc.formulaCells - calc.unsupported - calc.mismatched;
  add(
    'Formulas match Excel',
    calc.mismatched ? 'info' : 'pass',
    calc.formulaCells
      ? `${exact.toLocaleString()} of ${calc.formulaCells.toLocaleString()} formulas give Excel's saved value; ${calc.unsupported.toLocaleString()} use features not calculated here; ${calc.mismatched.toLocaleString()} differ (Excel's value is kept for those).`
      : 'No formulas.',
  );

  // Edit a copy and save it.
  const probes = chooseProbes(original);
  const session = WorkbookSession.open(bytes, file, 'en-US', { recalc: 'on' });
  session.ensureEngine();
  for (const p of probes) {
    const r = session.apply({ type: 'input', sheet: p.sheet, row: p.row, col: p.col, text: p.text });
    if (r.rejected) add(`Edit ${toA1(p.row, p.col)}`, 'fail', english.text(r.rejected));
  }
  report.probes = probes.map((p) => ({ sheet: session.model.sheets[p.sheet].info.name, ref: toA1(p.row, p.col), text: p.text }));
  let edited: Uint8Array;
  try {
    edited = session.save();
    report.outputs.edited = edited;
  } catch (e) {
    add('Saves after editing', 'fail', english.text(errorText(e)));
    return finish(report);
  }

  let reopened: WorkbookSession;
  try {
    reopened = WorkbookSession.open(edited, file);
    add('The edited copy opens again', 'pass', `${probes.length} cell${probes.length === 1 ? '' : 's'} edited.`);
  } catch (e) {
    add('The edited copy opens again', 'fail', english.text(errorText(e)));
    return finish(report);
  }

  // The edits read back.
  const wrong: string[] = [];
  for (const p of probes) {
    const grid = reopened.model.sheets[p.sheet].scan!.grid;
    const idx = findIndex(grid, p.row, p.col);
    const got = idx < 0 ? null : gridValue(grid, idx);
    if (got !== p.expect) wrong.push(`${reopened.model.sheets[p.sheet].info.name}!${toA1(p.row, p.col)} is ${JSON.stringify(got)}, expected ${JSON.stringify(p.expect)}`);
  }
  add('Edits read back', wrong.length ? 'fail' : 'pass', wrong.length ? wrong.slice(0, 5).join('; ') : 'Every edited cell holds what was typed.');

  // Everything else is as it was: constants equal, formulas unchanged.
  const probed = new Set(probes.map((p) => `${p.sheet}:${p.row}:${p.col}`));
  const diffs: string[] = [];
  let differing = 0;
  let compared = 0;
  original.model.sheets.forEach((m, s) => {
    const a = m.scan?.grid;
    const b = reopened.model.sheets[s]?.scan?.grid;
    if (!a || !b) return;
    // An array formula's values may change after an edit, as its saved value may.
    const spilled = new Set([...spillCells(m.scan), ...spillCells(reopened.model.sheets[s]?.scan)]);
    for (const { row, idx } of cellsOf(a)) {
      const col = a.cols[idx];
      if (probed.has(`${s}:${row}:${col}`) || spilled.has(`${row}:${col}`)) continue;
      compared++;
      const j = findIndex(b, row, col);
      const before = contentKey(a, idx);
      const formula = before.startsWith('=');
      const after = j < 0 ? '' : contentKey(b, j);
      // A formula's saved value may change after an edit; its formula may not.
      if (formula ? after !== before : after !== before && !(b.flags[j] & CellFlag.Formula)) {
        differing++;
        if (diffs.length < 10) diffs.push(`${m.info.name}!${toA1(row, col)}: ${before} → ${after || '(empty)'}`);
      }
    }
  });
  add('Every other cell is unchanged', differing ? 'fail' : 'pass', differing ? `${differing.toLocaleString()} differ: ${diffs.join('; ')}` : `${compared.toLocaleString()} cells compared.`);

  // Every part the editor doesn't write is byte for byte the same.
  const before = unzipSync(bytes);
  const after = unzipSync(edited);
  const touched = (path: string) =>
    /^xl\/worksheets\/sheet\d*\.xml$/.test(path) ||
    /^xl\/workbook\.xml$|^\[Content_Types\]\.xml$|^xl\/_rels\/workbook\.xml\.rels$|^xl\/styles\.xml$|^docProps\/custom\.xml$|^_rels\/\.rels$/.test(path) ||
    /pivotCacheDefinition\d*\.xml$/.test(path);
  // Folder entries in the zip aren't parts.
  const missing = Object.keys(before).filter((p) => !p.endsWith('/') && !(p in after) && !p.endsWith('calcChain.xml'));
  const changedParts = Object.keys(before).filter((p) => p in after && !touched(p) && !equalBytes(before[p], after[p]));
  add(
    'Charts, pivots, pictures and other parts are kept',
    missing.length || changedParts.length ? 'fail' : 'pass',
    missing.length || changedParts.length
      ? [missing.length ? `missing: ${missing.join(', ')}` : '', changedParts.length ? `changed: ${changedParts.join(', ')}` : ''].filter(Boolean).join('; ')
      : `${Object.keys(after).length} parts; only sheets with edits and workbook settings were rewritten${'xl/calcChain.xml' in before ? ', and the calculation chain removed so Excel rebuilds it' : ''}.`,
  );
  const sheetXml = Object.keys(after).filter((p) => /^xl\/worksheets\/sheet\d*\.xml$/.test(p));
  const badXml = sheetXml.filter((p) => !/<\/(\w+:)?worksheet>\s*$/.test(strFromU8(after[p])));
  if (badXml.length) add('Sheet XML is complete', 'fail', badXml.join(', '));

  // An independent reader.
  const python = findPython(options.python);
  if (!python) {
    add('An independent reader (openpyxl) reads the edits', 'skip', 'No Python with openpyxl found; pass --python to use one.');
  } else {
    const cells = probes.map((p) => [reopened.model.sheets[p.sheet].info.name, toA1(p.row, p.col)]);
    const r = spawnSync(python, ['-c', OPENPYXL_READ, JSON.stringify(cells)], { input: Buffer.from(edited), encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    if (r.status !== 0) {
      add('An independent reader (openpyxl) reads the edits', 'fail', (r.stderr || 'openpyxl could not read the file').trim().split('\n').slice(-1)[0]);
    } else {
      const values = JSON.parse(r.stdout) as unknown[];
      const off = probes.flatMap((p, i) => (values[i] === p.expect ? [] : [`${cells[i].join('!')} read as ${JSON.stringify(values[i])}`]));
      add('An independent reader (openpyxl) reads the edits', off.length ? 'fail' : 'pass', off.length ? off.join('; ') : `openpyxl reads all ${probes.length} edits.`);
    }
  }

  if (options.structure) {
    const first = original.model.sheets.findIndex((m) => m.info.kind === 'worksheet' && !m.info.protected);
    if (first < 0) {
      add('Inserting a row keeps every reference working', 'skip', 'No editable worksheet.');
    } else {
      try {
        const next = WorkbookSession.open(bytes, file).restructure({ type: 'insertRows', sheet: first, at: 0, count: 1 });
        const out = next.save();
        report.outputs.restructured = out;
        const again = WorkbookSession.open(out, file);
        const count = (sess: WorkbookSession) => sess.model.sheets.reduce((n, m) => n + (m.scan?.grid.formulas.size ?? 0), 0);
        const refs = (sess: WorkbookSession) => sess.model.sheets.reduce((n, m) => n + [...(m.scan?.grid.formulas.values() ?? [])].filter((f) => f.includes('#REF!')).length, 0);
        const ok = count(again) === count(original) && refs(again) === refs(original);
        add(
          'Inserting a row keeps every reference working',
          ok ? 'pass' : 'fail',
          ok
            ? `Row inserted at the top of ${original.model.sheets[first].info.name}; ${count(again).toLocaleString()} formulas kept, no new #REF!.`
            : `formulas ${count(original)} → ${count(again)}, #REF! ${refs(original)} → ${refs(again)}`,
        );
      } catch (e) {
        add('Inserting a row keeps every reference working', 'info', `Refused: ${english.text(errorText(e))}`);
      }
    }
  }
  return finish(report);
}

function finish(report: VerifyReport): VerifyReport {
  report.ok = report.checks.every((c) => c.status !== 'fail');
  return report;
}

function findIndex(grid: SheetGrid, row: number, col: number): number {
  if (row >= grid.rowCount) return -1;
  for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) if (grid.cols[i] === col) return i;
  return -1;
}

function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

const MARK: Record<CheckStatus, string> = { pass: '✓', fail: '✗', skip: '–', info: 'ℹ' };

/** The reports as Markdown, ending with what to check by hand in Excel. */
export function reportMarkdown(reports: VerifyReport[], outDir?: string): string {
  const lines = ['# Workbook round-trip report', ''];
  const failed = reports.filter((r) => !r.ok).length;
  lines.push(`${reports.length} workbook${reports.length === 1 ? '' : 's'} checked; ${failed ? `${failed} with failures` : 'no failures'}.`, '');
  for (const r of reports) {
    lines.push(`## ${r.file} ${r.ok ? '✓' : '✗'}`, '');
    for (const c of r.checks) lines.push(`- ${MARK[c.status]} **${c.name}**: ${c.detail}`);
    if (r.calc?.samples.length) {
      lines.push('', 'Formulas not calculated here or calculated differently (first 20):', '', '| Cell | Formula | Excel | Here | Why |', '| --- | --- | --- | --- | --- |');
      // A | inside a cell, in a formula or in a text value, would end the cell.
      const cell = (text = '') => text.replace(/\|/g, '\\|');
      for (const x of r.calc.samples) lines.push(`| ${cell(x.where)} | \`${cell(x.formula)}\` | ${cell(x.excel)} | ${cell(x.engine)} | ${cell(x.reason)} |`);
    }
    if (r.probes.length) lines.push('', `Edited in the copy: ${r.probes.map((p) => `${p.sheet}!${p.ref} = ${p.text.replace(/^'/, '')}`).join(', ')}.`);
    lines.push('');
  }
  lines.push(
    '## The step that needs Excel',
    '',
    `Open each edited copy${outDir ? ` in \`${outDir}\`` : ''} in desktop Excel and check:`,
    '',
    '1. It opens without "We found a problem with some content" (a repair prompt).',
    '2. The edited cells listed above show their new values, and formulas that use them have recalculated.',
    '3. Charts, pivot tables (they refresh on open), pictures, comments, validation and conditional formatting look as in the original.',
    '4. Save it from Excel and close: Excel shouldn\'t ask to save again on reopening.',
    '',
  );
  return lines.join('\n');
}
