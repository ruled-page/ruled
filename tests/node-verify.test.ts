import { spawnSync } from 'node:child_process';
import { chmodSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { afterAll, afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { msg, TextError } from '../src/i18n';
import { reportMarkdown, verifyWorkbook, type VerifyReport } from '../src/node/verify';
import { runVerify, VERIFY_USAGE, type VerifyIO } from '../src/node/verifyCli';
import { sample } from './helpers';

async function workbook(build: (wb: ExcelJS.Workbook) => void | Promise<void>): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  await build(wb);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const check = (r: VerifyReport, name: string) => r.checks.find((c) => c.name === name);
const READER = 'An independent reader (openpyxl) reads the edits';
const INSERT = 'Inserting a row keeps every reference working';

// Stand-ins for Python: they claim to have openpyxl, then answer the read as told.
const bin = mkdtempSync(join(tmpdir(), 'verify-python-'));
afterAll(() => rmSync(bin, { recursive: true, force: true }));
function fakePython(name: string, answer: string): string {
  const path = join(bin, name);
  writeFileSync(path, `#!/bin/sh\nif [ "$2" = "import openpyxl" ]; then exit 0; fi\ncat > /dev/null\n${answer}\n`);
  chmodSync(path, 0o755);
  return path;
}
const hasOpenpyxl = ['.venv/bin/python', 'python3', 'python'].some((p) => spawnSync(p, ['-c', 'import openpyxl']).status === 0);

describe('verifying a workbook', () => {
  it('checks a one-sheet workbook without formulas, editing the cell below its data', async () => {
    const r = verifyWorkbook(await workbook((wb) => void (wb.addWorksheet('Data').getCell('A1').value = 'Header')), 'one.xlsx', { python: false });
    expect(r.ok).toBe(true);
    expect(r.probes).toEqual([{ sheet: 'Data', ref: 'A3', text: 'verify' }]);
    expect(check(r, 'Opens')?.detail).toBe('1 sheet.');
    expect(check(r, 'Formulas match Excel')).toEqual({ name: 'Formulas match Excel', status: 'pass', detail: 'No formulas.' });
    expect(check(r, 'The edited copy opens again')?.detail).toBe('1 cell edited.');
    expect(check(r, 'Edits read back')?.status).toBe('pass');
    expect(check(r, INSERT)).toBeUndefined();
  });

  it('edits numbers and text, but nothing on a protected sheet', async () => {
    const bytes = await workbook(async (wb) => {
      const open = wb.addWorksheet('Open');
      open.getCell('A1').value = 'Name';
      open.getCell('A2').value = 'Widget';
      open.getCell('B2').value = 41;
      const locked = wb.addWorksheet('Locked');
      locked.getCell('A1').value = 1;
      await locked.protect('', {});
    });
    const r = verifyWorkbook(bytes, 'mixed.xlsx', { python: false });
    expect(r.ok).toBe(true);
    expect(r.probes).toEqual([
      { sheet: 'Open', ref: 'A2', text: "'Widget (edited)" },
      { sheet: 'Open', ref: 'B2', text: '42' },
      { sheet: 'Open', ref: 'A4', text: 'verify' },
    ]);
  });

  it('skips inserting a row when no worksheet can be edited', async () => {
    const bytes = await workbook(async (wb) => {
      const ws = wb.addWorksheet('Locked');
      ws.getCell('A1').value = 1;
      await ws.protect('', {});
    });
    const r = verifyWorkbook(bytes, 'locked.xlsx', { structure: true, python: false });
    expect(r.probes).toEqual([]);
    expect(check(r, INSERT)).toEqual({ name: INSERT, status: 'skip', detail: 'No editable worksheet.' });
    expect(r.outputs.restructured).toBeUndefined();
  });

  it('copes with data in the last row: no edit below it, and says in words why no row can go in', async () => {
    const bytes = await workbook((wb) => {
      const ws = wb.addWorksheet('Tall');
      ws.getCell('A1').value = 1;
      ws.getCell('A1048576').value = 'last';
    });
    const r = verifyWorkbook(bytes, 'tall.xlsx', { structure: true, python: false });
    expect(r.probes.map((p) => p.ref)).toEqual(['A1', 'A1048576']);
    expect(check(r, INSERT)).toEqual({ name: INSERT, status: 'info', detail: "Refused: There isn't room: inserting would push cells with data off the bottom of the sheet." });
    expect(r.ok).toBe(true);
  });

  it('notes formulas whose saved value differs from this engine’s, and lists them in the report', async () => {
    const bytes = await workbook((wb) => {
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = { formula: 'IF(A1>0,"x|y","z")', result: 'a|b' };
    });
    const r = verifyWorkbook(bytes, 'calc.xlsx', { python: false });
    expect(check(r, 'Formulas match Excel')).toEqual({
      name: 'Formulas match Excel',
      status: 'info',
      detail: "0 of 1 formulas give Excel's saved value; 0 use features not calculated here; 1 differ (Excel's value is kept for those).",
    });
    expect(r.ok).toBe(true);
    expect(r.calc).toMatchObject({ formulas: 1, mismatched: 1, samples: [{ where: 'Data!B1', formula: '=IF(A1>0,"x|y","z")', excel: '"a|b"', engine: '"x|y"' }] });
    const row = reportMarkdown([r]).split('\n').find((l) => l.startsWith('| Data!B1'))!;
    // Every | inside a cell is escaped, so the row keeps its five columns.
    expect(row.replace(/\\\|/g, '').split('|')).toHaveLength(7);
    expect(row).toMatch(/^\| Data!B1 \| `=IF\(A1>0,"x\\\|y","z"\)` \| "a\\\|b" \| "x\\\|y" \| /);
  });
});

describe('when the editor gets it wrong', () => {
  afterEach(() => vi.restoreAllMocks());
  const bytes = () =>
    workbook((wb) => {
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 'Name';
      ws.getCell('A2').value = 1;
    });

  it('catches edits that land in the wrong cell', async () => {
    const apply = WorkbookSession.prototype.apply;
    vi.spyOn(WorkbookSession.prototype, 'apply').mockImplementation(function (this: WorkbookSession, op) {
      return apply.call(this, op.type === 'input' ? { ...op, row: op.row - 1 } : op);
    });
    const r = verifyWorkbook(await bytes(), 'x.xlsx', { python: false });
    expect(r.ok).toBe(false);
    expect(check(r, 'Edits read back')).toMatchObject({ status: 'fail', detail: 'Data!A2 is 1, expected 2; Data!A4 is null, expected "verify"' });
    expect(check(r, 'Every other cell is unchanged')).toMatchObject({ status: 'fail', detail: '1 differ: Data!A1: Name → 2' });
  });

  it('stops at a copy that won’t save, or won’t open again, saying why in English', async () => {
    const original = await bytes();
    const save = WorkbookSession.prototype.save;
    const saving = vi.spyOn(WorkbookSession.prototype, 'save').mockImplementation(function (this: WorkbookSession, ...args) {
      if (this.editedCount) throw new Error('Disk full.');
      return save.apply(this, args);
    });
    const unsaved = verifyWorkbook(original, 'x.xlsx', { python: false });
    expect(unsaved.checks.at(-1)).toEqual({ name: 'Saves after editing', status: 'fail', detail: 'Disk full.' });
    expect(unsaved.ok).toBe(false);
    saving.mockRestore();

    const open = WorkbookSession.open;
    vi.spyOn(WorkbookSession, 'open').mockImplementation((data, ...rest) => {
      if (data !== original) throw new TextError(msg('core.open.notPackage'));
      return open(data, ...rest);
    });
    const unopened = verifyWorkbook(original, 'x.xlsx', { python: false });
    expect(unopened.checks.at(-1)).toEqual({ name: 'The edited copy opens again', status: 'fail', detail: 'This file is not a valid Office Open XML package.' });
    expect(unopened.outputs.edited).toBeDefined();
  });
});

describe('the independent reading', () => {
  const probed = () =>
    workbook((wb) => {
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 'text';
    });

  it('is skipped without a Python that has openpyxl', async () => {
    const r = verifyWorkbook(await probed(), 'x.xlsx', { python: join(bin, 'no-such-python') });
    expect(check(r, READER)).toEqual({ name: READER, status: 'skip', detail: 'No Python with openpyxl found; pass --python to use one.' });
  });

  it('fails with the reader’s last words when it can’t read the copy', async () => {
    const python = fakePython('broken', 'echo "Traceback (most recent call last):" >&2\necho "zipfile.BadZipFile: File is not a zip file" >&2\nexit 1');
    const r = verifyWorkbook(await probed(), 'x.xlsx', { python });
    expect(check(r, READER)).toEqual({ name: READER, status: 'fail', detail: 'zipfile.BadZipFile: File is not a zip file' });
    expect(r.ok).toBe(false);
  });

  it('fails naming each cell the reader saw differently', async () => {
    const python = fakePython('wrong', `echo '[2, "nope", null]'`);
    const r = verifyWorkbook(await probed(), 'x.xlsx', { python });
    expect(r.probes.map((p) => p.ref)).toEqual(['A1', 'A2', 'A4']);
    expect(check(r, READER)).toEqual({ name: READER, status: 'fail', detail: 'Data!A2 read as "nope"; Data!A4 read as null' });
  });

  it.skipIf(!hasOpenpyxl)('passes when openpyxl, found on its own, reads every edit', async () => {
    const r = verifyWorkbook(await probed(), 'x.xlsx');
    expect(check(r, READER)).toEqual({ name: READER, status: 'pass', detail: 'openpyxl reads all 3 edits.' });
  });
});

describe('the report', () => {
  it('summarises several workbooks, failures first in the count, and says where the copies are', () => {
    const good = verifyWorkbook(sample('orders.xlsx'), 'orders.xlsx', { python: false });
    const bad = verifyWorkbook(new TextEncoder().encode('nope'), 'nope.xlsx', { python: false });
    const md = reportMarkdown([good, bad], 'out');
    expect(md).toMatch(/^# Workbook round-trip report\n\n2 workbooks checked; 1 with failures\.\n/);
    expect(md).toContain('## orders.xlsx ✓');
    expect(md).toContain('## nope.xlsx ✗\n\n- ✗ **Opens**: This file is not an Excel workbook (.xlsx or .xlsm).');
    expect(md).toContain('- – **An independent reader (openpyxl) reads the edits**');
    expect(md).toContain('Open each edited copy in `out` in desktop Excel');
    const one = reportMarkdown([good]);
    expect(one).toContain('1 workbook checked; no failures.');
    expect(one).toContain('Open each edited copy in desktop Excel');
    expect(one).toMatch(/Edited in the copy: Orders!\w+\d+ = /);
  });
});

describe('workbook-verify', () => {
  /** Runs the command line over files held in memory. */
  async function cli(argv: string[], files: Record<string, Uint8Array> = {}) {
    const out: string[] = [];
    const err: string[] = [];
    const written: Record<string, Uint8Array | string> = {};
    const io: VerifyIO = {
      list: async (p) => {
        const found = Object.keys(files).filter((f) => f === p || f.startsWith(`${p}/`));
        if (!found.length) throw new Error(`ENOENT: no such file or directory, stat '${p}'`);
        return found;
      },
      readFile: async (p) => files[p],
      writeFile: async (p, d) => void (written[p] = d),
      out: (t) => out.push(t),
      err: (t) => err.push(t),
    };
    return { code: await runVerify(argv, io), out: out.join('\n'), err: err.join('\n'), written };
  }

  it('shows the usage for --help and -h', async () => {
    expect(await cli(['--help'])).toEqual({ code: 0, out: VERIFY_USAGE, err: '', written: {} });
    expect((await cli(['in', '-h'])).code).toBe(0);
  });

  it('turns down a command line it can’t use', async () => {
    expect(await cli(['in', '--fast'])).toMatchObject({ code: 2, err: `Unknown option --fast.\n\n${VERIFY_USAGE}` });
    expect(await cli([])).toMatchObject({ code: 2, err: VERIFY_USAGE });
    expect(await cli(['in', '--out'])).toMatchObject({ code: 2, err: VERIFY_USAGE });
    expect(await cli(['missing'])).toMatchObject({ code: 2, err: "ENOENT: no such file or directory, stat 'missing'" });
  });

  it('prints the reports as JSON, without the copies, and exits 1 when one fails', async () => {
    const r = await cli(['in', '--json', '--no-python'], { 'in/orders.xlsx': sample('orders.xlsx'), 'in/nope.xlsx': new TextEncoder().encode('nope') });
    expect(r.code).toBe(1);
    expect(r.err).toBe('');
    const reports = JSON.parse(r.out);
    expect(reports.map((x: VerifyReport) => [x.file, x.ok])).toEqual([
      ['orders.xlsx', true],
      ['nope.xlsx', false],
    ]);
    expect('outputs' in reports[0]).toBe(false);
  });

  it('writes no copy of a file it couldn’t open, and marks it failed', async () => {
    const r = await cli(['in', '--out', 'out', '--no-python'], { 'in/nope.xlsx': new TextEncoder().encode('nope') });
    expect(r.code).toBe(1);
    expect(r.err).toBe('✗ nope.xlsx');
    expect(Object.keys(r.written).sort()).toEqual(['out/report.json', 'out/report.md']);
  });

  it('writes the edited and row-inserted copies with the workbook’s own extension', async () => {
    const r = await cli(['in/orders.xlsm', '--out', 'out', '--structure', '--python', join(bin, 'no-such-python')], { 'in/orders.xlsm': sample('orders.xlsx') });
    expect(r.code).toBe(0);
    expect(r.err).toBe('✓ orders.xlsm');
    expect(Object.keys(r.written).sort()).toEqual(['out/orders.edited.xlsm', 'out/orders.row-inserted.xlsm', 'out/report.json', 'out/report.md']);
    expect(r.out).toBe(r.written['out/report.md']);
    expect(String(r.written['out/report.md'])).toContain('No Python with openpyxl found');
    expect(JSON.parse(String(r.written['out/report.json']))[0].checks.find((c: { name: string }) => c.name === INSERT).status).toBe('pass');
  });
});
