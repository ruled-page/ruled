import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { SAMPLES } from './helpers';

/**
 * The command-line entry points run on import: they read process.argv, write
 * to stdout and stderr, and set the exit code.
 */
let dir: string;
let argv: string[];
let out: string;
let err: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'workbook-bin-'));
  argv = process.argv;
  out = '';
  err = '';
  vi.spyOn(process.stdout, 'write').mockImplementation((chunk) => ((out += String(chunk)), true));
  vi.spyOn(process.stderr, 'write').mockImplementation((chunk) => ((err += String(chunk)), true));
});
afterEach(() => {
  vi.restoreAllMocks();
  process.argv = argv;
  process.exitCode = undefined;
  rmSync(dir, { recursive: true, force: true });
});

/** Runs an entry point with these arguments; resolves to its exit code. */
async function main(args: string[]): Promise<number> {
  process.argv = ['node', 'verify', ...args];
  process.exitCode = undefined;
  vi.resetModules();
  await import('../src/verify/bin');
  await vi.waitFor(() => expect(process.exitCode).toBeDefined(), { timeout: 10_000 });
  return Number(process.exitCode);
}

describe('workbook-verify on the command line', () => {
  it('checks the workbooks in a folder, passing over lock files and other files, and writes the report', async () => {
    const input = join(dir, 'in');
    mkdirSync(input);
    copyFileSync(join(SAMPLES, 'orders.xlsx'), join(input, 'orders.xlsx'));
    writeFileSync(join(input, '~$orders.xlsx'), 'Excel’s lock file');
    writeFileSync(join(input, 'notes.txt'), 'not a workbook');
    const report = join(dir, 'out', 'nested');
    expect(await main([input, '--out', report, '--no-python'])).toBe(0);
    expect(err).toBe('✓ orders.xlsx\n');
    expect(out).toContain('1 workbook checked; no failures.');
    expect(JSON.parse(readFileSync(join(report, 'report.json'), 'utf8')).map((r: { file: string }) => r.file)).toEqual(['orders.xlsx']);
    expect(existsSync(join(report, 'orders.edited.xlsx'))).toBe(true);
  });

  it('takes a single file, and exits 2 for a path that isn’t there', async () => {
    expect(await main([join(SAMPLES, 'orders.xlsx'), '--json', '--no-python'])).toBe(0);
    expect(JSON.parse(out)[0].file).toBe('orders.xlsx');
    out = '';
    expect(await main([join(dir, 'nothing-here')])).toBe(2);
    expect(err).toMatch(/ENOENT/);
  });

  it('exits 2 with the reason when the report can’t be written', async () => {
    const blocked = join(dir, 'a-file');
    writeFileSync(blocked, '');
    expect(await main([join(SAMPLES, 'orders.xlsx'), '--out', join(blocked, 'out'), '--no-python'])).toBe(2);
    expect(err).toMatch(/ENOTDIR|EEXIST/);
  });
});
