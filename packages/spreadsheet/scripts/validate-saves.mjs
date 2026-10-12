/**
 * Checks that saving never makes a workbook invalid. Edited copies of the
 * samples and of tests/corpus (from tests/roundtrip-out.test.ts, and from
 * workbook-verify with a row inserted) are validated with Microsoft's Open
 * XML SDK, the closest automated stand-in for Excel opening a file without
 * offering to repair it, and each is compared with its original: an error
 * the original didn't have fails the check (and any error, for a copy with
 * no original here). Some samples were generated with
 * libraries that write small schema errors of their own; those don't count.
 *
 * Needs the .NET SDK (10 or later) and `npm run build:verify`.
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const workspace = fileURLToPath(new URL('../../..', import.meta.url));
const out = mkdtempSync(join(tmpdir(), 'workbook-saves-'));
const sources = [join(workspace, 'examples', 'demo', 'public', 'samples'), join(root, 'tests', 'corpus')];
const originals = sources.flatMap((dir) => readdirSync(dir).filter((f) => /\.xls[xm]$/.test(f) && f !== 'big.xlsx' && f !== 'lookups.xlsx').map((f) => join(dir, f)));

// The safety tests add a workbook whose macros were removed.
execFileSync('npx', ['vitest', 'run', 'packages/spreadsheet/tests/roundtrip-out.test.ts', 'packages/spreadsheet/tests/safety.test.ts'], { cwd: workspace, stdio: 'inherit', env: { ...process.env, ROUNDTRIP_OUT: out } });
const verified = join(out, 'verify');
mkdirSync(verified);
spawnSync('node', [join(root, 'dist-verify', 'verify.js'), ...originals, '--structure', '--no-python', '--out', verified], { cwd: root, stdio: 'ignore' });
const saved = [
  ...readdirSync(out).filter((f) => f.endsWith('.xlsx')).map((f) => join(out, f)),
  ...readdirSync(verified).filter((f) => /\.xls[xm]$/.test(f)).map((f) => join(verified, f)),
];

execFileSync('dotnet', ['build', join(root, 'tools', 'ooxml-validate'), '-c', 'Release', '-v', 'q', '--nologo'], { stdio: 'inherit' });
const validator = join(root, 'tools', 'ooxml-validate', 'bin', 'Release', 'net10.0', 'OoxmlValidate.dll');
const validate = (files) => {
  const run = spawnSync('dotnet', [validator, ...files], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  return new Map(run.stdout.trim().split('\n').filter(Boolean).map((line) => {
    const r = JSON.parse(line);
    return [r.file, r.errors];
  }));
};
const key = (e) => `${e.part} ${e.id} ${e.description}`;
const before = validate(originals);
const after = validate(saved);

let failures = 0;
for (const file of saved) {
  const name = basename(file).split('.')[0];
  const original = originals.find((o) => basename(o).split('.')[0] === name);
  const known = new Set((original ? before.get(original) : []).map(key));
  const added = after.get(file).filter((e) => !known.has(key(e)));
  if (added.length) {
    failures++;
    console.log(`✗ ${basename(file)}`);
    for (const e of added.slice(0, 10)) console.log(`    ${e.part} ${e.path ?? ''}\n      ${e.description}`);
  } else {
    console.log(`✓ ${basename(file)}`);
  }
}
rmSync(out, { recursive: true, force: true });
console.log(failures ? `${failures} saved workbook(s) have errors their originals didn't.` : `All ${saved.length} saved workbooks are as valid as their originals.`);
process.exit(failures ? 1 : 0);
