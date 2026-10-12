import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { runVerify } from './cli';

runVerify(process.argv.slice(2), {
  list: async (path) => {
    if (!(await stat(path)).isDirectory()) return [path];
    return (await readdir(path)).filter((f) => /\.(xlsx|xlsm)$/i.test(f) && !f.startsWith('~$')).sort().map((f) => join(path, f));
  },
  readFile: async (path) => new Uint8Array(await readFile(path)),
  writeFile: async (path, data) => {
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, data);
  },
  out: (text) => process.stdout.write(`${text}\n`),
  err: (text) => process.stderr.write(`${text}\n`),
}).then(
  (code) => {
    process.exitCode = code;
  },
  (e: Error) => {
    process.stderr.write(`${e.stack ?? e.message}\n`);
    process.exitCode = 2;
  },
);
