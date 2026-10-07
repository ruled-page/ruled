import { readFile, writeFile } from 'node:fs/promises';
import { run } from './cli';

run(process.argv.slice(2), {
  readFile: async (path) => new Uint8Array(await readFile(path)),
  writeFile: (path, data) => writeFile(path, data),
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
