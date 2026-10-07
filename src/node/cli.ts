import { schemaFromJSON } from '../schema/json';
import type { WorkbookSchema } from '../schema/types';
import { ServerWorkbook, type WorkbookIssue } from './index';

export interface CliIO {
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array): Promise<void>;
  out(text: string): void;
  err(text: string): void;
}

export const USAGE = `Checks an Excel workbook against schemas, optionally changes cells, and saves it.

Usage: workbook-check <file.xlsx> [options]

  --schema <file.schema.json>   A schema to check against (repeat for several; the first that fits applies)
  --set <Sheet!A1=value>        Type a value into a cell, as a user would (repeat for several)
  --out <file.xlsx>             Save the workbook here (stamped with the schema's version)
  --values                      Print the values in the shape of the schema
  --json                        Print the result as JSON
  --help                        Show this

Exit status: 0 when the workbook matches a schema with no errors (or no schema
was given), 1 when it has errors or matches none, 2 when it can't be read.`;

interface Args {
  file?: string;
  schemas: string[];
  sets: { sheet: string; ref: string; value: string }[];
  out?: string;
  values: boolean;
  json: boolean;
  help: boolean;
}

function parseArgs(argv: string[]): Args {
  const args: Args = { schemas: [], sets: [], values: false, json: false, help: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => {
      const v = argv[++i];
      if (v === undefined) throw new Error(`${a} needs a value.`);
      return v;
    };
    if (a === '--schema') args.schemas.push(next());
    else if (a === '--set') {
      const m = /^(?:'([^']+)'|([^!]+))!(\$?[A-Za-z]{1,3}\$?\d+)=(.*)$/s.exec(next());
      if (!m) throw new Error('--set takes Sheet!A1=value.');
      args.sets.push({ sheet: m[1] ?? m[2], ref: m[3].replace(/\$/g, '').toUpperCase(), value: m[4] });
    } else if (a === '--out') args.out = next();
    else if (a === '--values') args.values = true;
    else if (a === '--json') args.json = true;
    else if (a === '--help' || a === '-h') args.help = true;
    else if (a.startsWith('--')) throw new Error(`Unknown option ${a}.`);
    else if (!args.file) args.file = a;
    else throw new Error(`Only one workbook at a time (got ${args.file} and ${a}).`);
  }
  return args;
}

const where = (i: WorkbookIssue) => (i.ref ? `${i.sheet}!${i.ref}` : i.sheet || 'workbook');

/** Runs the command line; resolves to the exit status. */
export async function run(argv: string[], io: CliIO): Promise<number> {
  let args: Args;
  try {
    args = parseArgs(argv);
  } catch (e) {
    io.err(`${(e as Error).message}\n\n${USAGE}`);
    return 2;
  }
  if (args.help || !args.file) {
    (args.help ? io.out : io.err)(USAGE);
    return args.help ? 0 : 2;
  }

  let wb: ServerWorkbook;
  try {
    const schemas: WorkbookSchema[] = [];
    for (const path of args.schemas) {
      const text = new TextDecoder().decode(await io.readFile(path));
      try {
        schemas.push(schemaFromJSON(text));
      } catch (e) {
        throw new Error(`${path}: ${(e as Error).message}`);
      }
    }
    wb = ServerWorkbook.open(await io.readFile(args.file), { fileName: args.file.split(/[\\/]/).pop(), schemas });
  } catch (e) {
    io.err((e as Error).message);
    return 2;
  }

  const rejected: string[] = [];
  for (const s of args.sets) {
    try {
      const r = await wb.set(s.sheet, s.ref, s.value);
      if (r.rejected) rejected.push(`${s.sheet}!${s.ref}: ${r.rejected}`);
      for (const w of r.warnings) rejected.push(`${s.sheet}!${s.ref}: ${w}`);
    } catch (e) {
      rejected.push(`${s.sheet}!${s.ref}: ${(e as Error).message}`);
    }
  }

  const issues = await wb.issues();
  const errors = issues.filter((i) => i.severity === 'error').length;
  const ok = !wb.mismatch && errors === 0 && rejected.length === 0;
  if (args.out) {
    const check = await wb.checkBeforeSave();
    if (!check.ok && !check.canOverride && wb.schema?.saveWithErrors !== 'confirm') {
      rejected.push(`Not saved: ${check.message}`);
    } else {
      await io.writeFile(args.out, wb.save());
    }
  }

  if (args.json) {
    io.out(
      JSON.stringify(
        {
          ok,
          file: wb.fileName,
          schema: wb.schema && { id: wb.schema.id, name: wb.schema.name, version: wb.schema.version },
          mismatch: wb.mismatch,
          note: wb.note,
          errors,
          warnings: issues.length - errors,
          issues,
          rejected,
          changes: wb.changes(),
          saved: args.out && !rejected.some((r) => r.startsWith('Not saved')) ? args.out : undefined,
          values: args.values ? wb.values() : undefined,
        },
        null,
        2,
      ),
    );
    return ok ? 0 : 1;
  }

  const lines: string[] = [];
  if (wb.schema) lines.push(`${wb.fileName}: matches “${wb.schema.name}”${wb.schema.version ? ` (version ${wb.schema.version})` : ''}.`);
  else if (wb.mismatch) {
    lines.push(`${wb.fileName}: doesn't match ${wb.mismatch.schema ? `“${wb.mismatch.schema}”` : 'any schema given'}.`);
    for (const p of wb.mismatch.problems) lines.push(`  ${p}`);
  } else lines.push(`${wb.fileName}: read ${wb.sheetNames.length} sheet${wb.sheetNames.length === 1 ? '' : 's'}; no schema given.`);
  if (wb.note) lines.push(wb.note);
  for (const i of issues) lines.push(`  ${i.severity === 'error' ? 'error  ' : 'warning'}  ${where(i)}  ${i.message}`);
  for (const r of rejected) lines.push(`  not done  ${r}`);
  if (wb.schema) lines.push(`${errors} error${errors === 1 ? '' : 's'}, ${issues.length - errors} warning${issues.length - errors === 1 ? '' : 's'}.`);
  if (args.out && !rejected.some((r) => r.startsWith('Not saved'))) lines.push(`Saved ${args.out} (${wb.editedCount} change${wb.editedCount === 1 ? '' : 's'}).`);
  if (args.values) lines.push(JSON.stringify(wb.values() ?? null, null, 2));
  io.out(lines.join('\n'));
  return ok ? 0 : 1;
}
