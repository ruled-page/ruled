import { reportMarkdown, verifyWorkbook, type VerifyReport } from './verify';

export interface VerifyIO {
  /** The workbook files under a path: the file itself, or the .xlsx and .xlsm files in a folder. */
  list(path: string): Promise<string[]>;
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array | string): Promise<void>;
  out(text: string): void;
  err(text: string): void;
}

export const VERIFY_USAGE = `Round-trips workbooks through the editor and checks nothing is lost.

Usage: workbook-verify <file.xlsx | folder>... [options]

  --out <folder>     Write the edited copies and report.md / report.json here
  --structure        Also insert a row at the top of each workbook's first sheet
  --python <path>    A Python with openpyxl for an independent reading (found automatically)
  --no-python        Skip the independent reading
  --json             Print the report as JSON instead of Markdown

Exit status: 0 when every check passes, 1 when any fails, 2 for usage errors.`;

export async function runVerify(argv: string[], io: VerifyIO): Promise<number> {
  const paths: string[] = [];
  let out: string | undefined;
  let structure = false;
  let json = false;
  let python: string | false | undefined;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') out = argv[++i];
    else if (a === '--structure') structure = true;
    else if (a === '--json') json = true;
    else if (a === '--python') python = argv[++i];
    else if (a === '--no-python') python = false;
    else if (a === '--help' || a === '-h') {
      io.out(VERIFY_USAGE);
      return 0;
    } else if (a.startsWith('--')) {
      io.err(`Unknown option ${a}.\n\n${VERIFY_USAGE}`);
      return 2;
    } else paths.push(a);
  }
  if (!paths.length || (out === undefined && argv.includes('--out'))) {
    io.err(VERIFY_USAGE);
    return 2;
  }

  const reports: VerifyReport[] = [];
  for (const p of paths) {
    let files: string[];
    try {
      files = await io.list(p);
    } catch (e) {
      io.err((e as Error).message);
      return 2;
    }
    for (const file of files) {
      const name = file.split(/[\\/]/).pop()!;
      const report = verifyWorkbook(await io.readFile(file), name, { structure, python });
      reports.push(report);
      if (out) {
        const base = `${out}/${name.replace(/\.(xlsx|xlsm)$/i, '')}`;
        const ext = /\.xlsm$/i.test(name) ? '.xlsm' : '.xlsx';
        if (report.outputs.edited) await io.writeFile(`${base}.edited${ext}`, report.outputs.edited);
        if (report.outputs.restructured) await io.writeFile(`${base}.row-inserted${ext}`, report.outputs.restructured);
      }
      if (!json) io.err(`${report.ok ? '✓' : '✗'} ${name}`);
    }
  }
  const markdown = reportMarkdown(reports, out);
  const plain = reports.map(({ outputs: _outputs, ...r }) => r);
  if (out) {
    await io.writeFile(`${out}/report.md`, markdown);
    await io.writeFile(`${out}/report.json`, JSON.stringify(plain, null, 2));
  }
  io.out(json ? JSON.stringify(plain, null, 2) : markdown);
  return reports.every((r) => r.ok) ? 0 : 1;
}
