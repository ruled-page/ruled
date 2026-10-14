'use server';

import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';

/** Keeps the saved workbook beside the one opened; a cell changed to "refuse" is answered with an error instead. */
export async function save(saved) {
  if (saved.changes.some((change) => change.after === 'refuse')) return { error: 'The server refused this workbook.' };
  await writeFile(join(process.cwd(), 'workbooks', 'saved.xlsx'), saved.bytes);
}
