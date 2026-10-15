import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { FileSource, PluginContext, PluginText } from '@ruled/sheet/plugin';
import { HeadlessWorkbook } from '@ruled/sheet/testing/headless.ts';
import { blankWorkbook } from '../src/blankWorkbook';
import { newWorkbook } from '../src/newWorkbook';

describe('New workbook', () => {
  it('opens as one empty sheet, with no findings', () => {
    const wb = HeadlessWorkbook.open(blankWorkbook('Tabelle1'), { fileName: 'Mappe1.xlsx' });
    expect(wb.sheetNames).toEqual(['Tabelle1']);
    expect(wb.findings).toEqual([]);
  });

  it('differs each time it is made, so drafts of one never come back for another', () => {
    const created = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['docProps/core.xml']);
    expect(created(blankWorkbook('Sheet1', new Date(1)))).not.toBe(created(blankWorkbook('Sheet1', new Date(2))));
  });

  it('is offered in the Open menu, named in the editor’s language, and can’t be saved back', async () => {
    let source!: FileSource;
    const ctx = {
      text: (t: PluginText) => (t as { de: string }).de,
      files: { registerSource: (s: FileSource) => ((source = s), () => undefined) },
    } as unknown as PluginContext;
    newWorkbook().setup(ctx);
    expect(source.label).toEqual({ en: 'New workbook', de: 'Neue Arbeitsmappe' });
    const file = await source.pick!();
    expect(file!.origin).toEqual({ source: 'new-workbook', name: 'Mappe1.xlsx', ref: null, writable: false });
    expect(HeadlessWorkbook.open(file!.bytes as Uint8Array).sheetNames).toEqual(['Tabelle1']);
  });
});
