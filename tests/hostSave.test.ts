import { describe, expect, it } from 'vitest';
import { schemaFromJSON } from '../src/schema/json';
import type { SavedWorkbook } from '../src/state/store';
import { forecastSchema } from '../src/demo/schemas';
import { sample } from './helpers';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

/** The forecast sample breaks its schema on purpose, so saving asks first; this says yes. */
async function saveAnyway(store: WorkbookStore): Promise<void> {
  const saving = store.download();
  for (let i = 0; i < 100 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  store.confirmRequest?.resolve('confirm');
  await saving;
}

const versioned = (version: string) => ({ ...schemaFromJSON(JSON.stringify(forecastSchema)), version, saveWithErrors: 'confirm' as const });

describe('saving to the host', () => {
  it('hands over the file, the schema values and readers for any cell', async () => {
    let got: SavedWorkbook | undefined;
    const schema = versioned('3');
    const store = await testStore({ schemas: [schema], onSave: (f) => void (got = f), saveLabel: 'Save to ITAS' });
    await opened(store, sample('forecast.xlsx'), 'forecast.xlsx');
    expect(store.schema?.schema.id).toBe(schema.id);
    expect(store.saveLabel).toBe('Save to ITAS');

    await store.input(1, 1, '1500');
    expect(store.unsaved).toBe(true);
    await saveAnyway(store);
    expect(got).toBeDefined();
    expect(got!.schema).toMatchObject({ id: schema.id, version: '3', errors: 2 });
    const quarters = Object.values(got!.values!)[0].tables;
    const first = Object.values(quarters)[0][0];
    expect(Object.values(first)).toContain(1500);
    expect(got!.cell('Forecast', 'B2')).toBe(1500);
    expect(got!.range('Forecast', 'A1:B2')[1][1]).toBe(1500);
    expect(got!.changes.map((c) => c.ref)).toContain('B2');
    expect(store.unsaved).toBe(false);
    store.dispose();
  });

  it('stamps the file with the schema version, so reopening finds it', async () => {
    let bytes: Uint8Array | undefined;
    const schema = versioned('3');
    const store = await testStore({ schemas: [schema], onSave: (f) => void (bytes = f.bytes) });
    await opened(store, sample('forecast.xlsx'), 'forecast.xlsx');
    await store.input(1, 1, '1500');
    await saveAnyway(store);
    store.dispose();

    const newer = versioned('4');
    const again = await testStore({ schemas: [newer] });
    await opened(again, bytes!, 'forecast.xlsx');
    expect(again.properties).toEqual({ WorkbookSchema: `${schema.id}@3` });
    expect(again.schemaNote).toMatch(/last saved with version 3 .* checked against version 4/);
    again.dispose();
  });
});
