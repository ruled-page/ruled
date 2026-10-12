import { expect, test } from '@playwright/test';
import { serverUrl, setCell } from './helpers';

/** The <workbook-editor> element in a plain HTML page, saving to the host through its save event. */
test('hands the host the workbook and its values on save', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();

  await setCell(page, 'B2', '125000');
  await expect(page.locator('#status')).toHaveText('1 change, not saved');

  // The sample has a negative cost in E3, which the host's schema doesn't allow.
  const save = page.getByRole('button', { name: 'Save from the host' });
  await save.click();
  await expect(page.getByText('Not saved. 1 cell breaks the “Forecast (from the host page)” schema.')).toBeVisible();
  await expect(page.locator('#received')).toHaveText('What the host receives on save appears here.');

  await setCell(page, 'E3', '30000');
  await save.click();
  await expect(page.locator('#status')).toHaveText('2 changes, saved');

  const received = JSON.parse((await page.locator('#received').textContent()) ?? '{}');
  expect(received.fileName).toBe('forecast.xlsx');
  expect(received.bytes).toBeGreaterThan(1000);
  expect(received.schema).toMatchObject({ id: 'example.forecast', version: '1', errors: 0 });
  expect(received.values.Forecast.tables.quarters[0]).toMatchObject({ revenue: 125000 });
  expect(received.values.Forecast.tables.quarters[3]).toMatchObject({ cost: 30000 });
  expect(received.changes.map((c: { ref: string }) => c.ref)).toEqual(['B2', 'E3']);
});

test('the element says when a workbook opens or can’t be loaded, and keeps it when moved', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await setCell(page, 'B2', '125000');

  const edits = await page.evaluate(async () => {
    const editor = document.getElementById('editor') as HTMLElement & { editedCount: number };
    const heard: string[] = ((window as unknown as { heard: string[] }).heard = []);
    editor.addEventListener('opened', (e) => heard.push(`opened ${(e as CustomEvent<{ fileName: string }>).detail.fileName}`));
    editor.addEventListener('editorerror', (e) => heard.push(`error ${(e as CustomEvent<{ message: string }>).detail.message}`));
    // Moved to another place in the page: the edit is still there.
    const holder = document.createElement('div');
    document.body.append(holder);
    holder.append(editor);
    await new Promise((r) => setTimeout(r, 50));
    return editor.editedCount;
  });
  expect(edits).toBe(1);

  // Another workbook, while this one has an unsaved edit: the user is asked first.
  await page.evaluate(() => document.getElementById('editor')!.setAttribute('src', '/samples/orders.xlsx'));
  await page.getByRole('button', { name: 'Discard and open' }).click();
  await expect.poll(() => page.evaluate(() => (window as unknown as { heard: string[] }).heard)).toEqual(['opened orders.xlsx']);

  // One that isn't there.
  await page.evaluate(() => document.getElementById('editor')!.setAttribute('src', '/nothing/here.xlsx'));
  await expect
    .poll(() => page.evaluate(() => (window as unknown as { heard: string[] }).heard[1]))
    .toMatch(/^error (Could not load \/nothing\/here\.xlsx: 404|This file is not an Excel workbook)/);
});
