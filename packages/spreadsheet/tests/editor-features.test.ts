import ExcelJS from 'exceljs';
import { createElement, type ReactElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { I18nextProvider } from 'react-i18next';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { NAMESPACE } from '../src/i18n';
import type { EditorFeatures, WorkbookStore } from '../src/state/store';
import type { MenuItem } from '../src/ui/ContextMenu';
import { controllerWorkbook, logSheet, workbookSchema } from './schema-fixture';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

// The sheet tab menu, caught before it would be shown: showing it needs a browser.
const shown = vi.hoisted(() => ({ items: undefined as MenuItem[] | undefined }));
vi.mock('../src/ui/ContextMenu', async (original) => ({
  ...(await original<typeof import('../src/ui/ContextMenu')>()),
  ContextMenu: ({ items }: { items: MenuItem[] }) => {
    shown.items = items;
    return null;
  },
}));
// Read through a function, which TypeScript doesn't narrow after `shown.items = undefined`.
const menuShown = () => shown.items;
const { SheetTabs, StatusBar, Toolbar } = await import('../src/ui/Chrome');
const { cellNote } = await import('../src/ui/FormulaBar');

let sheets: Uint8Array;
let schemaBook: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = 1;
  data.getCell('A2').value = 2;
  await data.protect('', {});
  wb.addWorksheet('Other').getCell('A1').value = 'x';
  wb.addWorksheet('Hidden').state = 'hidden';
  sheets = new Uint8Array(await wb.xlsx.writeBuffer());
  schemaBook = await controllerWorkbook();
});

async function store(features: Partial<EditorFeatures> = {}, schema = false): Promise<WorkbookStore> {
  const s = await testStore({ features, schemas: schema ? [workbookSchema([logSheet()])] : undefined });
  await opened(s, schema ? schemaBook : sheets, 'features.xlsx');
  return s;
}

function html(s: WorkbookStore, el: ReactElement): string {
  return renderToStaticMarkup(createElement(I18nextProvider, { i18n: s.i18n, defaultNS: NAMESPACE }, el));
}

const toolbar = (s: WorkbookStore) => html(s, createElement(Toolbar, { store: s, onOpen: () => undefined }));
const tabs = (s: WorkbookStore) => html(s, createElement(SheetTabs, { store: s }));

describe('features', () => {
  it('are all on unless the host turns them off', async () => {
    const s = await testStore({ features: { titleBar: false } });
    expect(Object.entries(s.features).filter(([, on]) => !on)).toEqual([['titleBar', false]]);
  });

  it('titleBar hides the name, file name, changes and schema chip, keeping the actions', async () => {
    const on = await store({}, true);
    expect(toolbar(on)).toContain('wbe-brand');
    expect(toolbar(on)).toContain('features.xlsx');
    const off = await store({ titleBar: false }, true);
    const bar = toolbar(off);
    expect(bar).not.toContain('wbe-brand');
    expect(bar).not.toContain('features.xlsx');
    expect(bar).not.toContain('wbe-chip');
    expect(bar).toContain('wbe-toolbar-actions');
  });

  it('statusBar hides the status bar', async () => {
    const on = await store();
    expect(html(on, createElement(StatusBar, { store: on, view: on.view, selection: on.selection }))).toContain('wbe-status-bar');
    const off = await store({ statusBar: false });
    expect(html(off, createElement(StatusBar, { store: off, view: off.view, selection: off.selection }))).toBe('');
  });

  it('sheetTabButtons hides adding a sheet and the sheet menu button, keeping the tabs', async () => {
    const on = tabs(await store());
    expect(on).toContain('aria-label="New sheet"');
    expect(on).toContain('aria-label="Sheet menu"');
    const off = tabs(await store({ sheetTabButtons: false }));
    expect(off).not.toContain('aria-label="New sheet"');
    expect(off).not.toContain('aria-label="Sheet menu"');
    expect(off).toContain('role="tab"');
  });

  it('sheetMenus takes away the tab menu, dragging, unhiding and renaming, keeping the tabs and adding a sheet', async () => {
    const on = await store();
    expect(tabs(on)).toContain('draggable="true"');
    expect(tabs(on)).toMatch(/<button[^>]*aria-haspopup="dialog"[^>]*>1 hidden/);
    on.set('sheetMenu', { sheet: 0, x: 0, y: 0 });
    shown.items = undefined;
    tabs(on);
    expect(menuShown()?.length).toBeGreaterThan(0);

    const off = await store({ sheetMenus: false });
    const markup = tabs(off);
    expect(markup).not.toContain('draggable="true"');
    expect(markup).not.toContain('aria-label="Sheet menu"');
    expect(markup).toContain('aria-label="New sheet"');
    expect(markup).toMatch(/<span[^>]*>1 hidden/);
    expect(markup.match(/role="tab"/g)).toHaveLength(2);
    off.set('sheetMenu', { sheet: 0, x: 0, y: 0 });
    off.set('renameSheetFor', 0);
    shown.items = undefined;
    expect(tabs(off)).not.toContain('role="dialog"');
    expect(menuShown()).toBeUndefined();
    // Choosing a tab still shows its sheet.
    off.setActive(1);
    expect(off.active).toBe(1);
  });

  it('lockNotes leaves out the note that a cell is locked, going on to what else the cell has', async () => {
    const on = await store();
    expect(cellNote(on.t, on, on.view!, on.selection)).toMatchObject({ tone: 'lock' });
    const off = await store({ lockNotes: false });
    expect(cellNote(off.t, off, off.view!, off.selection)?.tone).not.toBe('lock');
    // Still read-only: only the note goes.
    expect(off.view!.readOnlyReason(0, 0)).toBeTruthy();
  });

  it.each<[keyof EditorFeatures, string[]]>([
    ['undoButtons', ['aria-label="Undo"', 'aria-label="Redo"']],
    ['recordButtons', ['aria-label="Insert row"', 'aria-label="Delete row"']],
    ['exactNumbers', ['Exact numbers']],
    ['highlightEdits', ['Highlight changes']],
    ['schemaLabels', ['Mark schema labels']],
    ['shortcuts', ['aria-label="Keyboard shortcuts"']],
  ])('%s hides its toolbar actions', async (flag, marks) => {
    const on = toolbar(await store({}, true));
    for (const mark of marks) expect(on).toContain(mark);
    const off = toolbar(await store({ [flag]: false }, true));
    for (const mark of marks) expect(off).not.toContain(mark);
  });

  it('leave the More panel empty when every action in it is off, so its button can go', async () => {
    const off = await store({ exactNumbers: false, highlightEdits: false, schemaLabels: false, shortcuts: false, schemas: false, report: false, format: false, download: false });
    expect(toolbar(off)).toMatch(/<div class="wbe-toolbar-more" id="[^"]*"><\/div><button[^>]*wbe-toolbar-more-btn/);
  });
});
