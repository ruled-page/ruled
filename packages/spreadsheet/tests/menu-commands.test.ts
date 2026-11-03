import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { I18nextProvider } from 'react-i18next';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { EditorHandle } from '../src/editor/handle';
import { applyLiveOptions, changedOptions } from '../src/editor/liveOptions';
import { editorKey, type GridKey } from '../src/grid/keys';
import { english, NAMESPACE } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { keepItems, type MenuItem } from '../src/ui/ContextMenu';
import { MENU_COMMANDS, SHORTCUT_COMMANDS, shortcutOn, type MenuCommandId, type MenuOptions } from '../src/ui/menuCommands';
import { shortcuts } from '../src/ui/shortcuts';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

// The menu as GridMenu builds it, caught before it would be shown: showing it needs a browser.
const shown = vi.hoisted(() => ({ items: [] as MenuItem[] }));
vi.mock('../src/ui/ContextMenu', async (original) => ({
  ...(await original<typeof import('../src/ui/ContextMenu')>()),
  ContextMenu: ({ items }: { items: MenuItem[] }) => {
    shown.items = items;
    return null;
  },
}));
const { GridMenu } = await import('../src/ui/GridMenu');

let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Data');
  sheet.getCell('A1').value = 'Region';
  sheet.getCell('A2').value = 'North';
  // Beyond A1, so A1 alone isn't whole rows or columns.
  sheet.getCell('C3').value = 1;
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

async function store(language: string, menus?: MenuOptions): Promise<WorkbookStore> {
  const s = await testStore({ language, menus });
  await opened(s, bytes, 'menus.xlsx');
  return s;
}

/** The menu GridMenu shows: of a cell, or of a row or column heading. */
function menu(s: WorkbookStore, kind: 'cell' | 'row' | 'column'): MenuItem[] {
  shown.items = [];
  s.set('menu', { x: 0, y: 0, kind });
  renderToStaticMarkup(createElement(I18nextProvider, { i18n: s.i18n, defaultNS: NAMESPACE }, createElement(GridMenu, { store: s })));
  return shown.items;
}

/** The menu's entries as `id label`, submenus' indented, separators as `-`. */
function lines(items: MenuItem[], indent = ''): string[] {
  return items.flatMap((item) => ('separator' in item ? [`${indent}-`] : [`${indent}${item.id} ${item.label}`, ...('items' in item ? lines(item.items, `${indent}  `) : [])]));
}

const ids = (items: MenuItem[]): string[] => items.flatMap((item) => ('separator' in item ? [] : [item.id!, ...('items' in item ? ids(item.items) : [])]));

describe('menu command ids', () => {
  it('names every entry of the cell and heading menus with an id from MENU_COMMANDS, in the menus it says', async () => {
    const s = await store('en');
    const where = { cell: menu(s, 'cell'), header: [...menu(s, 'row'), ...menu(s, 'column')] };
    for (const [kind, items] of Object.entries(where)) {
      for (const id of ids(items)) {
        expect(MENU_COMMANDS, id).toHaveProperty([id]);
        expect(['both', kind], id).toContain(MENU_COMMANDS[id as MenuCommandId].menus);
      }
    }
    // Labels that change with the selection or state keep their entry's id.
    s.selectLines('rows', [[0, 2]], { row: 0, col: 0 });
    expect(lines(menu(s, 'row'))).toContain('menu.insertRowsAbove Insert 3 rows above');
    expect(lines(menu(s, 'row'))).toContain('menu.filterOn Add filter buttons');
    await s.toggleFilter();
    expect(lines(menu(s, 'row'))).toContain('menu.filterOn Remove filter buttons');
  });

  it('keeps the entries listed, the same in English and German, dropping separators left with nothing between', async () => {
    const menus: MenuOptions = { cell: ['menu.copy', 'menu.columns', 'outline.menu.groupRows', 'menu.find'], header: ['menu.copy', 'menu.hideRows'] };
    const en = await store('en', menus);
    const de = await store('de', menus);
    expect(lines(menu(en, 'cell'))).toEqual([
      'menu.copy Copy',
      '-',
      'menu.rows Rows',
      '  outline.menu.groupRows Group rows',
      'menu.columns Columns',
      '  menu.insertColumnsLeft Insert 1 column left',
      '  menu.insertColumnsRight Insert 1 column right',
      '  menu.deleteColumns Delete 1 column',
      '  -',
      '  menu.columnWidth Column width…',
      '  menu.autofitColumnWidth AutoFit column width',
      '  menu.hideColumns Hide column',
      '  -',
      '  outline.menu.groupColumns Group columns',
      '-',
      'menu.find Find and replace…',
    ]);
    expect(ids(menu(de, 'cell'))).toEqual(ids(menu(en, 'cell')));
    expect(lines(menu(de, 'cell'))[0]).toBe('menu.copy Kopieren');
    expect(lines(menu(en, 'row'))).toEqual(['menu.copy Copy', '-', 'menu.hideRows Hide row']);
    expect(lines(menu(de, 'row'))).toEqual(['menu.copy Kopieren', '-', 'menu.hideRows Zeile ausblenden']);
    // The column heading's menu has neither, but for Copy.
    expect(ids(menu(de, 'column'))).toEqual(['menu.copy']);
  });

  it('keeps the rows the cell menu shows in place of its Rows submenu with it, when whole rows are selected', async () => {
    const s = await store('de', { cell: ['menu.rows'] });
    s.selectLines('rows', [[0, 0]], { row: 0, col: 0 });
    expect(ids(menu(s, 'cell'))).toEqual([
      'menu.insertRowsAbove',
      'menu.insertRowsBelow',
      'menu.deleteRows',
      'menu.rowHeight',
      'menu.autofitRowHeight',
      'menu.hideRows',
      'outline.menu.groupRows',
    ]);
  });

  it('opens no menu when it would be empty', async () => {
    const s = await store('en', { header: [] });
    expect(menu(s, 'column')).toEqual([]);
    expect(menu(s, 'cell').length).toBeGreaterThan(10);
  });
});

describe('keeping menu entries', () => {
  const item = (id: string, within?: string): MenuItem => ({ id, within, label: id, onSelect: () => undefined });
  const sep = { separator: true } as const;
  const items: MenuItem[] = [sep, item('a'), sep, item('b'), sep, sep, { id: 'sub', label: 'sub', items: [item('c'), sep, item('d')] }, item('e', 'sub'), sep];
  const kept = (keep?: string[]) => lines(keepItems(items, keep));

  it('keeps everything, as it is, without a list', () => {
    expect(keepItems(items, undefined)).toBe(items);
  });

  it('drops separators at either end and in a row', () => {
    expect(kept(['a', 'b'])).toEqual(['a a', '-', 'b b']);
    expect(kept(['b'])).toEqual(['b b']);
    expect(kept([])).toEqual([]);
  });

  it('keeps a whole submenu, or the entries of it listed, and what stands in for it', () => {
    expect(kept(['sub'])).toEqual(['sub sub', '  c c', '  -', '  d d', 'e e']);
    expect(kept(['d'])).toEqual(['sub sub', '  d d']);
    expect(kept(['a', 'c', 'e'])).toEqual(['a a', '-', 'sub sub', '  c c', 'e e']);
  });
});

describe('shortcuts of menu entries', () => {
  it('are on without the option, or while either menu keeps their entry', () => {
    for (const id of SHORTCUT_COMMANDS) expect(shortcutOn(undefined, id)).toBe(true);
    expect(shortcutOn({ cell: [] }, 'menu.copy')).toBe(true); // the heading menu, left whole, has Copy
    expect(shortcutOn({ cell: [], header: [] }, 'menu.copy')).toBe(false);
    expect(shortcutOn({ cell: [], header: ['menu.copy'] }, 'menu.copy')).toBe(true);
    expect(shortcutOn({ cell: [], header: [], shortcuts: ['menu.copy'] }, 'menu.copy')).toBe(true);
    // Insert Cells is only in the cell menu, so a whole heading menu doesn't keep it.
    expect(shortcutOn({ cell: [] }, 'menu.insertCells')).toBe(false);
    // Kept with the submenu they're in.
    expect(shortcutOn({ cell: ['menu.rows'], header: [] }, 'outline.menu.groupRows')).toBe(true);
    expect(shortcutOn({ cell: ['charts.menu.insert'], header: [] }, 'core.chartType.column')).toBe(true);
    // Shortcuts that aren't a menu entry's aren't touched.
    expect(shortcutOn({ cell: [], header: [] }, 'menu.merge')).toBe(true);
  });

  it('know where their entries are', () => {
    for (const id of SHORTCUT_COMMANDS) {
      const inHeader = MENU_COMMANDS[id].menus !== 'cell';
      expect(shortcutOn({ cell: [], header: [id] }, id), id).toBe(inHeader);
      expect(shortcutOn({ cell: [id], header: [] }, id), id).toBe(true);
    }
  });

  it('are taken by the grid only while on', async () => {
    const key = (k: Partial<GridKey>): GridKey => ({ key: '', keyCode: 0, shiftKey: false, ctrlKey: false, metaKey: false, altKey: false, ...k });
    const altF5 = key({ key: 'F5', altKey: true });
    const off = await store('de', { cell: [], header: [] });
    expect(editorKey(off, altF5, () => undefined)).toBe(false);
    off.setMenus({ cell: [], header: [], shortcuts: ['pivot.menu.refresh'] });
    expect(editorKey(off, altF5, () => undefined)).toBe(true);
    expect(editorKey(off, key({ key: 'F5', altKey: true, ctrlKey: true }), () => undefined)).toBe(false);
  });

  it('leave the list of shortcuts with their entries', () => {
    const t = english;
    const all = shortcuts(t, false).flatMap((g) => g.keys.map(([, does]) => does));
    const without = shortcuts(t, false, (id) => shortcutOn({ cell: [], header: [], shortcuts: ['menu.copy'] }, id)).flatMap((g) => g.keys.map(([, does]) => does));
    expect(all).toContain(t('shortcuts.find'));
    expect(without).not.toContain(t('shortcuts.find'));
    expect(without).toContain(t('shortcuts.clipboard'));
    expect(without).toContain(t('shortcuts.undoRedo'));
  });
});

describe('the menus option', () => {
  it('changes while the editor is open, a new literal of the same lists being no change', () => {
    const calls: unknown[] = [];
    const handle = { setMenus: (m: unknown) => calls.push(m) } as unknown as EditorHandle;
    expect(changedOptions({ menus: { cell: ['menu.copy'] } }, { menus: { cell: ['menu.copy'] } })).toEqual([]);
    applyLiveOptions(handle, { menus: { cell: ['menu.copy'] } }, { menus: { cell: ['menu.copy', 'menu.find'] } });
    expect(calls).toEqual([{ cell: ['menu.copy', 'menu.find'] }]);
  });

  it('has every id documented', () => {
    const page = readFileSync(new URL('../../../docs/src/content/docs/guides/menus-and-shortcuts.md', import.meta.url), 'utf8');
    for (const id of Object.keys(MENU_COMMANDS)) expect(page, id).toContain(`\`${id}\``);
  });
});
