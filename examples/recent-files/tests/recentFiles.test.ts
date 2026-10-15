import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { FileSource, RuledPlugin } from '@ruled/sheet/plugin';
import { handleFor } from '../../../packages/spreadsheet/src/editor/handle';
import type { WorkbookStore } from '../../../packages/spreadsheet/src/state/store';
import { sample } from '../../../packages/spreadsheet/tests/helpers';
import { testStore } from '../../../packages/spreadsheet/tests/storeHarness';
import { recentFiles, RECENT_FILES_ID, type RecentFile, type RecentFilesOptions } from '../src/recentFiles';

const BUDGET = sample('budget.xlsx');
const FORECAST = sample('forecast.xlsx');
const buffer = (bytes: Uint8Array) => bytes.slice().buffer as ArrayBuffer;

/** Just enough DOM for the plugin's list: elements that hold children, text and click listeners. */
class FakeElement {
  children: FakeElement[] = [];
  textContent = '';
  className = '';
  title = '';
  style = { cssText: '' };
  private attrs = new Map<string, string>();
  private clicks: (() => void)[] = [];
  constructor(readonly tagName: string) {}
  append(...kids: FakeElement[]) {
    this.children.push(...kids);
  }
  replaceChildren(...kids: FakeElement[]) {
    this.children = kids;
  }
  setAttribute(name: string, value: string) {
    this.attrs.set(name, value);
  }
  getAttribute(name: string) {
    return this.attrs.get(name) ?? null;
  }
  addEventListener(_: 'click', fn: () => void) {
    this.clicks.push(fn);
  }
  click() {
    for (const fn of this.clicks) fn();
  }
  /** Its text and its children's, joined by spaces. */
  text(): string {
    return [this.textContent, ...this.children.map((c) => c.text())].filter(Boolean).join(' ');
  }
  /** Every element under it with this tag. */
  all(tag: string): FakeElement[] {
    return this.children.flatMap((c) => [...(c.tagName === tag ? [c] : []), ...c.all(tag)]);
  }
}

let stored: Map<string, string>;

beforeEach(() => {
  stored = new Map();
  vi.stubGlobal('localStorage', { getItem: (k: string) => stored.get(k) ?? null, setItem: (k: string, v: string) => void stored.set(k, v) });
  vi.stubGlobal('document', { createElement: (tag: string) => new FakeElement(tag) });
});

afterEach(() => vi.unstubAllGlobals());

/** A drive that holds forecast.xlsx, and fetches it again by its origin. */
function fakeDrive() {
  const fetched: unknown[] = [];
  const source: FileSource = {
    id: 'drive',
    label: 'Drive',
    pick: async () => ({ bytes: FORECAST, origin: { source: 'drive', name: 'forecast.xlsx', ref: { item: 7 }, version: '1', location: 'Shared/forecast.xlsx' } }),
    save: async (_, origin) => origin,
    fetch: async (origin) => {
      fetched.push(origin.ref);
      return { bytes: FORECAST, origin: { ...origin, version: '2' } };
    },
  };
  const plugin: RuledPlugin = { id: 'test.drive', apiVersion: 1, setup: (ctx) => void ctx.files.registerSource(source) };
  return { source, plugin, fetched };
}

/** An editor running the plugin, as createWorkbookEditor makes one. */
async function editor(options: RecentFilesOptions = {}, more: RuledPlugin[] = []) {
  const store = await testStore({ plugins: [recentFiles(options), ...more] });
  const handle = handleFor(store);
  await store.startPlugins(handle);
  const host = store.plugins!;
  const panel = host.panels.find((p) => p.plugin === RECENT_FILES_ID)!;
  const clear = host.commands.find((c) => c.plugin === RECENT_FILES_ID)!;
  const source = host.sources.find((s) => s.plugin === RECENT_FILES_ID)!;
  /** The panel, drawn into an element. */
  const showPanel = () => {
    const body = new FakeElement('div');
    const cleanup = panel.render(body as unknown as HTMLElement);
    return { body, cleanup: cleanup as () => void };
  };
  return { store, handle, host, panel, clear, source, showPanel };
}

/** Opens a file as the Open button does, and waits for the plugin to list it. */
async function openFromDevice(store: WorkbookStore, name: string, bytes: Uint8Array) {
  await store.open(name, buffer(bytes), { from: 'device' });
  await settled();
}

/** Lets the plugin's storage writes finish. */
const settled = () => new Promise((r) => setTimeout(r, 0));

const kept = (): RecentFile[] => JSON.parse(stored.get(`ruled-plugin:${RECENT_FILES_ID}`) ?? '{}').files ?? [];

describe('recent files', () => {
  it('lists small files from this device with their bytes, newest first, and not the host’s', async () => {
    const { store, handle, showPanel } = await editor({ keepBytesUpTo: 10_000 });
    const { body } = showPanel();
    expect(body.text()).toBe('Workbooks you open from this device or a file source appear here.');

    await openFromDevice(store, 'budget.xlsx', BUDGET);
    await openFromDevice(store, 'forecast.xlsx', FORECAST);
    await handle.open(BUDGET, 'from-the-host.xlsx');
    await settled();

    expect(kept().map((f) => f.origin.name)).toEqual(['forecast.xlsx', 'budget.xlsx']);
    expect(atob(kept()[1].bytes!).length).toBe(BUDGET.byteLength);
    // The open panel is drawn again as the list changes.
    expect(body.all('li').map((li) => li.text())).toEqual(['forecast.xlsx This device ×', 'budget.xlsx This device ×']);
    store.dispose();
  });

  it('leaves out files too large to keep, and keeps at most `max`', async () => {
    const { store } = await editor({ keepBytesUpTo: 6_900, max: 1 });
    await openFromDevice(store, 'budget.xlsx', BUDGET);
    await openFromDevice(store, 'kitchen-sink.xlsx', sample('kitchen-sink.xlsx'));
    await openFromDevice(store, 'forecast.xlsx', FORECAST);
    expect(kept().map((f) => f.origin.name)).toEqual(['forecast.xlsx']);
    store.dispose();
  });

  it('opens a file from this device again from the bytes kept, moving it to the top', async () => {
    const { store, showPanel } = await editor();
    await openFromDevice(store, 'budget.xlsx', BUDGET);
    await openFromDevice(store, 'forecast.xlsx', FORECAST);

    const { body } = showPanel();
    body.all('button').find((b) => b.text().startsWith('budget.xlsx'))!.click();
    await vi.waitFor(() => expect(store.fileName).toBe('budget.xlsx'));
    await settled();
    expect(new Uint8Array(store.openedFile()!.bytes())).toEqual(BUDGET);
    expect(store.origin).toMatchObject({ source: 'recent', location: 'This device' });
    // Saving downloads, as it did when the file was first opened.
    expect(store.saveTarget).toBe('download');
    expect(kept().map((f) => f.origin.name)).toEqual(['budget.xlsx', 'forecast.xlsx']);
    store.dispose();
  });

  it('keeps only the origin of a file from a source, and opens it again through that source', async () => {
    const drive = fakeDrive();
    const { store, host, showPanel } = await editor({}, [drive.plugin]);
    await host.openFrom(host.sources.find((s) => s.id === 'drive')!);
    await settled();
    expect(kept()).toEqual([expect.objectContaining({ id: 'drive:{"item":7}', origin: expect.objectContaining({ name: 'forecast.xlsx', version: '1' }) })]);
    expect(kept()[0].bytes).toBeUndefined();

    await openFromDevice(store, 'budget.xlsx', BUDGET);
    showPanel().body.all('li')[1].children[0].click();
    await vi.waitFor(() => expect(store.fileName).toBe('forecast.xlsx'));
    expect(drive.fetched).toEqual([{ item: 7 }]);
    // It is the source's file again, as it is now, so it saves back there.
    expect(store.origin).toMatchObject({ source: 'drive', version: '2' });
    expect(store.saveTarget).toBe('source');
    store.dispose();
  });

  it('offers the files as a submenu of Open, opening them through their sources', async () => {
    const drive = fakeDrive();
    const { store, host, source } = await editor({}, [drive.plugin]);
    await host.openFrom(host.sources.find((s) => s.id === 'drive')!);
    await openFromDevice(store, 'budget.xlsx', BUDGET);

    const items = host.items(source);
    expect(items).toEqual([
      { source: 'recent', name: 'budget.xlsx', ref: 'device:budget.xlsx', location: 'This device' },
      expect.objectContaining({ source: 'drive', name: 'forecast.xlsx', location: 'Shared/forecast.xlsx' }),
    ]);
    // No Browse…: there's nothing to pick.
    expect(source.pick).toBeUndefined();

    expect(await host.openItem(items[1])).toBe(true);
    expect(store.origin).toMatchObject({ source: 'drive', version: '2' });
    expect(await host.openItem(items[0])).toBe(true);
    expect(store.origin).toMatchObject({ source: 'recent', name: 'budget.xlsx' });
    store.dispose();
  });

  it('says so when a file from this device is no longer kept', async () => {
    const { store, host, source } = await editor();
    await openFromDevice(store, 'budget.xlsx', BUDGET);
    const [item] = host.items(source);
    stored.clear();
    const again = await editor();
    expect(await again.host.openItem(item)).toBe(false);
    expect(again.store.t.text(again.store.toasts.at(-1)!.text)).toBe('budget.xlsx is no longer kept in this browser. Open it again from this device.');
    store.dispose();
    again.store.dispose();
  });

  it('forgets a file, or clears the list after asking, the command enabled only while there is one', async () => {
    const { store, host, clear, showPanel } = await editor();
    expect(host.commandState(clear, {}).enabled).toBe(false);
    await openFromDevice(store, 'budget.xlsx', BUDGET);
    await openFromDevice(store, 'forecast.xlsx', FORECAST);
    expect(host.commandState(clear, {}).enabled).toBe(true);

    const { body } = showPanel();
    const forget = body.all('button').find((b) => b.getAttribute('aria-label') === 'Remove forecast.xlsx from the list')!;
    forget.click();
    await settled();
    expect(kept().map((f) => f.origin.name)).toEqual(['budget.xlsx']);

    host.run(clear, {});
    await vi.waitFor(() => expect(store.confirmRequest).toBeDefined());
    expect(store.t.text(store.confirmRequest!.title)).toBe('Clear recent files');
    store.confirmRequest!.resolve('confirm');
    await vi.waitFor(() => expect(kept()).toEqual([]));
    expect(host.commandState(clear, {}).enabled).toBe(false);
    store.dispose();
  });

  it('remembers the list in this browser, and speaks the editor’s language', async () => {
    const first = await editor();
    await openFromDevice(first.store, 'budget.xlsx', BUDGET);
    first.store.dispose();

    const { store, showPanel } = await editor();
    store.setLanguage('de');
    await settled();
    const { body, cleanup } = showPanel();
    expect(body.all('li').map((li) => li.text())).toEqual(['budget.xlsx Dieses Gerät ×']);
    cleanup();
    store.dispose();
  });
});
