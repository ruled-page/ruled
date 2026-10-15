import { afterEach, describe, expect, it, vi } from 'vitest';
import { forecastSchema } from '@ruled/demo/schemas';
import { handleFor, type EditorHandle } from '../src/editor/handle';
import { english } from '../src/i18n';
import type { FileOrigin, FileSource, PluginContext, RuledPlugin } from '../src/plugin/types';
import { definePlugin } from '../src/plugin';
import { schemaFromJSON } from '../src/schema/json';
import type { EditorEvent, SavedWorkbook, StoreOptions, WorkbookStore } from '../src/state/store';
import { sample } from './helpers';
import { opened, testStore } from './storeHarness';

const FORECAST = sample('forecast.xlsx');

/** A store and its handle, with its plugins started, as createWorkbookEditor makes them. */
async function editor(options: StoreOptions = {}): Promise<{ store: WorkbookStore; handle: EditorHandle; events: EditorEvent[] }> {
  const events: EditorEvent[] = [];
  const store = await testStore({ onEvent: (e) => void events.push(e), ...options });
  const handle = handleFor(store);
  await store.startPlugins(handle);
  return { store, handle, events };
}

/** A plugin that runs `setup` with its context. */
const plugin = (setup: (ctx: PluginContext) => void | (() => void), extra: Partial<RuledPlugin> = {}): RuledPlugin =>
  definePlugin({ id: 'test.plugin', apiVersion: 1, setup, ...extra });

/** Answers the dialog the store is about to show. */
async function answer(store: WorkbookStore, with_: 'confirm' | 'cancel'): Promise<string> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  const req = store.confirmRequest!;
  req.resolve(with_);
  return english.text(req.title);
}

/** A drive holding one file, with a version that moves on each save, as an ETag does. */
function fakeDrive(options: { writable?: boolean } = {}) {
  let version = 1;
  const writes: { file: SavedWorkbook; origin: FileOrigin }[] = [];
  const copies: SavedWorkbook[] = [];
  const source: FileSource = {
    id: 'fake.drive',
    label: { key: 'plugin.test.plugin.drive' },
    pick: async () => ({ bytes: FORECAST, origin: { source: 'fake.drive', name: 'forecast.xlsx', ref: { item: 7 }, version: '1', writable: options.writable, location: 'Shared/forecast.xlsx' } }),
    save: async (file, origin) => {
      if (origin.version !== String(version)) throw Object.assign(new Error(''), { name: 'SaveConflictError' });
      writes.push({ file, origin });
      return { ...origin, version: String(++version) };
    },
    saveAs: async (file) => {
      copies.push(file);
      return { source: 'fake.drive', name: 'copy.xlsx', ref: { item: 8 } };
    },
  };
  return {
    source,
    writes,
    copies,
    /** Someone else saves the file meanwhile. */
    bump: () => version++,
    plugin: plugin((ctx) => void ctx.files.registerSource(source), { translations: { en: { drive: 'Fake Drive' }, de: { drive: 'Testlaufwerk' } } }),
  };
}

afterEach(() => vi.unstubAllGlobals());

describe('plugin registries', () => {
  it('tells plugins the largest file the editor opens', async () => {
    const seen: number[] = [];
    await editor({ plugins: [plugin((ctx) => void seen.push(ctx.limits.maxFileBytes))] });
    await editor({ limits: { maxFileBytes: 1024 }, plugins: [plugin((ctx) => void seen.push(ctx.limits.maxFileBytes))] });
    expect(seen).toEqual([50 * 1024 * 1024, 1024]);
  });

  it('gives plugins their text in the editor\'s language, for their own DOM', async () => {
    let ctx!: PluginContext;
    const { handle } = await editor({ plugins: [plugin((c) => void (ctx = c), { translations: { en: { hi: 'Hello {{name}}' }, de: { hi: 'Hallo {{name}}' } } })] });
    expect(ctx.text({ key: 'plugin.test.plugin.hi', params: { name: 'Ada' } })).toBe('Hello Ada');
    handle.setLanguage('de');
    await vi.waitFor(() => expect(ctx.text({ key: 'plugin.test.plugin.hi', params: { name: 'Ada' } })).toBe('Hallo Ada'));
    expect(ctx.text({ en: 'Plain', de: 'Schlicht' })).toBe('Schlicht');
  });

  it('keeps what plugins register under their id, and drops it when they remove it', async () => {
    const removers: (() => void)[] = [];
    const { store } = await editor({
      plugins: [
        plugin((ctx) => {
          removers.push(
            ctx.commands.register({ id: 'hello', label: 'Hello', where: 'toolbar', run: () => undefined }),
            ctx.commands.register({ id: 'more', label: 'More', where: 'more', run: () => undefined }),
            ctx.files.registerSource({ id: 'drive', label: 'Drive', pick: async () => null }),
            ctx.ui.registerPanel({ id: 'recent', title: 'Recent', render: () => undefined }),
          );
        }),
      ],
    });
    expect(store.plugins!.active).toBe(true);
    expect(store.plugins!.commands.map((c) => [c.key, c.where])).toEqual([
      ['test.plugin/hello', 'toolbar'],
      ['test.plugin/more', 'more'],
    ]);
    expect(store.plugins!.sources.map((s) => s.key)).toEqual(['test.plugin/drive']);
    expect(store.plugins!.panels.map((p) => p.key)).toEqual(['test.plugin/recent']);
    for (const remove of removers) remove();
    expect([store.plugins!.commands, store.plugins!.sources, store.plugins!.panels]).toEqual([[], [], []]);
    store.dispose();
  });

  it('reads commands’ state again when a plugin says what they depend on changed', async () => {
    let ctx: PluginContext | undefined;
    const { store, handle } = await editor({ plugins: [plugin((c) => void (ctx = c))] });
    const heard = vi.fn();
    handle.subscribe(heard);
    ctx!.commands.refresh();
    expect(heard).toHaveBeenCalled();
    store.dispose();
  });

  it('refuses a second command or source with the same id', async () => {
    const { store, events } = await editor({
      plugins: [
        plugin((ctx) => {
          ctx.files.registerSource({ id: 'drive', label: 'Drive', pick: async () => null });
          ctx.files.registerSource({ id: 'drive', label: 'Drive', pick: async () => null });
        }),
        plugin((ctx) => void ctx.commands.register({ label: 'No id', where: 'more', run: () => undefined } as never), { id: 'other.plugin' }),
      ],
    });
    expect(events.filter((e) => e.type === 'pluginError').map((e) => e.type === 'pluginError' && e.message)).toEqual([
      '"drive" is already registered.',
      "Every one of a plugin's commands needs an id.",
    ]);
    // Setup threw, so the plugin was stopped and what it added went with it.
    expect(store.plugins!.sources).toEqual([]);
    store.dispose();
  });

  it('merges plugin messages under plugin.<id>., the host rewording them', async () => {
    const { store } = await editor({
      language: 'de',
      plugins: [plugin(() => undefined, { translations: { en: { open: 'Open from Drive', save: 'Save to Drive' }, de: { open: 'Aus Drive öffnen' } } })],
      translations: { de: { 'plugin.test.plugin.save': 'In Firmendateien speichern' } },
    });
    expect(store.t.text({ key: 'plugin.test.plugin.open' } as never)).toBe('Aus Drive öffnen');
    expect(store.t.text({ key: 'plugin.test.plugin.save' } as never)).toBe('In Firmendateien speichern');
    store.setLanguage('en');
    expect(store.t.text({ key: 'plugin.test.plugin.save' } as never)).toBe('Save to Drive');
    store.dispose();
  });

  it('gives cell-menu commands the selection, and re-reads whether they show', async () => {
    const runs: unknown[] = [];
    const { store } = await editor({
      plugins: [
        plugin((ctx) => {
          ctx.commands.register({
            id: 'cell',
            label: 'Look up',
            where: 'cellMenu',
            run: (c) => void runs.push(c.selection),
            visible: (c) => c.selection?.sheet === 'Forecast',
            enabled: (c) => c.selection?.active !== 'A1',
          });
        }),
      ],
    });
    await opened(store, FORECAST, 'forecast.xlsx');
    const [command] = store.plugins!.commands;
    const at = () => ({ selection: store.commandSelection() });
    expect(store.plugins!.commandState(command, at())).toEqual({ visible: true, enabled: false });
    store.setSelection({ row: 1, col: 1, rects: [{ r1: 1, c1: 1, r2: 2, c2: 3 }] });
    expect(store.plugins!.commandState(command, at())).toEqual({ visible: true, enabled: true });
    store.plugins!.run(command, at());
    expect(runs).toEqual([{ sheet: 'Forecast', active: 'B2', ranges: ['B2:D3'] }]);
    store.dispose();
  });

  it('lets plugins hear events, keep values in the browser, and clean up when the editor goes', async () => {
    const items = new Map<string, string>();
    vi.stubGlobal('localStorage', { getItem: (k: string) => items.get(k) ?? null, setItem: (k: string, v: string) => void items.set(k, v) });
    const heard: string[] = [];
    let ctx: PluginContext | undefined;
    const cleanup = vi.fn();
    const { store, handle } = await editor({
      plugins: [
        plugin((c) => {
          ctx = c;
          c.on((e) => void heard.push(e.type));
          c.on(() => void heard.push('never'))();
          return cleanup;
        }),
      ],
    });
    expect(ctx!.editor).toBe(handle);
    await ctx!.storage.set('recent', ['a.xlsx']);
    await ctx!.storage.set('folder', 'Finance');
    await ctx!.storage.delete('folder');
    expect(await ctx!.storage.get('recent')).toEqual(['a.xlsx']);
    expect(await ctx!.storage.get('folder')).toBeUndefined();
    expect(await ctx!.storage.get('constructor')).toBeUndefined();
    expect(JSON.parse(items.get('ruled-plugin:test.plugin')!)).toEqual({ recent: ['a.xlsx'] });

    await opened(store, FORECAST, 'forecast.xlsx');
    expect(heard).toContain('opened');
    expect(heard).not.toContain('never');
    handle.destroy();
    expect(ctx!.signal.aborted).toBe(true);
    expect(cleanup).toHaveBeenCalledOnce();
  });

  it('keeps nothing, without an error, when the browser has no storage', async () => {
    vi.stubGlobal('localStorage', {
      getItem: () => 'not json',
      setItem: () => {
        throw new Error('QuotaExceededError');
      },
    });
    let ctx: PluginContext | undefined;
    const { store } = await editor({ plugins: [plugin((c) => void (ctx = c))] });
    await ctx!.storage.set('recent', 1);
    expect(await ctx!.storage.get('recent')).toBeUndefined();
    store.dispose();
  });

  it('shows plugin toasts, questions and dialogs in the editor', async () => {
    let ctx: PluginContext | undefined;
    const { store } = await editor({ plugins: [plugin((c) => void (ctx = c))] });
    ctx!.ui.toast('info', 'Signed in');
    expect(store.toasts.at(-1)).toMatchObject({ kind: 'info', text: 'Signed in' });

    const asked = ctx!.ui.confirm({ title: 'Sign out?', body: 'You can sign in again.', confirmLabel: 'Sign out' });
    expect(await answer(store, 'confirm')).toBe('Sign out?');
    expect(await asked).toBe(true);

    const first = ctx!.ui.dialog({ title: 'Pick', render: () => undefined });
    // A second dialog closes the first.
    const second = ctx!.ui.dialog({ title: 'Folder', render: () => undefined, actions: [{ id: 'ok', label: 'OK', primary: true }] });
    expect(await first).toBeNull();
    expect(store.pluginDialog).toMatchObject({ plugin: 'test.plugin', title: 'Folder' });
    store.pluginDialog!.resolve('ok');
    expect(await second).toBe('ok');
    expect(store.pluginDialog).toBeUndefined();
    store.dispose();
  });
});

describe('refusing and containing plugins', () => {
  it('refuses a plugin for a contract version it doesn’t know, and starts the rest', async () => {
    const started = vi.fn();
    const { store, events } = await editor({
      plugins: [
        { id: 'future', apiVersion: 2, setup: started } as unknown as RuledPlugin,
        { id: 'nothing' } as unknown as RuledPlugin,
        plugin(() => undefined),
        plugin(started),
      ],
    });
    expect(started).not.toHaveBeenCalled();
    expect(events.filter((e) => e.type === 'pluginError')).toEqual([
      { type: 'pluginError', plugin: 'future', message: expect.stringContaining('version 2 of the plugin contract') },
      { type: 'pluginError', plugin: 'nothing', message: expect.stringContaining('needs an id and a setup function') },
      { type: 'pluginError', plugin: 'test.plugin', message: expect.stringContaining('another plugin has the same id') },
    ]);
    expect(store.toasts.filter((t) => t.kind === 'error')).toHaveLength(3);
    expect(store.plugins!.active).toBe(true);
    store.dispose();
  });

  it('stops a plugin whose setup throws or rejects, removing what it added, and carries on', async () => {
    let late: PluginContext | undefined;
    const { store, events } = await editor({
      plugins: [
        plugin((ctx) => {
          ctx.commands.register({ id: 'a', label: 'A', where: 'toolbar', run: () => undefined });
          throw new Error('no client id');
        }),
        plugin(
          (ctx) => {
            late = ctx;
            ctx.commands.register({ id: 'b', label: 'B', where: 'toolbar', run: () => undefined });
            ctx.ui.registerPanel({ id: 'p', title: 'P', render: () => undefined });
            return new Promise((_, reject) => setTimeout(() => reject(new Error('later')))) as unknown as void;
          },
          { id: 'async.plugin' },
        ),
      ],
    });
    // Its panel is open as it fails: the panel closes with it.
    store.set('panel', 'plugin:async.plugin/p');
    await new Promise((r) => setTimeout(r, 5));
    expect(store.panel).toBeNull();
    // A plugin stopped can't add anything more.
    late!.commands.register({ id: 'c', label: 'C', where: 'toolbar', run: () => undefined })();
    expect(store.plugins!.commands).toEqual([]);
    expect(store.plugins!.active).toBe(false);
    expect(events.filter((e) => e.type === 'pluginError').map((e) => e.type === 'pluginError' && [e.plugin, e.message])).toEqual([
      ['test.plugin', 'no client id'],
      ['async.plugin', 'later'],
    ]);
    expect(english.text(store.toasts[0].text)).toBe('The plugin test.plugin failed: no client id');
    await opened(store, FORECAST, 'forecast.xlsx');
    expect(store.status).toBe('ready');
    store.dispose();
  });

  it('reports commands that throw, hiding one whose checks throw', async () => {
    let ctx: PluginContext | undefined;
    const { store, events } = await editor({
      plugins: [
        plugin((c) => {
          ctx = c;
          c.commands.register({ id: 'run', label: 'Run', where: 'more', run: () => Promise.reject(new Error('offline')), preload: () => { throw new Error('no SDK'); } });
          c.commands.register({ id: 'shy', label: 'Shy', where: 'more', run: () => undefined, visible: () => { throw new Error('bad check'); } });
          c.on(() => {
            throw new Error('listener');
          });
        }),
      ],
    });
    const [run, shy] = store.plugins!.commands;
    store.plugins!.run(run, {});
    store.plugins!.preload([run]);
    expect(store.plugins!.commandState(shy, {})).toEqual({ visible: false, enabled: false });
    expect(store.plugins!.commandState(shy, {})).toEqual({ visible: false, enabled: false });
    await new Promise((r) => setTimeout(r));
    // The listener throws on the error it hears, but doesn't hear its own report, which would never end.
    store.toast('error', 'Something else');
    const messages = events.filter((e) => e.type === 'pluginError').map((e) => e.type === 'pluginError' && e.message);
    expect(messages).toEqual(['no SDK', 'bad check', 'offline', 'listener']);
    expect(ctx!.signal.aborted).toBe(false);
    store.dispose();
  });

  it('says so when a source fails to pick, or gives something that isn’t a file', async () => {
    const { store, events } = await editor({
      plugins: [
        plugin((ctx) => {
          ctx.files.registerSource({ id: 'broken', label: 'Broken', pick: () => Promise.reject(new Error('popup blocked')) });
          ctx.files.registerSource({ id: 'odd', label: 'Odd', pick: async () => ({ bytes: 'nope', origin: { source: 'odd', name: 'x.xlsx', ref: null } }) as never });
          ctx.files.registerSource({ id: 'cancel', label: 'Cancel', pick: async () => null });
        }),
      ],
    });
    const [broken, odd, cancel] = store.plugins!.sources;
    expect(await store.plugins!.openFrom(broken)).toBe(false);
    expect(await store.plugins!.openFrom(odd)).toBe(false);
    expect(await store.plugins!.openFrom(cancel)).toBe(false);
    expect(events.find((e) => e.type === 'pluginError')).toMatchObject({ plugin: 'test.plugin', message: 'popup blocked' });
    expect(english.text(store.toasts.at(-1)!.text)).toMatch(/can't open: it needs the file's bytes/);
    expect(store.status).toBe('empty');
    store.dispose();
  });
});

describe('saving to a plugin’s source', () => {
  it('chooses the primary save: the host’s, then the source’s, then a copy, then a download', async () => {
    const drive = fakeDrive();
    const readOnly = fakeDrive({ writable: false });
    const target = async (options: StoreOptions, d = drive) => {
      const { store } = await editor({ plugins: [d.plugin], ...options });
      await store.plugins!.openFrom(store.plugins!.sources[0]);
      const t = store.saveTarget;
      store.dispose();
      return t;
    };
    expect(await target({})).toBe('source');
    expect(await target({ onSave: () => undefined })).toBe('host');
    expect(await target({}, readOnly)).toBe('copy');
    expect(await target({ features: { sourceSave: false } })).toBe('download');
    expect(await target({ features: { sourceSave: false } }, readOnly)).toBe('download');

    const { store } = await editor({ plugins: [drive.plugin] });
    await opened(store, FORECAST, 'forecast.xlsx');
    expect(store.saveTarget).toBe('download');
    expect(store.copySources.map((s) => s.id)).toEqual(['fake.drive']);
    store.dispose();
  });

  it('writes back to the source beside the host’s Save, each with its own unsaved state', async () => {
    const drive = fakeDrive();
    const hosted: SavedWorkbook[] = [];
    const { store, handle, events } = await editor({ plugins: [drive.plugin], onSave: (f) => void hosted.push(f) });
    expect(await store.plugins!.openFrom(store.plugins!.sources[0])).toBe(true);
    for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
    expect(handle.origin).toMatchObject({ source: 'fake.drive', version: '1', location: 'Shared/forecast.xlsx' });
    expect(store.sourceLabel(store.originSource!)).toBe('Fake Drive');
    expect(store.canSaveToSource).toBe(true);

    await store.input(1, 1, '1500');
    expect([store.unsaved, store.sourceUnsaved]).toEqual([true, true]);

    expect(await store.plugins!.saveToSource()).toBe('saved');
    expect(drive.writes).toHaveLength(1);
    expect(drive.writes[0].origin.version).toBe('1');
    expect(drive.writes[0].file.origin).toMatchObject({ source: 'fake.drive' });
    expect(handle.origin?.version).toBe('2');
    expect([store.unsaved, store.sourceUnsaved]).toEqual([true, false]);
    expect(events.find((e) => e.type === 'savedToSource')).toMatchObject({ source: 'fake.drive', name: 'forecast.xlsx' });
    expect(english.text(store.toasts.at(-1)!.text)).toBe('Saved forecast.xlsx to Fake Drive.');

    // The host's Save is still its own, and is told where the workbook came from.
    await handle.save();
    expect(hosted[0].origin).toMatchObject({ source: 'fake.drive', version: '2' });
    expect([store.unsaved, store.sourceUnsaved]).toEqual([false, false]);
    store.dispose();
  });

  it('without a host’s Save, writing to the source is saving', async () => {
    const drive = fakeDrive();
    const { store, handle } = await editor({ plugins: [drive.plugin] });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    await store.input(1, 1, '1500');
    expect(await handle.save()).toBe('saved');
    expect(drive.writes).toHaveLength(1);
    expect([store.unsaved, store.sourceUnsaved]).toEqual([false, false]);
    store.dispose();
  });

  it('keeps the changes and offers a copy when the file changed since it was opened', async () => {
    const drive = fakeDrive();
    const { store } = await editor({ plugins: [drive.plugin] });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    await store.input(1, 1, '1500');
    drive.bump();
    const saving = store.plugins!.saveToSource();
    expect(await answer(store, 'cancel')).toBe('This workbook was changed elsewhere');
    expect(await saving).toBe('failed');
    expect(drive.writes).toEqual([]);
    expect(store.origin?.version).toBe('1');
    expect([store.unsaved, store.sourceUnsaved]).toEqual([true, true]);
    store.dispose();
  });

  it('keeps the changes when the source fails, naming it', async () => {
    const drive = fakeDrive();
    drive.source.save = () => Promise.reject(new Error('Network down.'));
    const { store, events } = await editor({ plugins: [drive.plugin] });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    await store.input(1, 1, '1500');
    expect(await store.plugins!.saveToSource()).toBe('failed');
    expect(english.text(store.toasts.at(-1)!.text)).toBe('Not saved to Fake Drive: Network down. Your changes are still here.');
    expect(events.find((e) => e.type === 'pluginError')).toMatchObject({ plugin: 'test.plugin', message: 'Network down.' });
    expect(store.sourceUnsaved).toBe(true);
    store.plugins!.sources[0].save = async () => undefined as never;
    expect(await store.plugins!.saveToSource()).toBe('failed');
    store.dispose();
  });

  it('asks, but doesn’t stop, when the schema has errors', async () => {
    const drive = fakeDrive();
    const { store } = await editor({ plugins: [drive.plugin], schemas: [schemaFromJSON(JSON.stringify(forecastSchema))], onSave: () => undefined });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
    // The forecast sample breaks its schema on purpose: the host's Save is blocked.
    expect(await store.save()).toBe('blocked');
    const saving = store.plugins!.saveToSource();
    expect(await answer(store, 'cancel')).toBe('Save with errors?');
    expect(await saving).toBe('cancelled');
    const again = store.plugins!.saveToSource();
    await answer(store, 'confirm');
    expect(await again).toBe('saved');
    store.dispose();
  });

  it('saves copies to sources, leaving the workbook where it came from', async () => {
    const drive = fakeDrive({ writable: false });
    const { store, events } = await editor({ plugins: [drive.plugin] });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    await store.input(1, 1, '1500');
    expect(store.canSaveToSource).toBe(false);
    expect(await store.plugins!.saveToSource()).toBe('cancelled');
    expect(await store.savePrimary()).toBe('saved');
    expect(drive.copies).toHaveLength(1);
    expect(store.origin?.name).toBe('forecast.xlsx');
    expect(store.unsaved).toBe(false);
    expect(events.find((e) => e.type === 'savedToSource')).toMatchObject({ name: 'copy.xlsx' });

    store.plugins!.sources[0].saveAs = async () => null;
    expect(await store.plugins!.saveCopyTo(store.plugins!.sources[0])).toBe('cancelled');
    store.dispose();
  });

  it('writes nowhere when the host turns saving to sources off', async () => {
    const drive = fakeDrive();
    const { store } = await editor({ plugins: [drive.plugin], features: { sourceSave: false } });
    await store.plugins!.openFrom(store.plugins!.sources[0]);
    expect(store.copySources).toEqual([]);
    expect(await store.plugins!.saveToSource()).toBe('cancelled');
    expect(await store.plugins!.saveCopyTo(store.plugins!.sources[0])).toBe('cancelled');
    store.dispose();
  });

  it('opens files a plugin found some other way, and the editor reopens them from where they came', async () => {
    let ctx: PluginContext | undefined;
    const drive = fakeDrive();
    const { store } = await editor({
      plugins: [
        plugin((c) => {
          ctx = c;
          c.files.registerSource(drive.source);
        }),
      ],
    });
    expect(await ctx!.files.open({ bytes: FORECAST.slice().buffer as ArrayBuffer, origin: { source: 'fake.drive', name: 'recent.xlsx', ref: 1 } })).toBe(true);
    expect(store.fileName).toBe('recent.xlsx');
    expect(store.origin?.ref).toBe(1);
    expect(store.saveTarget).toBe('source');
    store.dispose();
  });

  it('tells plugins what is open, as opened, and where from', async () => {
    let ctx: PluginContext | undefined;
    const drive = fakeDrive();
    const { store, handle } = await editor({
      plugins: [
        plugin((c) => {
          ctx = c;
          c.files.registerSource(drive.source);
        }),
      ],
    });
    expect(ctx!.files.current()).toBeUndefined();

    await store.open('forecast.xlsx', FORECAST.slice().buffer as ArrayBuffer, { from: 'device' });
    const device = ctx!.files.current()!;
    expect(device).toMatchObject({ name: 'forecast.xlsx', from: 'device', size: FORECAST.byteLength, origin: undefined });
    // Each call is a copy, unchanged by edits.
    await store.input(1, 1, '1500');
    const bytes = device.bytes();
    expect(new Uint8Array(bytes)).toEqual(FORECAST);
    expect(device.bytes()).not.toBe(bytes);

    const opening = handle.open(FORECAST, 'from-host.xlsx');
    await answer(store, 'confirm');
    await opening;
    expect(ctx!.files.current()).toMatchObject({ name: 'from-host.xlsx', from: 'host' });

    await store.plugins!.openFrom(store.plugins!.sources[0]);
    expect(ctx!.files.current()).toMatchObject({ name: 'forecast.xlsx', from: 'source', origin: { source: 'fake.drive', ref: { item: 7 } } });
    store.dispose();
  });

  it('lets a source offer files by name, opened through the source each names', async () => {
    const drive = fakeDrive();
    drive.source.fetch = async (origin) => ({ bytes: FORECAST, origin });
    const offered: FileOrigin = { source: 'fake.drive', name: 'forecast.xlsx', ref: { item: 7 } };
    let items: () => unknown = () => [offered, null, { name: 'no source' }];
    const { store, events } = await editor({
      plugins: [drive.plugin, definePlugin({ id: 'test.list', apiVersion: 1, setup: (c) => void c.files.registerSource({ id: 'list', label: 'List', items: () => items() as FileOrigin[] }) })],
    });
    const list = store.plugins!.sources.find((s) => s.id === 'list')!;
    // Only origins are offered; a source with nothing to pick picks nothing.
    expect(store.plugins!.items(list)).toEqual([offered]);
    expect(await store.plugins!.openFrom(list)).toBe(false);
    expect(await store.plugins!.openItem(offered)).toBe(true);
    expect(store.origin).toMatchObject({ source: 'fake.drive', name: 'forecast.xlsx' });

    items = () => {
      throw new Error('No list.');
    };
    expect(store.plugins!.items(list)).toEqual([]);
    expect(events.at(-1)).toMatchObject({ type: 'pluginError', plugin: 'test.list', message: 'No list.' });
    items = () => 'not a list';
    expect(store.plugins!.items(list)).toEqual([]);
    store.dispose();
  });

  it('fetches a file again through the source its origin names, even another plugin’s', async () => {
    let ctx: PluginContext | undefined;
    const drive = fakeDrive();
    const fetched: FileOrigin[] = [];
    drive.source.fetch = async (origin) => {
      fetched.push(origin);
      if (origin.ref === 'gone') throw new Error('It was deleted.');
      return { bytes: FORECAST, origin: { ...origin, version: '9' } };
    };
    const { store, events } = await editor({ plugins: [drive.plugin, definePlugin({ id: 'test.recent', apiVersion: 1, setup: (c) => void (ctx = c) })] });
    const origin: FileOrigin = { source: 'fake.drive', name: 'forecast.xlsx', ref: { item: 7 }, version: '1' };

    const file = await ctx!.files.fetch(origin);
    expect(file?.origin.version).toBe('9');
    expect(fetched).toEqual([origin]);
    expect(await ctx!.files.open(file!)).toBe(true);
    expect(store.origin?.version).toBe('9');

    // A failing source is named; one that isn't here, or can't fetch, is explained.
    expect(await ctx!.files.fetch({ ...origin, ref: 'gone' })).toBeNull();
    expect(events.at(-1)).toMatchObject({ type: 'pluginError', plugin: 'test.plugin', message: 'It was deleted.' });
    expect(await ctx!.files.fetch({ ...origin, source: 'elsewhere' })).toBeNull();
    expect(english.text(store.toasts.at(-1)!.text)).toMatch(/forecast\.xlsx can't be opened again here/);
    delete drive.source.fetch;
    store.dispose();
  });
});
