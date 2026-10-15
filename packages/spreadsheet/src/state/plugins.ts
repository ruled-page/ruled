import type { EditorHandle } from '../editor/handle';
import { english } from '../i18n';
import { errorText, msg, type Text } from '../i18n/message';
import type { Command, CommandContext, FileOrigin, FileSource, JsonValue, PluginContext, PluginPanel, PluginStorage, PluginText, RemoteFile, RuledPlugin } from '../plugin/types';
import type { EditorEvent, SavedWorkbook, SaveResult, WorkbookStore } from './store';

/** A plugin's command, source or panel as the editor keeps it: under the plugin's id. */
type Registered<T> = T & { plugin: string; key: string };
export type RegisteredCommand = Registered<Command> & { broken?: boolean };
export type RegisteredSource = Registered<FileSource>;
export type RegisteredPanel = Registered<PluginPanel>;

/** The plugin contract versions this editor knows. */
const API_VERSIONS: readonly unknown[] = [1];

/** The plugin's text as the editor's: its messages are looked up like the editor's own. */
export const pluginText = (text: PluginText) => text as Text;

/** Values for one plugin, in localStorage; storage being unavailable (a private window) is not an error. */
function pluginStorage(id: string): PluginStorage {
  const name = `ruled-plugin:${id}`;
  const read = (): Record<string, JsonValue> => {
    try {
      const data: unknown = JSON.parse(localStorage.getItem(name) ?? '{}');
      return data && typeof data === 'object' && !Array.isArray(data) ? (data as Record<string, JsonValue>) : {};
    } catch {
      return {};
    }
  };
  const write = (change: (data: Record<string, JsonValue>) => void) => {
    const data = read();
    change(data);
    try {
      localStorage.setItem(name, JSON.stringify(data));
    } catch {
      // See above.
    }
  };
  return {
    get: async (key) => (Object.hasOwn(read(), key) ? read()[key] : undefined),
    set: async (key, value) => write((d) => void (d[key] = value)),
    delete: async (key) => write((d) => void delete d[key]),
  };
}

const isConflict = (err: unknown) => (err as Error | undefined)?.name === 'SaveConflictError';

/**
 * The plugins of one editor: what they registered, every call into them,
 * each wrapped so that a plugin that throws is reported and never takes the
 * editor down, and opening from and saving to their file sources. Loaded
 * only for an editor given plugins.
 */
export class PluginHost {
  commands: RegisteredCommand[] = [];
  sources: RegisteredSource[] = [];
  panels: RegisteredPanel[] = [];
  /** Plugins that can be started, by id. */
  private readonly accepted = new Map<string, RuledPlugin>();
  /** Plugins refused, and why, said once the editor is made. */
  private readonly refused: { plugin: string; reason: string }[] = [];
  private readonly running = new Map<string, { abort: AbortController; cleanup?: () => void }>();
  private readonly listeners = new Set<{ plugin: string; fn: (event: EditorEvent) => void }>();
  /** Set while plugins hear an event, or a plugin's failure is reported: events then aren't passed to plugins. */
  private dispatching = false;

  constructor(
    private readonly store: WorkbookStore,
    plugins: readonly RuledPlugin[] = [],
  ) {
    for (const p of plugins) {
      const id = typeof p?.id === 'string' && p.id ? p.id : undefined;
      // Said in English: these are mistakes in the host's code, for its developers to read.
      const refuse = (reason: string) => this.refused.push({ plugin: id ?? String(p), reason });
      if (!id || typeof p.setup !== 'function') refuse('a plugin needs an id and a setup function.');
      else if (!API_VERSIONS.includes(p.apiVersion)) refuse(`it was written for version ${String(p.apiVersion)} of the plugin contract, which this editor doesn't know, so it wasn't started.`);
      else if (this.accepted.has(id)) refuse('another plugin has the same id, so only the first was started.');
      else this.accepted.set(id, p);
    }
  }

  /** Whether any plugin is running, so the editor loads the parts that show what they add. */
  get active(): boolean {
    return this.running.size > 0;
  }

  /** Starts each plugin, once the editor's handle is made. */
  start(editor: EditorHandle): void {
    for (const { plugin, reason } of this.refused) this.fail(plugin, new Error(reason));
    for (const [id, plugin] of this.accepted) {
      const abort = new AbortController();
      const entry: { abort: AbortController; cleanup?: () => void } = { abort };
      this.running.set(id, entry);
      try {
        const result: unknown = plugin.setup(this.context(id, editor, abort.signal));
        if (typeof result === 'function') entry.cleanup = result as () => void;
        // An async setup that rejects ends the plugin as one that throws does.
        else if (result instanceof Promise) result.catch((err: unknown) => this.fail(id, err, true));
      } catch (err) {
        this.fail(id, err, true);
      }
    }
    this.store.notify();
  }

  /** Ends every plugin, when the editor is destroyed. */
  stop(): void {
    for (const id of this.running.keys()) this.stopPlugin(id);
  }

  /**
   * Calls into a plugin: a throw is reported, naming the plugin, and gives
   * `fallback`; a promise it returns is reported if it rejects.
   */
  call<T, F>(plugin: string, fn: () => T, fallback: F): T | F {
    try {
      const result = fn();
      if (result instanceof Promise) result.catch((err: unknown) => this.fail(plugin, err));
      return result;
    } catch (err) {
      this.fail(plugin, err);
      return fallback;
    }
  }

  /** Reports a plugin's failure; `stop` also ends it, removing what it added. */
  fail(plugin: string, err: unknown, stop = false, text?: Text): void {
    this.report(plugin, text ?? msg('store.plugin.failed', { plugin, reason: errorText(err) }), errorText(err));
    if (stop) this.stopPlugin(plugin);
  }

  /** Shows the problem in the editor and tells the host's onEvent; plugins don't hear it, so one that throws on hearing it can't go round for ever. */
  private report(plugin: string, text: Text, reason: Text = text): void {
    const was = this.dispatching;
    this.dispatching = true;
    try {
      this.store.toast('error', text);
      this.store.emit({ type: 'pluginError', plugin, message: english.text(reason) });
    } finally {
      this.dispatching = was;
    }
  }

  private stopPlugin(id: string): void {
    const entry = this.running.get(id);
    if (!entry) return;
    this.running.delete(id);
    entry.abort.abort();
    if (entry.cleanup) this.call(id, entry.cleanup, undefined);
    this.commands = this.commands.filter((c) => c.plugin !== id);
    this.sources = this.sources.filter((s) => s.plugin !== id);
    this.panels = this.panels.filter((p) => p.plugin !== id);
    for (const l of this.listeners) if (l.plugin === id) this.listeners.delete(l);
    if (this.store.panel?.startsWith(`plugin:${id}/`)) this.store.panel = null;
    this.store.notify();
  }

  /** Tells the plugins what the editor did. */
  dispatch(event: EditorEvent): void {
    if (this.dispatching) return;
    this.dispatching = true;
    try {
      for (const l of this.listeners) this.call(l.plugin, () => l.fn(event), undefined);
    } finally {
      this.dispatching = false;
    }
  }

  /** A command's state for the selection: hidden, or shown and whether it can run. One that throws is hidden from then on. */
  commandState(command: RegisteredCommand, context: CommandContext): { visible: boolean; enabled: boolean } {
    if (command.broken) return { visible: false, enabled: false };
    const visible = command.visible ? this.call(command.plugin, () => command.visible!(context) !== false, 'broken' as const) : true;
    const enabled = visible === true && command.enabled ? this.call(command.plugin, () => command.enabled!(context) !== false, 'broken' as const) : true;
    if (visible === 'broken' || enabled === 'broken') {
      command.broken = true;
      return { visible: false, enabled: false };
    }
    return { visible, enabled };
  }

  /** Runs a command, inside the click that chose it. */
  run(command: RegisteredCommand, context: CommandContext): void {
    this.call(command.plugin, () => command.run(context), undefined);
  }

  /** Lets plugins load what they need, as a menu holding their commands or sources opens. */
  preload(items: readonly (RegisteredCommand | RegisteredSource)[]): void {
    for (const item of items) if (item.preload) this.call(item.plugin, () => item.preload!(), undefined);
  }

  /** The source a workbook's origin names. */
  source(id: string | undefined): RegisteredSource | undefined {
    return this.sources.find((s) => s.id === id);
  }

  // ---------------------------------------------------------------------------
  // Opening from and saving to sources

  /**
   * Opens a file a plugin's source fetched, keeping where it came from:
   * whether it opened. The editor has already said why when it didn't.
   */
  async openRemote(file: RemoteFile): Promise<boolean> {
    const store = this.store;
    const origin = file?.origin;
    if (!origin || typeof origin.name !== 'string' || !(file.bytes instanceof ArrayBuffer || ArrayBuffer.isView(file.bytes))) {
      store.toast('error', msg('store.open.badRemote'));
      return false;
    }
    const bytes = file.bytes instanceof ArrayBuffer ? file.bytes.slice(0) : (new Uint8Array(file.bytes.buffer, file.bytes.byteOffset, file.bytes.byteLength).slice().buffer as ArrayBuffer);
    return (await store.open(origin.name, bytes, { csv: store.features.csv, origin: { ...origin } })).opened;
  }

  /**
   * The Open menu's entry for a source: lets the person pick a file and opens
   * it. The source is asked inside the click, so its popup isn't blocked;
   * unsaved changes are asked about after, once there is a file to open.
   */
  async openFrom(source: RegisteredSource): Promise<boolean> {
    let file: RemoteFile | null;
    try {
      file = await source.pick();
    } catch (err) {
      this.fail(source.plugin, err);
      return false;
    }
    return file ? this.openRemote(file) : false;
  }

  /**
   * "Save to <source>": writes the workbook back over the file it was opened
   * from. A schema's errors don't stop it, as the file is the person's own,
   * but they are asked first. On success the source's new origin replaces
   * the old, so the next save is checked against the version just written.
   */
  async saveToSource(): Promise<SaveResult> {
    const store = this.store;
    const source = store.originSource;
    const origin = store.origin;
    if (!source?.save || !origin || !store.canSaveToSource) return 'cancelled';
    if (store.schema) {
      await store.schema.settled();
      const { errors } = store.schema.summary();
      if (
        errors &&
        !(await store.confirm({
          title: msg('store.source.withErrorsTitle'),
          body: msg('store.source.withErrorsBody', { count: errors, source: store.sourceLabel(source) }),
          confirmLabel: msg('store.source.withErrorsConfirm'),
          danger: false,
        }))
      )
        return 'cancelled';
    }
    return this.write(source, async (saved) => {
      const next = await source.save!(saved, origin);
      if (!next || typeof next !== 'object' || typeof next.name !== 'string') throw new Error('save() did not resolve with the file’s origin.');
      store.savedToSource({ ...next });
      return { origin: next, text: msg('store.source.saved', { fileName: next.name, source: store.sourceLabel(source) }) };
    });
  }

  /** "Save a copy" to a source: a new file, where the person chooses. The open workbook stays where it came from. */
  async saveCopyTo(source: RegisteredSource): Promise<SaveResult> {
    if (!source.saveAs || !this.store.copySources.includes(source)) return 'cancelled';
    return this.write(source, async (saved) => {
      const made = await source.saveAs!(saved);
      if (!made) return undefined;
      return { origin: made, text: msg('store.source.copied', { fileName: made.name, source: this.store.sourceLabel(source) }) };
    });
  }

  /** Writes the workbook through a source: a conflict is handled as one from onSave, and a failure keeps the changes. */
  private async write(source: RegisteredSource, write: (saved: SavedWorkbook) => Promise<{ origin: FileOrigin; text: Text } | undefined>): Promise<SaveResult> {
    const store = this.store;
    const started = performance.now();
    const saved = await store.snapshot();
    if (!saved) return 'failed';
    store.busy = true;
    store.notify();
    try {
      const done = await write(saved);
      if (!done) return 'cancelled';
      // Without a host's Save, the source is where the workbook is kept.
      if (!store.hasSaveHandler) store.afterSave();
      store.toast('info', done.text);
      store.emit({ type: 'savedToSource', source: source.id, name: done.origin.name, ms: Math.round(performance.now() - started) });
      return 'saved';
    } catch (err) {
      if (isConflict(err)) await store.conflict(err as Error, saved);
      else this.fail(source.plugin, err, false, msg('store.source.notSaved', { source: store.sourceLabel(source), reason: errorText(err) }));
      return 'failed';
    } finally {
      store.busy = false;
      store.notify();
    }
  }

  private context(plugin: string, editor: EditorHandle, signal: AbortSignal): PluginContext {
    const store = this.store;
    // What a plugin adds, until it removes it or is stopped.
    const add = <K extends 'commands' | 'sources' | 'panels'>(list: K, item: PluginHost[K][number]) => {
      if (signal.aborted) return () => undefined;
      (this[list] as PluginHost[K][number][]) = [...this[list], item];
      store.notify();
      return () => {
        (this[list] as PluginHost[K][number][]) = this[list].filter((x) => x !== item);
        store.notify();
      };
    };
    const own = (list: 'commands' | 'sources' | 'panels', id: string) => {
      if (typeof id !== 'string' || !id) throw new Error(`Every one of a plugin's ${list} needs an id.`);
      const key = `${plugin}/${id}`;
      if (this[list].some((x) => x.key === key) || (list === 'sources' && this.source(id))) throw new Error(`"${id}" is already registered.`);
      return key;
    };
    return {
      editor,
      signal,
      commands: { register: (command) => add('commands', { ...command, plugin, key: own('commands', command.id) }) },
      files: {
        registerSource: (source) => add('sources', { ...source, plugin, key: own('sources', source.id) }),
        open: (file) => this.openRemote(file),
      },
      ui: {
        toast: (kind, text) => store.toast(kind, pluginText(text)),
        confirm: ({ title, body, confirmLabel }) => store.confirm({ title: pluginText(title), body: pluginText(body), confirmLabel: pluginText(confirmLabel), danger: false }),
        dialog: (options) => store.showPluginDialog(plugin, options),
        registerPanel: (panel) => add('panels', { ...panel, plugin, key: own('panels', panel.id) }),
      },
      on: (fn) => {
        const entry = { plugin, fn };
        this.listeners.add(entry);
        return () => void this.listeners.delete(entry);
      },
      storage: pluginStorage(plugin),
    };
  }
}
