import type { EditorHandle } from '../editor/handle';
import type { Translations } from '../i18n';
import type { LocalizedMap } from '../i18n/message';
import type { EditorEvent, SavedWorkbook } from '../state/store';

/**
 * Text a plugin shows: written text, host-style text in several languages
 * ({ en: 'Recent', de: 'Zuletzt' }), or one of its own messages to translate,
 * `{ key: 'plugin.<id>.<key>', params }`, from its `translations`.
 */
export type PluginText = string | LocalizedMap | { key: string; params?: Record<string, unknown> };

/** An icon: SVG path data drawn on a 24×24 grid with the current text colour, as the `icons` option takes. */
export type IconSource = string;

export type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue };

/**
 * A plugin: a plain object a host passes in `plugins`. It needs nothing from
 * `@ruled/sheet` at run time, only its types.
 */
export interface RuledPlugin {
  /** Unique, reverse-DNS style: "page.ruled.onedrive". Commands, messages and storage are kept under it. */
  id: string;
  /** The contract this plugin was written for. The editor refuses a version it doesn't know, saying so. */
  apiVersion: 1;
  /** Messages by language, keyed as the plugin uses them; the editor shows them as `plugin.<id>.<key>`. */
  translations?: Translations;
  /** Called once, as soon as the editor is made and the code that runs plugins has loaded. Return a function to clean up, or nothing. */
  setup(ctx: PluginContext): void | (() => void);
}

export interface PluginContext {
  /** The same handle the host has: open, save, getWorkbook, subscribe, status, origin… */
  editor: EditorHandle;
  /** Aborted when the editor is destroyed, or the plugin is stopped after failing; pass it to fetches and listeners. */
  signal: AbortSignal;
  /** The editor's limits, so a source can refuse a file it won't open before downloading it. */
  limits: { maxFileBytes: number };
  /** Text in the editor's language as it is now, for the plugin's own DOM in dialogs and panels. */
  text(text: PluginText): string;
  commands: PluginCommands;
  files: PluginFiles;
  ui: PluginUI;
  /** What the editor does, as `onEvent` hears it. Returns an unsubscribe function. */
  on(listener: (event: EditorEvent) => void): () => void;
  /** A small key–value store kept in this browser for this plugin: recent files, the last folder. */
  storage: PluginStorage;
}

/** The selection a cell-menu command is for: the sheet's name, the active cell and the selected ranges, in A1 style. */
export interface CommandSelection {
  sheet: string;
  active: string;
  ranges: string[];
}

export interface CommandContext {
  /** Given to commands in the cell menu. */
  selection?: CommandSelection;
}

/** An action with a label, shown on the toolbar, in More, or in the cell menu. */
export interface Command {
  /** Unique within the plugin. */
  id: string;
  label: PluginText;
  icon?: IconSource;
  where: 'toolbar' | 'more' | 'cellMenu';
  /** Runs inside the click, so a popup it opens straight away isn't blocked. */
  run(context: CommandContext): void | Promise<void>;
  /** Read again whenever the editor's state changes, or the plugin calls `commands.refresh()`. */
  enabled?(context: CommandContext): boolean;
  visible?(context: CommandContext): boolean;
  /** Called when the menu holding it opens, or the pointer rests on it: load an SDK here. */
  preload?(): void;
}

export interface PluginCommands {
  /** Adds a command; returns a function that removes it. */
  register(command: Command): () => void;
  /** Reads every command's `enabled` and `visible` again, after something they depend on changed outside the editor. */
  refresh(): void;
}

/** Where the open workbook came from. Kept by the editor, handed back on save, never written to the file. */
export interface FileOrigin {
  /** The id of the FileSource it came from. */
  source: string;
  /** Shown as the file's name. */
  name: string;
  /** The source's own identity for the file, such as a drive and item id. Plain JSON. */
  ref: JsonValue;
  /** Its version as opened (an ETag, a revision id), for telling a stale save. */
  version?: string;
  /** False when saving back isn't possible or wise; only "Save a copy" is offered then. */
  writable?: boolean;
  /** Where the file is, such as "Shared/Finance/prices.xlsx", shown when the pointer rests on its name. */
  location?: string;
}

/** A file a source fetched, with where it came from. */
export interface RemoteFile {
  bytes: ArrayBuffer | Uint8Array;
  origin: FileOrigin;
}

/** Somewhere files are opened from and saved to, such as a cloud drive. */
export interface FileSource {
  /** Unique within the plugin; FileOrigin.source names it. */
  id: string;
  /** Shown in the Open menu, as "Save to <label>" and in "Save a copy". */
  label: PluginText;
  icon?: IconSource;
  /** Called when a menu holding the source opens: load an SDK here, so `pick` can open a popup inside the click. */
  preload?(): void;
  /**
   * Lets the person choose a file, and fetches it. Resolves null when they
   * cancel. An error thrown here, or from `save` or `saveAs`, is shown with
   * its `text` when it has one (any PluginText, so it can be translated),
   * otherwise its message.
   */
  pick(): Promise<RemoteFile | null>;
  /**
   * Writes the workbook back over the file it was opened from, resolving
   * with its origin as now saved. Throw an error named "SaveConflictError"
   * (SaveConflictError from `@ruled/sheet` is one) when it changed since.
   */
  save?(file: SavedWorkbook, origin: FileOrigin): Promise<FileOrigin>;
  /** Writes the workbook as a new file, letting the person choose where and what to call it. Resolves null when they cancel. */
  saveAs?(file: SavedWorkbook): Promise<FileOrigin | null>;
  /**
   * Fetches the file an origin names again, as it is now, for opening it
   * without `pick`: from a list of recent files, or a link. The origin may
   * be old, so resolve with the file's current one. Resolves null when the
   * person cancels (a sign-in they close). Without it, files from this
   * source can only be opened by picking them.
   */
  fetch?(origin: FileOrigin): Promise<RemoteFile | null>;
}

/** The workbook open in the editor, as it was opened. */
export interface OpenedFile {
  /** Its file name, as shown. */
  name: string;
  /**
   * Where it came from: this device (the Open button, or a file dropped on
   * the editor), the host's code (its `file` option, or `open` on the handle),
   * or a plugin's source, which `origin` names.
   */
  from: 'device' | 'host' | 'source';
  /** Where it is kept, for a file from a source. */
  origin?: FileOrigin;
  /** Its size in bytes, as opened. */
  size: number;
  /** A copy of its bytes as opened, before any edits. */
  bytes(): ArrayBuffer;
}

export interface PluginFiles {
  /** Adds a source to the Open and "Save a copy" menus; returns a function that removes it. */
  registerSource(source: FileSource): () => void;
  /** Opens a file as if picked from a source, for files found some other way (a recent-files list, a link). Resolves whether it opened. */
  open(file: RemoteFile): Promise<boolean>;
  /**
   * Fetches a file again through the source its origin names, which may be
   * another plugin's, so `open` can open it. Resolves null when the person
   * cancels, or when it can't: the source isn't in this editor, has no
   * `fetch`, or failed, which the editor has then said.
   */
  fetch(origin: FileOrigin): Promise<RemoteFile | null>;
  /** The workbook open now, as it was opened; undefined while none is. Read it when the `opened` event comes, to know what was opened and from where. */
  current(): OpenedFile | undefined;
}

export interface DialogAction {
  id: string;
  label: PluginText;
  /** The dialog's main action, styled as such. */
  primary?: boolean;
}

export interface PluginDialog {
  title: PluginText;
  /**
   * Fills the dialog's body with plain DOM; may return a function to clean
   * up when it closes. `close` ends the dialog from inside the body, as
   * choosing an item in a list does, resolving `dialog` with that id. It is
   * called again, after cleaning up, when the editor's language changes.
   */
  render(body: HTMLElement, close: (action: string | null) => void): void | (() => void);
  /** Buttons under the body; without them, the dialog has a Close button. */
  actions?: DialogAction[];
}

export interface PluginPanel {
  /** Unique within the plugin. */
  id: string;
  title: PluginText;
  icon?: IconSource;
  /**
   * Fills the panel with plain DOM; may return a function to clean up when
   * it closes. It is called again, after cleaning up, when the editor's
   * language changes. To show new content meanwhile, change the body's DOM.
   */
  render(body: HTMLElement): void | (() => void);
}

/**
 * The editor's own messages and frames, in its style and language. Plugin
 * DOM picks up the theme through its CSS custom properties (--wbe-*), and
 * the `wbe-btn` and `wbe-input` classes.
 */
export interface PluginUI {
  toast(kind: 'info' | 'error', text: PluginText): void;
  confirm(options: { title: PluginText; body: PluginText; confirmLabel: PluginText }): Promise<boolean>;
  /** A dialog in the editor's frame, with its focus handling. Resolves with the action chosen, or null when it is closed. */
  dialog(options: PluginDialog): Promise<string | null>;
  /** A side panel beside Issues, Changes and the report, with a button to show it. Returns a function that removes it. */
  registerPanel(panel: PluginPanel): () => void;
}

/**
 * Values kept in this browser for the plugin, as JSON in `localStorage`,
 * which a page shares and browsers cap at about 5 MB: keep them small.
 * Storage may be unavailable or full (a private window): then nothing is
 * kept, without an error, so read back what you rely on.
 */
export interface PluginStorage {
  /** A value as set; `T` says what you kept there, unchecked. */
  get<T = JsonValue>(key: string): Promise<T | undefined>;
  /** Keeps a value JSON can hold: a list of objects, a string. */
  set<T>(key: string, value: T): Promise<void>;
  delete(key: string): Promise<void>;
}
