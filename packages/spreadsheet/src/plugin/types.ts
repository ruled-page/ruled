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
  commands: PluginCommands;
  files: PluginFiles;
  ui: PluginUI;
  /** What the editor does, as `onEvent` hears it. Returns an unsubscribe function. */
  on(listener: (event: EditorEvent) => void): () => void;
  /** A small key–value store kept in this browser for this plugin (recent files, the last folder). Never workbook contents. */
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
  /** Read again whenever the editor's state changes. */
  enabled?(context: CommandContext): boolean;
  visible?(context: CommandContext): boolean;
  /** Called when the menu holding it opens, or the pointer rests on it: load an SDK here. */
  preload?(): void;
}

export interface PluginCommands {
  /** Adds a command; returns a function that removes it. */
  register(command: Command): () => void;
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
}

export interface PluginFiles {
  /** Adds a source to the Open and "Save a copy" menus; returns a function that removes it. */
  registerSource(source: FileSource): () => void;
  /** Opens a file as if picked from one of this plugin's sources, for files found some other way (a recent-files list, a link). */
  open(file: RemoteFile): Promise<boolean>;
}

export interface DialogAction {
  id: string;
  label: PluginText;
  /** The dialog's main action, styled as such. */
  primary?: boolean;
}

export interface PluginDialog {
  title: PluginText;
  /** Fills the dialog's body with plain DOM; may return a function to clean up when it closes. */
  render(body: HTMLElement): void | (() => void);
  /** Buttons under the body; without them, the dialog has a Close button. */
  actions?: DialogAction[];
}

export interface PluginPanel {
  /** Unique within the plugin. */
  id: string;
  title: PluginText;
  icon?: IconSource;
  /** Fills the panel with plain DOM; may return a function to clean up when it closes. */
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

/** Values kept in this browser for the plugin. Storage may be unavailable (a private window): then nothing is kept, without an error. */
export interface PluginStorage {
  get(key: string): Promise<JsonValue | undefined>;
  set(key: string, value: JsonValue): Promise<void>;
  delete(key: string): Promise<void>;
}
