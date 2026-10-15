import type { FileOrigin, PluginContext, PluginText, RuledPlugin } from '@ruled/sheet/plugin';

/**
 * A "Recent files" plugin: the workbooks opened lately, to open again in a
 * click. It is a file source whose files are a submenu of Open, a panel
 * listing them, and a More command that clears the list.
 *
 * Every file is kept as its origin. Files from a source (OneDrive) are
 * fetched again through that source, so they open as they are now and save
 * back to it. Files from this device are kept with their bytes too, as a
 * browser can't open a local file again by itself, and this source fetches
 * them; only small ones, since they are kept in localStorage. Files the host
 * opened are its own business, so they aren't listed.
 */

export const RECENT_FILES_ID = 'com.example.recent-files';
const SOURCE_ID = 'recent';

export interface RecentFilesOptions {
  /** How many files to list. Default 8. */
  max?: number;
  /** The largest file from this device to keep, in bytes. Larger ones aren't listed. Default 100 kB. */
  keepBytesUpTo?: number;
}

/** A file in the list. */
export interface RecentFile {
  /** Its source and identity there, or its name for a file from this device. */
  id: string;
  /** When it was last opened. */
  openedAt: number;
  /** Where it is kept: its source's origin, or this plugin's for a file from this device. */
  origin: FileOrigin;
  /** Its bytes, as base64, for a file from this device. */
  bytes?: string;
}

export const translations = {
  en: {
    title: 'Recent files',
    empty: 'Workbooks you open from this device or a file source appear here.',
    thisDevice: 'This device',
    forget: 'Remove {{name}} from the list',
    clear: 'Clear recent files',
    clearBody: 'The list is kept in this browser only. The files themselves are left as they are.',
    clearConfirm: 'Clear',
    gone: '{{name}} is no longer kept in this browser. Open it again from this device.',
  },
  de: {
    title: 'Zuletzt geöffnet',
    empty: 'Arbeitsmappen, die Sie von diesem Gerät oder aus einer Dateiquelle öffnen, erscheinen hier.',
    thisDevice: 'Dieses Gerät',
    forget: '{{name}} aus der Liste entfernen',
    clear: 'Liste leeren',
    clearBody: 'Die Liste wird nur in diesem Browser gespeichert. Die Dateien selbst bleiben unverändert.',
    clearConfirm: 'Leeren',
    gone: '{{name}} ist in diesem Browser nicht mehr gespeichert. Öffnen Sie die Datei erneut von diesem Gerät.',
  },
};

type Key = keyof (typeof translations)['en'];

/** One of this plugin's messages, for the editor to show in its language. */
const msg = (key: Key, params?: Record<string, unknown>): PluginText => ({ key: `plugin.${RECENT_FILES_ID}.${key}`, params });

/** A clock, on the editor's 24×24 icon grid. */
const ICON = 'M12 7v5l3 2M3 12a9 9 0 1 0 18 0a9 9 0 1 0-18 0';

export function recentFiles({ max = 8, keepBytesUpTo = 100 * 1024 }: RecentFilesOptions = {}): RuledPlugin {
  return {
    id: RECENT_FILES_ID,
    apiVersion: 1,
    translations,
    setup: (ctx) => setup(ctx, max, keepBytesUpTo),
  };
}

function setup(ctx: PluginContext, max: number, keepBytesUpTo: number): () => void {
  let list: RecentFile[] = [];
  /** Panels showing the list, drawn again when it changes. */
  const views = new Set<() => void>();
  const changed = async (next: RecentFile[]) => {
    list = next;
    await ctx.storage.set('files', list);
    for (const draw of views) draw();
    ctx.commands.refresh();
  };
  const loaded = ctx.storage.get<RecentFile[]>('files').then((kept) => {
    if (Array.isArray(kept)) list = kept.filter((f) => f?.origin);
    for (const draw of views) draw();
    ctx.commands.refresh();
  });

  /** Puts the workbook just opened at the top of the list. */
  async function remember() {
    const file = ctx.files.current();
    if (!file || file.from === 'host') return;
    await loaded;
    let entry: RecentFile | undefined;
    if (file.origin?.source === SOURCE_ID) {
      // Opened from this list: it is already kept.
      entry = list.find((f) => f.id === file.origin!.ref);
    } else if (file.origin) {
      entry = { id: `${file.origin.source}:${JSON.stringify(file.origin.ref)}`, openedAt: 0, origin: file.origin };
    } else if (file.size <= keepBytesUpTo) {
      const id = `device:${file.name}`;
      entry = { id, openedAt: 0, origin: { source: SOURCE_ID, name: file.name, ref: id }, bytes: toBase64(file.bytes()) };
    }
    if (!entry) return;
    await changed([{ ...entry, openedAt: Date.now() }, ...list.filter((f) => f.id !== entry.id)].slice(0, max));
  }

  const fromDevice = (origin: FileOrigin) => origin.source === SOURCE_ID;
  /** The origin as shown: a file from this device says so. */
  const shown = (origin: FileOrigin): FileOrigin => (fromDevice(origin) ? { ...origin, location: ctx.text(msg('thisDevice')) } : origin);

  /** Opens a listed file: fetched again through the source its origin names, which for this device's files is this one. */
  async function reopen(entry: RecentFile) {
    const file = await ctx.files.fetch(entry.origin);
    if (file) await ctx.files.open(file);
  }

  /** The list as buttons that open the files, each with a button that removes it. */
  function drawList(body: HTMLElement) {
    if (!list.length) {
      const empty = document.createElement('p');
      empty.textContent = ctx.text(msg('empty'));
      empty.style.cssText = 'margin: 12px; color: var(--wbe-text-muted)';
      body.replaceChildren(empty);
      return;
    }
    const ul = document.createElement('ul');
    ul.style.cssText = 'list-style: none; margin: 0; padding: 8px; display: grid; gap: 4px';
    for (const entry of list) {
      const li = document.createElement('li');
      li.style.cssText = 'display: flex; gap: 4px; align-items: center';
      const open = document.createElement('button');
      open.className = 'wbe-btn';
      open.style.cssText = 'flex: 1; min-width: 0; height: auto; padding-block: 6px; flex-direction: column; align-items: flex-start; text-align: start';
      const name = document.createElement('span');
      name.textContent = entry.origin.name;
      const where = document.createElement('small');
      const origin = shown(entry.origin);
      where.textContent = origin.location ?? origin.source;
      where.style.cssText = 'color: var(--wbe-text-subtle); font-size: var(--wbe-font-size-xs)';
      open.append(name, where);
      open.addEventListener('click', () => void reopen(entry));
      const remove = document.createElement('button');
      remove.className = 'wbe-btn';
      remove.textContent = '×';
      remove.setAttribute('aria-label', ctx.text(msg('forget', { name: entry.origin.name })));
      remove.title = remove.getAttribute('aria-label')!;
      remove.addEventListener('click', () => void changed(list.filter((f) => f !== entry)));
      li.append(open, remove);
      ul.append(li);
    }
    body.replaceChildren(ul);
  }

  const stop = ctx.on((event) => {
    if (event.type === 'opened') void remember();
  });

  ctx.files.registerSource({
    id: SOURCE_ID,
    label: msg('title'),
    icon: ICON,
    // A submenu of Open; choosing one fetches it through the source its origin names.
    items: () => list.map((f) => shown(f.origin)),
    // The files from this device, from the bytes kept. This source has no `save`, so saving downloads them as before.
    async fetch(origin) {
      const entry = list.find((f) => f.id === origin.ref);
      if (!entry?.bytes) {
        ctx.ui.toast('error', msg('gone', { name: origin.name }));
        return null;
      }
      return { bytes: fromBase64(entry.bytes), origin: shown(entry.origin) };
    },
  });

  ctx.ui.registerPanel({
    id: 'list',
    title: msg('title'),
    icon: ICON,
    // Drawn again whenever the list changes, until the panel closes.
    render: (body) => {
      const draw = () => drawList(body);
      draw();
      views.add(draw);
      return () => void views.delete(draw);
    },
  });

  ctx.commands.register({
    id: 'clear',
    label: msg('clear'),
    where: 'more',
    enabled: () => list.length > 0,
    async run() {
      if (await ctx.ui.confirm({ title: msg('clear'), body: msg('clearBody'), confirmLabel: msg('clearConfirm') })) await changed([]);
    },
  });

  return () => {
    stop();
    views.clear();
  };
}

function toBase64(bytes: ArrayBuffer): string {
  const all = new Uint8Array(bytes);
  let text = '';
  for (let i = 0; i < all.length; i += 0x8000) text += String.fromCharCode(...all.subarray(i, i + 0x8000));
  return btoa(text);
}

function fromBase64(text: string): Uint8Array {
  return Uint8Array.from(atob(text), (c) => c.charCodeAt(0));
}
