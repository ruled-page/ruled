import type { FileOrigin, PluginContext, PluginText, RemoteFile, RuledPlugin } from '@ruled/sheet/plugin';

/**
 * A "Recent files" plugin: the workbooks opened lately, to open again in a
 * click. It is a file source (Recent files, in the Open menu), a panel
 * listing them, and a More command that clears the list.
 *
 * Files from a source (OneDrive) are kept as their origin and fetched again
 * through that source, so they open as they are now and save back to it.
 * Files from this device are kept with their bytes, as a browser can't open
 * a local file again by itself; only small ones, since they are kept in
 * localStorage. Files the host opened are its own business, so they aren't
 * listed.
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
  name: string;
  /** When it was last opened. */
  openedAt: number;
  /** Where it is kept, for a file from a source. */
  origin?: FileOrigin;
  /** Its bytes, as base64, for a file from this device. */
  bytes?: string;
}

export const translations = {
  en: {
    title: 'Recent files',
    choose: 'Open a recent file',
    empty: 'Workbooks you open from this device or a file source appear here.',
    thisDevice: 'This device',
    forget: 'Remove {{name}} from the list',
    clear: 'Clear recent files',
    clearBody: 'The list is kept in this browser only. The files themselves are left as they are.',
    clearConfirm: 'Clear',
  },
  de: {
    title: 'Zuletzt geöffnet',
    choose: 'Zuletzt geöffnete Datei öffnen',
    empty: 'Arbeitsmappen, die Sie von diesem Gerät oder aus einer Dateiquelle öffnen, erscheinen hier.',
    thisDevice: 'Dieses Gerät',
    forget: '{{name}} aus der Liste entfernen',
    clear: 'Liste leeren',
    clearBody: 'Die Liste wird nur in diesem Browser gespeichert. Die Dateien selbst bleiben unverändert.',
    clearConfirm: 'Leeren',
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
  /** Panels and dialogs showing the list, drawn again when it changes. */
  const views = new Set<() => void>();
  const changed = async (next: RecentFile[]) => {
    list = next;
    await ctx.storage.set('files', list);
    for (const draw of views) draw();
    ctx.commands.refresh();
  };
  const loaded = ctx.storage.get<RecentFile[]>('files').then((kept) => {
    if (Array.isArray(kept)) list = kept;
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
      entry = { id: `${file.origin.source}:${JSON.stringify(file.origin.ref)}`, name: file.name, openedAt: 0, origin: file.origin };
    } else if (file.size <= keepBytesUpTo) {
      entry = { id: `device:${file.name}`, name: file.name, openedAt: 0, bytes: toBase64(file.bytes()) };
    }
    if (!entry) return;
    await changed([{ ...entry, openedAt: Date.now() }, ...list.filter((f) => f.id !== entry.id)].slice(0, max));
  }

  /** A listed file, ready to open: fetched again from its source, or from the bytes kept. */
  async function load(entry: RecentFile): Promise<RemoteFile | null> {
    if (entry.origin) return ctx.files.fetch(entry.origin);
    // Kept here, so the editor shows where it is; this source has no `save`, so saving downloads it as before.
    const origin: FileOrigin = { source: SOURCE_ID, name: entry.name, ref: entry.id, location: ctx.text(msg('thisDevice')) };
    return { bytes: fromBase64(entry.bytes ?? ''), origin };
  }

  async function reopen(entry: RecentFile) {
    const file = await load(entry);
    if (file) await ctx.files.open(file);
  }

  /** The list as buttons; `choose` runs when one is clicked, and with `forget` each has a button that removes it. */
  function drawList(body: HTMLElement, choose: (entry: RecentFile) => void, forget = false) {
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
      name.textContent = entry.name;
      const where = document.createElement('small');
      where.textContent = entry.origin ? (entry.origin.location ?? entry.origin.source) : ctx.text(msg('thisDevice'));
      where.style.cssText = 'color: var(--wbe-text-subtle); font-size: var(--wbe-font-size-xs)';
      open.append(name, where);
      open.addEventListener('click', () => choose(entry));
      li.append(open);
      if (forget) {
        const remove = document.createElement('button');
        remove.className = 'wbe-btn';
        remove.textContent = '×';
        remove.setAttribute('aria-label', ctx.text(msg('forget', { name: entry.name })));
        remove.title = remove.getAttribute('aria-label')!;
        remove.addEventListener('click', () => void changed(list.filter((f) => f !== entry)));
        li.append(remove);
      }
      ul.append(li);
    }
    body.replaceChildren(ul);
  }

  /** Shows the list, drawn again whenever it changes, until `cleanup`. */
  function show(draw: () => void): () => void {
    draw();
    views.add(draw);
    return () => void views.delete(draw);
  }

  const stop = ctx.on((event) => {
    if (event.type === 'opened') void remember();
  });

  ctx.files.registerSource({
    id: SOURCE_ID,
    label: msg('title'),
    icon: ICON,
    async pick() {
      await loaded;
      const chosen = await ctx.ui.dialog({
        title: msg('choose'),
        render: (body, close) => show(() => drawList(body, (entry) => close(entry.id))),
      });
      const entry = list.find((f) => f.id === chosen);
      return entry ? load(entry) : null;
    },
  });

  ctx.ui.registerPanel({
    id: 'list',
    title: msg('title'),
    icon: ICON,
    render: (body) => show(() => drawList(body, (entry) => void reopen(entry), true)),
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
