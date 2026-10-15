/**
 * Open workbooks from OneDrive and SharePoint, save them back, and save
 * copies there: `@ruled/plugin/onedrive`.
 *
 * ```ts
 * import { oneDrive } from '@ruled/plugin/onedrive';
 * const editor = useWorkbookEditor({ plugins: [oneDrive({ clientId: 'your-entra-app-id' })] });
 * ```
 */
import type { FileOrigin, FileSource, Localized, PluginContext, PluginText, RemoteFile, RuledPlugin, SavedWorkbook } from '@ruled/sheet/plugin';
import { isPopupBlocked, loadAuth, NeedsSignIn, type Auth } from './auth';
import { create, download, folderPath, getItem, GraphError, myDrive, replace, type DriveItem, type GetToken, type ItemRef } from './graph';
import { translations, type MessageKey } from './messages';
import { CONSUMER_PICKER, runPicker, type PickedItem } from './picker';

export interface OneDriveOptions {
  /** The application (client) id of your app registration in Microsoft Entra ID. */
  clientId: string;
  /**
   * Who can sign in: 'common' (personal, work and school accounts, the
   * default), 'organizations' (work and school only), 'consumers' (personal
   * only), or a tenant's id or domain (that organisation only).
   */
  authority?: 'common' | 'organizations' | 'consumers' | (string & {});
  /**
   * Where the sign-in popup returns: a page of yours that runs
   * `@ruled/plugin/onedrive/redirect`, registered as a single-page
   * application redirect URI. Default: /onedrive-redirect.html on the page's origin.
   */
  redirectUri?: string;
  /** Pick from a SharePoint site, such as "https://contoso.sharepoint.com/sites/finance", rather than the person's OneDrive. */
  sharePoint?: { baseUrl: string };
  /** Only open files: asks for read permission only, and offers no saving back or copies. */
  readOnly?: boolean;
  /** The name shown in the Open menu and on "Save to …". Default "OneDrive", or "SharePoint" with `sharePoint`. */
  label?: Localized;
}

/** The plugin's id; its messages are `plugin.page.ruled.onedrive.<key>`. */
export const ONEDRIVE_PLUGIN_ID = 'page.ruled.onedrive';

/** The file source's id, which a workbook's `origin.source` names when it came from OneDrive. */
export const ONEDRIVE_SOURCE_ID = 'onedrive';

const msg = (key: MessageKey, params?: Record<string, unknown>): PluginText => ({ key: `plugin.${ONEDRIVE_PLUGIN_ID}.${key}`, params });

/** An error the editor shows in its own words, translated. */
const textError = (text: PluginText, message: string) => Object.assign(new Error(message), { text });

/** The person declined to sign in or to open a window. */
class Cancelled extends Error {
  constructor() {
    super('Cancelled');
    this.name = 'Cancelled';
  }
}

/** A size in megabytes, to one decimal place, for messages. */
const mb = (bytes: number) => Math.round((bytes / 1024 / 1024) * 10) / 10;

/** "forecast (2).xlsx" after "forecast.xlsx", and "forecast (3).xlsx" after that. */
export function nextName(name: string): string {
  const [, stem, extension = ''] = /^(.*?)(\.[^.]+)?$/.exec(name)!;
  const numbered = /^(.*) \((\d+)\)$/.exec(stem);
  return numbered ? `${numbered[1]} (${Number(numbered[2]) + 1})${extension}` : `${stem} (2)${extension}`;
}

/** The connector, to pass in the editor's `plugins`. */
export function oneDrive(options: OneDriveOptions): RuledPlugin {
  if (!options?.clientId) throw new Error('oneDrive() needs the client id of your Microsoft Entra app registration.');
  return {
    id: ONEDRIVE_PLUGIN_ID,
    apiVersion: 1,
    translations,
    setup: (ctx) => setup(ctx, options),
  };
}

function setup(ctx: PluginContext, o: OneDriveOptions): void {
  const label: PluginText = o.label ?? msg(o.sharePoint ? 'sharePoint' : 'oneDrive');
  const named = { source: label };
  const graphScopes = [o.readOnly ? 'Files.Read.All' : 'Files.ReadWrite.All'];
  // Personal OneDrive's picker takes a token for these; work and school pickers, one for their SharePoint.
  const consumerPickerScopes = [o.readOnly ? 'OneDrive.ReadOnly' : 'OneDrive.ReadWrite'];

  let auth: Auth | undefined;
  let loading: Promise<Auth> | undefined;
  const load = () =>
    (loading ??= loadAuth({
      clientId: o.clientId,
      authority: o.authority ?? 'common',
      redirectUri: o.redirectUri ?? new URL('/onedrive-redirect.html', location.origin).href,
    }).then(
      (loaded) => (auth = loaded),
      (err: unknown) => {
        loading = undefined;
        throw err;
      },
    ));
  // Loading MSAL as a menu opens lets the click that follows open its popup straight away.
  const preload = () => void load().catch(() => undefined);

  /**
   * Runs `open`, which opens a popup. Browsers only allow one during a
   * click; when it was blocked, because the click was spent loading or
   * signing in, the person is asked, and their answer's click opens it.
   */
  async function withPopup<T>(open: () => Promise<T>, body: PluginText, confirmLabel: PluginText): Promise<T> {
    try {
      return await open();
    } catch (err) {
      if (!isPopupBlocked(err)) throw err;
    }
    if (!(await ctx.ui.confirm({ title: label, body, confirmLabel }))) throw new Cancelled();
    return open();
  }

  /** MSAL, with an account signed in; null when the person doesn't sign in. */
  async function signedIn(): Promise<Auth> {
    const a = auth ?? (await load());
    if (a.account()) return a;
    if (!(await withPopup(() => a.signIn(graphScopes), msg('signInBody', named), msg('signIn')))) throw new Cancelled();
    return a;
  }

  /** A token, asking the person to sign in again when only a popup can get one. */
  async function token(a: Auth, scopes: string[]): Promise<string> {
    try {
      return await a.token(scopes);
    } catch (err) {
      if (!(err instanceof NeedsSignIn)) throw err;
    }
    const got = await withPopup(() => a.tokenPopup(scopes), msg(a.account() ? 'signInAgain' : 'signInBody', named), msg('signIn'));
    if (!got) throw new Cancelled();
    return got;
  }

  /** Where the picker lives for this account, and the origin its tokens are for. */
  async function pickerPlace(a: Auth, graphToken: GetToken): Promise<{ url: string; site?: string; resource?: string }> {
    if (o.sharePoint) {
      const site = o.sharePoint.baseUrl.replace(/\/+$/, '');
      return { url: `${site}/_layouts/15/FilePicker.aspx`, site, resource: new URL(site).origin };
    }
    if (a.consumer()) return { url: CONSUMER_PICKER };
    const { webUrl } = await myDrive(graphToken, ctx.signal);
    const origin = new URL(webUrl).origin;
    return { url: `${origin}/_layouts/15/FilePicker.aspx`, resource: origin };
  }

  /** Lets the person choose a file, or a folder, in the picker. */
  async function choose(mode: 'files' | 'folders'): Promise<{ a: Auth; picked: PickedItem }> {
    const a = await signedIn();
    const graphToken = () => token(a, graphScopes);
    const picked = await withPopup(
      async () => {
        const win = window.open('', 'ruled-onedrive-picker', 'width=1080,height=680');
        if (!win) throw Object.assign(new Error('The browser blocked the picker’s window.'), { name: 'PopupBlocked' });
        try {
          const place = await pickerPlace(a, graphToken);
          return await runPicker({
            win,
            url: place.url,
            site: place.site,
            mode,
            pickLabel: mode === 'folders' ? ctx.text(msg('saveHere')) : undefined,
            locale: ctx.text({ en: 'en-us', de: 'de-de' }),
            token: (resource) =>
              place.resource ? token(a, [`${new URL(resource ?? place.resource).origin}/.default`]) : token(a, consumerPickerScopes),
            signal: ctx.signal,
          });
        } catch (err) {
          win.close();
          throw err;
        }
      },
      msg(mode === 'files' ? 'chooseFile' : 'chooseFolder', named),
      msg('choose'),
    );
    if (!picked) throw new Cancelled();
    return { a, picked };
  }

  /** A Graph failure in the editor's words; other errors as they are. */
  const described = (err: unknown) =>
    err instanceof GraphError ? textError(msg('refused', { ...named, status: err.status, reason: err.message }), `OneDrive refused the request (${err.status}): ${err.message}`) : err;

  const originOf = (drive: string, item: DriveItem): FileOrigin => {
    const folder = folderPath(item);
    return {
      source: ONEDRIVE_SOURCE_ID,
      name: item.name,
      ref: { drive, item: item.id },
      version: item.eTag,
      writable: !o.readOnly,
      location: folder ? `${folder}/${item.name}` : undefined,
    };
  };

  const refOf = (origin: FileOrigin): ItemRef => {
    const ref = origin.ref as { drive?: unknown; item?: unknown } | null;
    if (typeof ref?.drive !== 'string' || typeof ref.item !== 'string') throw new Error('This workbook’s origin doesn’t name a OneDrive file.');
    return { drive: ref.drive, item: ref.item };
  };

  /** Runs a step the person can back out of, resolving null when they do. */
  async function unlessCancelled<T>(step: () => Promise<T>): Promise<T | null> {
    try {
      return await step();
    } catch (err) {
      if (err instanceof Cancelled) return null;
      throw described(err);
    }
  }

  async function pick(): Promise<RemoteFile | null> {
    return unlessCancelled(async () => {
      const { a, picked } = await choose('files');
      const graphToken = () => token(a, graphScopes);
      const ref = { drive: picked.parentReference.driveId, item: picked.id };
      const item = await getItem(graphToken, ref, ctx.signal);
      const max = ctx.limits.maxFileBytes;
      if ((item.size ?? 0) > max) {
        ctx.ui.toast('error', msg('tooBig', { name: item.name, size: mb(item.size!), max: mb(max) }));
        return null;
      }
      const bytes = await download(graphToken, ref, item, ctx.signal);
      return { bytes, origin: originOf(ref.drive, item) };
    });
  }

  async function save(file: SavedWorkbook, origin: FileOrigin): Promise<FileOrigin> {
    const ref = refOf(origin);
    try {
      const a = auth ?? (await load());
      const item = await replace(() => token(a, graphScopes), ref, origin.name, file.bytes, origin.version, ctx.signal);
      return { ...origin, name: item.name, version: item.eTag };
    } catch (err) {
      // An empty message lets the editor say what a conflict means in its own words.
      if (err instanceof GraphError && err.status === 412) throw Object.assign(new Error(''), { name: 'SaveConflictError' });
      if (err instanceof Cancelled) throw textError(msg('signInBody', named), 'Not signed in to OneDrive.');
      throw described(err);
    }
  }

  /** Another name for the copy, when the folder has a file by this one; null when the person cancels. */
  async function askName(taken: string): Promise<string | null> {
    let input: HTMLInputElement | undefined;
    const action = await ctx.ui.dialog({
      title: msg('nameTakenTitle'),
      render(body) {
        const text = document.createElement('p');
        text.textContent = ctx.text(msg('nameTakenBody', { name: taken }));
        const field = document.createElement('label');
        field.textContent = ctx.text(msg('fileName'));
        field.style.display = 'grid';
        field.style.gap = '4px';
        input = document.createElement('input');
        input.className = 'wbe-input';
        input.value = nextName(taken);
        field.append(input);
        body.append(text, field);
        const focus = input;
        setTimeout(() => {
          focus.focus();
          focus.setSelectionRange(0, focus.value.lastIndexOf('.') > 0 ? focus.value.lastIndexOf('.') : focus.value.length);
        });
      },
      actions: [
        { id: 'cancel', label: msg('cancel') },
        { id: 'save', label: msg('save'), primary: true },
      ],
    });
    const name = input?.value.trim();
    if (action !== 'save' || !name) return null;
    return /\.[^.]+$/.test(name) ? name : `${name}${/\.[^.]+$/.exec(taken)?.[0] ?? '.xlsx'}`;
  }

  async function saveAs(file: SavedWorkbook): Promise<FileOrigin | null> {
    return unlessCancelled(async () => {
      const { a, picked } = await choose('folders');
      const folder = { drive: picked.parentReference.driveId, item: picked.id };
      for (let name = file.fileName; ; ) {
        try {
          return originOf(folder.drive, await create(() => token(a, graphScopes), folder, name, file.bytes, ctx.signal));
        } catch (err) {
          if (!(err instanceof GraphError && err.status === 409)) throw err;
          const next = await askName(name);
          if (!next) return null;
          name = next;
        }
      }
    });
  }

  const source: FileSource = { id: ONEDRIVE_SOURCE_ID, label, preload, pick };
  if (!o.readOnly) Object.assign(source, { save, saveAs });
  ctx.files.registerSource(source);

  ctx.commands.register({
    id: 'signOut',
    label: msg('signOut', named),
    where: 'more',
    preload,
    async run() {
      await (auth ?? (await load())).signOut();
      ctx.ui.toast('info', msg('signedOut', named));
    },
  });
}
