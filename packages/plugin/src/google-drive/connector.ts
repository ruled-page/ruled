/** The Google Drive connector's workings, loaded on demand by its entry (index.ts), so a host's bundle carries only the entry. */
import type { FileOrigin, PluginContext, PluginText, RemoteFile, SavedWorkbook } from '@ruled/sheet/plugin';
import { download, DriveError, exportXlsx, FOLDER, getFile, GOOGLE_SHEET, upload, XLSM, XLSX, type DriveFile, type GetToken } from './drive';
import { createAuth, isPopupBlocked, loadGoogle, showPicker, type Auth, type GoogleGlobals } from './google';
import type { GoogleDriveOptions } from './index';
import { GOOGLE_DRIVE_SOURCE_ID, msg } from './messages';
import { droppedFunctions } from './xlsx';

/** An error the editor shows in its own words, translated. */
const textError = (text: PluginText, message: string) => Object.assign(new Error(message), { text });

/** The person declined to sign in or closed a window. */
class Cancelled extends Error {
  constructor() {
    super('Cancelled');
    this.name = 'Cancelled';
  }
}

/** A size in megabytes, to one decimal place, for messages. */
const mb = (bytes: number) => Math.round((bytes / 1024 / 1024) * 10) / 10;

/** Drive's version number and modification time together: if either moved, someone else changed the file. */
const stamp = (file: DriveFile) => `${file.version ?? ''}@${file.modifiedTime ?? ''}`;

/**
 * The name the editor shows, which must end in .xlsx or .xlsm for it to
 * open the file: Drive names needn't have an extension, and a Google Sheet
 * opens as "<its name>.xlsx".
 */
export function workbookName(file: Pick<DriveFile, 'name' | 'mimeType'>): string {
  if (file.mimeType !== GOOGLE_SHEET && /\.xls[xm]$/i.test(file.name)) return file.name;
  return `${file.name}${file.mimeType === XLSM ? '.xlsm' : '.xlsx'}`;
}

/** The connector's workings, which its entry loads when they are first needed. */
export function connect(ctx: PluginContext, o: GoogleDriveOptions, label: PluginText) {
  const named = { source: label };

  let google: { globals: GoogleGlobals; auth: Auth } | undefined;
  let loading: Promise<{ globals: GoogleGlobals; auth: Auth }> | undefined;
  let signedInOnce = false;
  const load = () =>
    (loading ??= loadGoogle().then(
      (globals) => (google = { globals, auth: createAuth(globals, o.clientId) }),
      (err: unknown) => {
        loading = undefined;
        throw err;
      },
    ));
  // Loading Google's scripts as a menu opens lets the click that follows open its popup straight away.
  const preload = () => void load().catch(() => undefined);

  /**
   * A token for Drive. Getting a new one opens Google's popup, which
   * browsers allow only during a click; when it was blocked, because the
   * click was spent loading, the person is asked, and their answer's click opens it.
   */
  async function token(): Promise<string> {
    const { auth } = google ?? (await load());
    const current = auth.current();
    if (current) return current;
    const body = msg(signedInOnce ? 'signInAgain' : 'signInBody', named);
    let got: string | null;
    try {
      got = await auth.request();
    } catch (err) {
      if (!isPopupBlocked(err)) throw err;
      if (!(await ctx.ui.confirm({ title: label, body, confirmLabel: msg('signIn') }))) throw new Cancelled();
      got = await auth.request();
    }
    if (!got) throw new Cancelled();
    signedInOnce = true;
    return got;
  }
  const tok: GetToken = token;

  /** A Drive failure or a refused sign-in in the editor's words; other errors as they are. */
  function described(err: unknown): unknown {
    if ((err as Error)?.name === 'NotAllowed') return textError(msg('notAllowed', named), (err as Error).message);
    if (!(err instanceof DriveError)) return err;
    if (err.status === 401) {
      google?.auth.forget();
      return textError(msg('signInAgain', named), `Google Drive refused the token: ${err.message}`);
    }
    return textError(msg('refused', { ...named, status: err.status, reason: err.message }), `Google Drive refused the request (${err.status}): ${err.message}`);
  }

  /** Runs a step the person can back out of, resolving null when they do. */
  async function unlessCancelled<T>(step: () => Promise<T>): Promise<T | null> {
    try {
      return await step();
    } catch (err) {
      if (err instanceof Cancelled) return null;
      throw described(err);
    }
  }

  const originOf = (file: DriveFile): FileOrigin => ({
    source: GOOGLE_DRIVE_SOURCE_ID,
    name: workbookName(file),
    ref: { file: file.id },
    version: stamp(file),
    // Writing an .xlsx over a Google Sheet would have Google convert it back, losing what the export dropped.
    writable: file.mimeType !== GOOGLE_SHEET,
  });

  const idOf = (origin: FileOrigin): string => {
    const id = (origin.ref as { file?: unknown } | null)?.file;
    if (typeof id !== 'string') throw new Error('This workbook’s origin doesn’t name a Google Drive file.');
    return id;
  };

  const tooBig = (name: string, size: number) => {
    ctx.ui.toast('error', msg('tooBig', { name, size: mb(size), max: mb(ctx.limits.maxFileBytes) }));
    return null;
  };

  /** The file as it is now: an .xlsx as stored, a Google Sheet exported as one, with a notice saying so. */
  async function fetched(id: string): Promise<RemoteFile | null> {
    const file = await getFile(tok, id, ctx.signal);
    const max = ctx.limits.maxFileBytes;
    if (file.mimeType !== GOOGLE_SHEET) {
      if (Number(file.size ?? 0) > max) return tooBig(file.name, Number(file.size));
      return { bytes: await download(tok, id, ctx.signal), origin: originOf(file) };
    }
    let bytes: ArrayBuffer;
    try {
      bytes = await exportXlsx(tok, id, ctx.signal);
    } catch (err) {
      if (err instanceof DriveError && err.reason === 'exportSizeLimitExceeded') {
        throw textError(msg('exportTooBig', { name: file.name }), `${file.name} is larger than the 10 MB Google exports.`);
      }
      throw err;
    }
    if (bytes.byteLength > max) return tooBig(file.name, bytes.byteLength);
    const dropped = await droppedFunctions(bytes).catch(() => []);
    return {
      bytes,
      origin: originOf(file),
      notice: {
        kind: dropped.length ? 'warn' : 'info',
        text: msg(dropped.length ? 'googleSheetDropped' : 'googleSheet', { ...named, name: file.name }),
        items: dropped.map((d) => msg('droppedCells', { function: d.name, count: d.cells })),
      },
    };
  }

  /** Google's Picker, once there is a token for it. */
  async function picked(mode: 'files' | 'folders', parent?: string) {
    const t = await token();
    const { globals } = google!;
    const chosen = await showPicker(globals, {
      token: t,
      apiKey: o.apiKey,
      appId: o.appId,
      mode,
      locale: ctx.text({ en: 'en', de: 'de' }),
      title: ctx.text(msg(mode === 'files' ? 'pickFile' : 'pickFolder')),
      parent,
      mimeTypes: mode === 'files' ? [XLSX, XLSM, GOOGLE_SHEET] : [FOLDER],
    });
    if (!chosen) throw new Cancelled();
    return chosen;
  }

  const pick = () => unlessCancelled(async () => fetched((await picked('files')).id));

  /** A file opened before, by its origin: from a list of recent files. */
  function fetchAgain(origin: FileOrigin): Promise<RemoteFile | null> {
    const id = idOf(origin);
    return unlessCancelled(async () => {
      await token();
      return fetched(id);
    });
  }

  async function save(file: SavedWorkbook, origin: FileOrigin): Promise<FileOrigin> {
    const id = idOf(origin);
    try {
      // Drive v3 has no ETags, so no If-Match: compare the file's version and time just before writing.
      const now = await getFile(tok, id, ctx.signal);
      if (now.mimeType === GOOGLE_SHEET) throw new Error('A Google Sheet is never written over; save a copy instead.');
      // An empty message lets the editor say what a conflict means in its own words.
      if (stamp(now) !== origin.version) throw Object.assign(new Error(''), { name: 'SaveConflictError' });
      const saved = await upload(tok, { id, metadata: {}, bytes: file.bytes }, ctx.signal);
      return { ...origin, name: workbookName(saved), version: stamp(saved) };
    } catch (err) {
      if (err instanceof Cancelled) throw textError(msg('signInBody', named), 'Not signed in to Google Drive.');
      throw described(err);
    }
  }

  async function saveAs(file: SavedWorkbook): Promise<FileOrigin | null> {
    return unlessCancelled(async () => {
      const from = file.origin?.source === GOOGLE_DRIVE_SOURCE_ID ? file.origin : undefined;
      // The folder picker starts beside the file the workbook came from.
      const parent = from && (await token()) ? await getFile(tok, idOf(from), ctx.signal).then((f) => f.parents?.[0], () => undefined) : undefined;
      const folder = await picked('folders', parent);
      const made = await upload(tok, { metadata: { name: file.fileName, parents: [folder.id], mimeType: XLSX }, bytes: file.bytes }, ctx.signal);
      return originOf(made);
    });
  }

  async function signOut(): Promise<void> {
    await (google ?? (await load())).auth.signOut();
    signedInOnce = false;
    ctx.ui.toast('info', msg('signedOut', named));
  }

  return { preload, pick, fetch: fetchAgain, save, saveAs, signOut };
}

export type Connector = ReturnType<typeof connect>;
