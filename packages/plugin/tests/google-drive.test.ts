import { zipSync, strToU8 } from 'fflate';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { Command, FileOrigin, FileSource, PluginContext, PluginText, SavedWorkbook } from '@ruled/sheet/plugin';
import { API, SIMPLE_UPLOAD_LIMIT, UPLOAD, XLSX } from '../src/google-drive/drive';
import { GAPI_SCRIPT, GIS_SCRIPT, loadGoogle } from '../src/google-drive/google';
import { googleDrive, type GoogleDriveOptions } from '../src/google-drive';
import { workbookName } from '../src/google-drive/connector';
import { droppedFunctions, googleFunctions } from '../src/google-drive/xlsx';

const SHEET = 'application/vnd.google-apps.spreadsheet';
const BYTES = new Uint8Array([80, 75, 3, 4, 1, 2, 3]);
const Q = 'fields=id,name,mimeType,size,version,modifiedTime,parents';

/** A Google Sheets export: one sheet with a QUERY, an IMPORTRANGE inside a SUM, and a formula Excel has. */
const EXPORT = zipSync({
  '[Content_Types].xml': strToU8('<Types/>'),
  'xl/worksheets/sheet1.xml': strToU8(
    '<worksheet><sheetData>' +
      '<row r="1"><c r="A1"><f>IFERROR(__xludf.DUMMYFUNCTION("QUERY(B1:C9,&quot;&quot;select B where C &gt; 2&quot;&quot;)"),"Ada")</f><v>Ada</v></c></row>' +
      '<row r="2"><c r="A2"><f>IFERROR(__xludf.DUMMYFUNCTION("QUERY(B1:C9,""select C"")"),3)</f><v>3</v></c></row>' +
      '<row r="3"><c r="A3"><f>IFERROR(__xludf.DUMMYFUNCTION("SUM(IMPORTRANGE(""abc"",""Sheet1!A:A""))"),42)</f><v>42</v></c></row>' +
      '<row r="4"><c r="A4"><f>SUM(B1:B9)</f><v>9</v></c></row>' +
      '</sheetData></worksheet>',
  ),
});

/** Drive's metadata for the file. */
const driveFile = (over: Record<string, unknown> = {}) => ({ id: 'f1', name: 'prices.xlsx', mimeType: XLSX, size: String(BYTES.byteLength), version: '7', modifiedTime: '2026-10-15T10:00:00.000Z', parents: ['p1'], ...over });
const STAMP = '7@2026-10-15T10:00:00.000Z';

const json = (body: unknown, status = 200, headers: Record<string, string> = {}) => new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', ...headers } });
const driveError = (status: number, reason: string, message: string) => json({ error: { code: status, message, errors: [{ reason, message }] } }, status);

interface Call {
  url: string;
  method: string;
  headers: Headers;
  body: unknown;
}

/** Answers fetches from `routes` by "METHOD url"; records every call. */
function stubFetch(routes: Record<string, (call: Call) => Response | Promise<Response>>): Call[] {
  const calls: Call[] = [];
  vi.stubGlobal('fetch', async (url: string, init: RequestInit = {}) => {
    const call = { url, method: init.method ?? 'GET', headers: new Headers(init.headers), body: init.body };
    calls.push(call);
    const route = routes[`${call.method} ${url}`];
    if (!route) throw new Error(`Unexpected ${call.method} ${url}`);
    return route(call);
  });
  return calls;
}

const fileUrl = (id: string, query = Q) => `${API}/${id}?${query}&supportsAllDrives=true`;
const MEDIA = `${API}/f1?alt=media&supportsAllDrives=true`;
const EXPORT_URL = `${API}/f1/export?mimeType=${encodeURIComponent(XLSX)}`;

// Google's scripts, as the connector uses them: a token client whose popup can grant, be closed, be blocked or be refused; and the Picker.
const gis = vi.hoisted(() => ({
  popup: 'grant' as 'grant' | 'closed' | 'blocked' | 'refused',
  requests: [] as unknown[],
  revoked: [] as string[],
  config: undefined as undefined | { client_id: string; scope: string },
}));
const picker = vi.hoisted(() => ({ result: null as null | Record<string, unknown>, views: [] as Record<string, unknown>[], builder: {} as Record<string, unknown> }));

function fakeGoogle() {
  const chain = (record: Record<string, unknown>) =>
    new Proxy(record, { get: (target, key: string) => target[key] ?? ((...args: unknown[]) => ((target[`${key}Args`] = args.length > 1 ? args : args[0]), chain(target))) });
  return {
    accounts: {
      oauth2: {
        initTokenClient(config: { client_id: string; scope: string; callback: (r: unknown) => void; error_callback: (e: unknown) => void }) {
          gis.config = config;
          return {
            requestAccessToken(overrides: unknown) {
              gis.requests.push(overrides);
              if (gis.popup === 'grant') config.callback({ access_token: `token${gis.requests.length}`, expires_in: 3599 });
              else if (gis.popup === 'refused') config.callback({ error: 'access_denied' });
              else config.error_callback({ type: gis.popup === 'blocked' ? 'popup_failed_to_open' : 'popup_closed' });
            },
          };
        },
        hasGrantedAllScopes: () => true,
        revoke: (token: string, done: () => void) => (gis.revoked.push(token), done()),
      },
    },
    picker: {
      ViewId: { DOCS: 'all', FOLDERS: 'folders' },
      DocsViewMode: { LIST: 'list' },
      Feature: { SUPPORT_DRIVES: 'sdr' },
      Response: { ACTION: 'action', DOCUMENTS: 'docs' },
      Action: { PICKED: 'picked', CANCEL: 'cancel' },
      DocsView: function (this: unknown, id: string) {
        const view: Record<string, unknown> = { id };
        picker.views.push(view);
        return chain(view);
      },
      PickerBuilder: function () {
        picker.builder = {};
        return chain(
          Object.assign(picker.builder, {
            build: () => ({
              setVisible: () => (picker.builder.setCallbackArgs as (d: unknown) => void)(picker.result ? { action: 'picked', docs: [picker.result] } : { action: 'cancel' }),
            }),
          }),
        );
      },
    },
  };
}

/** The connector set up against a fake editor context: what it registered, and what it showed. */
function connector(options: Partial<GoogleDriveOptions> = {}, limits = { maxFileBytes: 50 * 1024 * 1024 }) {
  let source!: FileSource;
  const commands: Command[] = [];
  const toasts: { kind: string; text: PluginText }[] = [];
  const confirm = vi.fn(async () => true);
  const ctx = {
    signal: new AbortController().signal,
    limits,
    text: (text: PluginText) => (typeof text === 'string' ? text : 'key' in text ? text.key : (text as { en: string }).en),
    files: { registerSource: (s: FileSource) => ((source = s), () => undefined), open: vi.fn() },
    commands: { register: (c: Command) => (commands.push(c), () => undefined) },
    ui: { toast: (kind: string, text: PluginText) => void toasts.push({ kind, text }), confirm, dialog: vi.fn(), registerPanel: vi.fn() },
    on: () => () => undefined,
    storage: { get: vi.fn(), set: vi.fn(), delete: vi.fn() },
  } as unknown as PluginContext;
  googleDrive({ clientId: 'client.apps.googleusercontent.com', apiKey: 'key', appId: '1234', ...options }).setup(ctx);
  return { source, commands, toasts, confirm };
}

const saved = (bytes: Uint8Array = BYTES, fileName = 'prices.xlsx', origin?: FileOrigin) => ({ bytes, fileName, origin }) as SavedWorkbook;
const origin = (over: Partial<FileOrigin> = {}): FileOrigin => ({ source: 'google-drive', name: 'prices.xlsx', ref: { file: 'f1' }, version: STAMP, writable: true, ...over });
const key = (text: unknown) => (text as { key: string }).key.replace('plugin.page.ruled.google-drive.', '');

beforeEach(() => {
  gis.popup = 'grant';
  gis.requests = [];
  gis.revoked = [];
  picker.result = { id: 'f1', name: 'prices.xlsx', mimeType: XLSX };
  picker.views = [];
  vi.stubGlobal('window', { google: fakeGoogle(), gapi: {} });
});

afterEach(() => vi.unstubAllGlobals());

describe('googleDrive()', () => {
  it('needs the client id, API key and project number', () => {
    expect(() => googleDrive({ clientId: 'c', apiKey: 'k' } as GoogleDriveOptions)).toThrow(/appId/);
    expect(() => googleDrive({ apiKey: 'k', appId: '1' } as GoogleDriveOptions)).toThrow(/clientId/);
  });

  it('is a plugin with English and German messages, a source that can save back and copy, and a sign-out command', () => {
    const plugin = googleDrive({ clientId: 'c', apiKey: 'k', appId: '1' });
    expect(plugin).toMatchObject({ id: 'page.ruled.google-drive', apiVersion: 1 });
    expect(Object.keys(plugin.translations!)).toEqual(['en', 'de']);
    expect(Object.keys(plugin.translations!.de)).toEqual(Object.keys(plugin.translations!.en));
    const { source, commands } = connector();
    expect(source).toMatchObject({ id: 'google-drive', label: { key: 'plugin.page.ruled.google-drive.googleDrive' } });
    expect(typeof source.save).toBe('function');
    expect(typeof source.saveAs).toBe('function');
    expect(commands.map((c) => [c.id, c.where])).toEqual([['signOut', 'more']]);
  });

  it('loads Identity Services and the Picker from Google when the page hasn’t them', async () => {
    const added: string[] = [];
    const google = fakeGoogle();
    const w: Record<string, unknown> = {};
    vi.stubGlobal('window', w);
    vi.stubGlobal('document', {
      createElement: () => ({ remove: vi.fn() }),
      head: {
        append(el: { src: string; onload: () => void }) {
          added.push(el.src);
          if (el.src === GIS_SCRIPT) w.google = { ...(w.google as object), accounts: google.accounts };
          else w.gapi = { load: (_: string, cb: { callback: () => void }) => ((w.google = { ...(w.google as object), picker: google.picker }), cb.callback()) };
          el.onload();
        },
      },
    });
    const loaded = await loadGoogle();
    expect(added.sort()).toEqual([GAPI_SCRIPT, GIS_SCRIPT].sort());
    expect(loaded.picker).toBe(google.picker);
    expect(loaded.accounts).toBe(google.accounts);
  });
});

describe('opening', () => {
  it('asks for drive.file only, picks with the project number, and downloads an .xlsx as stored', async () => {
    const calls = stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile()), [`GET ${MEDIA}`]: () => new Response(BYTES) });
    const { source } = connector();
    const file = await source.pick!();
    expect(gis.config).toMatchObject({ client_id: 'client.apps.googleusercontent.com', scope: 'https://www.googleapis.com/auth/drive.file' });
    expect(picker.builder).toMatchObject({ setAppIdArgs: '1234', setDeveloperKeyArgs: 'key', setOAuthTokenArgs: 'token1', setLocaleArgs: 'en' });
    expect(picker.views[0]).toMatchObject({ id: 'all', setMimeTypesArgs: `${XLSX},application/vnd.ms-excel.sheet.macroEnabled.12,${SHEET}` });
    expect(new Uint8Array(file!.bytes as ArrayBuffer)).toEqual(BYTES);
    expect(file!.origin).toEqual(origin());
    expect(file!.notice).toBeUndefined();
    expect(calls.every((c) => c.headers.get('Authorization') === 'Bearer token1')).toBe(true);
  });

  it('opens a Google Sheet as an Excel copy that can’t be written back, naming the functions only Google has', async () => {
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ name: 'Budget', mimeType: SHEET, size: undefined })), [`GET ${EXPORT_URL}`]: () => new Response(EXPORT) });
    picker.result = { id: 'f1', name: 'Budget', mimeType: SHEET };
    const { source } = connector();
    const file = await source.pick!();
    expect(file!.origin).toEqual(origin({ name: 'Budget.xlsx', writable: false }));
    expect(file!.notice!.kind).toBe('warn');
    expect(key(file!.notice!.text)).toBe('googleSheetDropped');
    expect((file!.notice!.text as { params: object }).params).toMatchObject({ name: 'Budget' });
    expect(file!.notice!.items).toEqual([
      { key: 'plugin.page.ruled.google-drive.droppedCells', params: { function: 'QUERY', count: 2 } },
      { key: 'plugin.page.ruled.google-drive.droppedCells', params: { function: 'IMPORTRANGE', count: 1 } },
    ]);
  });

  it('says only that it is a copy when the export kept every formula', async () => {
    const plain = zipSync({ 'xl/worksheets/sheet1.xml': strToU8('<worksheet><f>SUM(A1:A3)</f></worksheet>') });
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ name: 'Budget', mimeType: SHEET })), [`GET ${EXPORT_URL}`]: () => new Response(plain) });
    const file = await connector().source.pick!();
    expect(file!.notice).toMatchObject({ kind: 'info', items: [] });
    expect(key(file!.notice!.text)).toBe('googleSheet');
  });

  it('says plainly when a Google Sheet is too large for Google to export', async () => {
    stubFetch({
      [`GET ${fileUrl('f1')}`]: () => json(driveFile({ name: 'Huge', mimeType: SHEET })),
      [`GET ${EXPORT_URL}`]: () => driveError(403, 'exportSizeLimitExceeded', 'This file is too large to be exported.'),
    });
    const err = (await connector().source.pick!().catch((e: unknown) => e)) as { text: unknown };
    expect(key(err.text)).toBe('exportTooBig');
    expect((err.text as { params: object }).params).toEqual({ name: 'Huge' });
  });

  it('refuses an .xlsx larger than the editor opens, before downloading it', async () => {
    const calls = stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ size: String(3 * 1024 * 1024) })) });
    const { source, toasts } = connector({}, { maxFileBytes: 2 * 1024 * 1024 });
    expect(await source.pick!()).toBeNull();
    expect(calls).toHaveLength(1);
    expect(toasts[0]).toMatchObject({ kind: 'error', text: { params: { name: 'prices.xlsx', size: 3, max: 2 } } });
  });

  it('resolves null when the sign-in popup or the Picker is closed', async () => {
    stubFetch({});
    gis.popup = 'closed';
    expect(await connector().source.pick!()).toBeNull();
    gis.popup = 'grant';
    picker.result = null;
    expect(await connector().source.pick!()).toBeNull();
  });

  it('asks, and signs in on the answer, when the browser blocked the popup', async () => {
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile()), [`GET ${MEDIA}`]: () => new Response(BYTES) });
    gis.popup = 'blocked';
    const { source, confirm } = connector();
    confirm.mockImplementation(async () => ((gis.popup = 'grant'), true));
    expect(await source.pick!()).not.toBeNull();
    expect(confirm).toHaveBeenCalledTimes(1);
    expect(gis.requests).toHaveLength(2);
  });

  it('says so when the person doesn’t allow access to Drive', async () => {
    gis.popup = 'refused';
    const err = (await connector().source.pick!().catch((e: unknown) => e)) as { text: unknown };
    expect(key(err.text)).toBe('notAllowed');
  });

  it('opens a file again by its origin, as it is now, without the Picker', async () => {
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ version: '9' })), [`GET ${MEDIA}`]: () => new Response(BYTES) });
    const file = await connector().source.fetch!(origin());
    expect(picker.views).toHaveLength(0);
    expect(file!.origin.version).toBe('9@2026-10-15T10:00:00.000Z');
  });

  it('gives the name an extension the editor opens', () => {
    expect(workbookName({ name: 'prices', mimeType: XLSX })).toBe('prices.xlsx');
    expect(workbookName({ name: 'macros', mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12' })).toBe('macros.xlsm');
    expect(workbookName({ name: 'Budget.xlsx', mimeType: SHEET })).toBe('Budget.xlsx.xlsx');
  });
});

describe('saving back', () => {
  const UPDATE = `${UPLOAD}/f1?uploadType=multipart&${Q}&supportsAllDrives=true`;

  it('writes over the file only if its version and time are as opened, taking the new ones', async () => {
    let body: Blob | undefined;
    const calls = stubFetch({
      [`GET ${fileUrl('f1')}`]: () => json(driveFile()),
      [`PATCH ${UPDATE}`]: (c) => ((body = c.body as Blob), json(driveFile({ version: '8', modifiedTime: '2026-10-15T11:00:00.000Z' }))),
    });
    const { source } = connector();
    const next = await source.save!(saved(), origin());
    expect(next).toEqual(origin({ version: '8@2026-10-15T11:00:00.000Z' }));
    expect(calls.map((c) => c.method)).toEqual(['GET', 'PATCH']);
    expect(calls[1].headers.get('Content-Type')).toMatch(/^multipart\/related; boundary=/);
    const text = await body!.text();
    expect(text).toContain(`Content-Type: ${XLSX}`);
    expect(text).toContain('\r\n{}\r\n');
  });

  it('throws SaveConflictError when the file changed since it was opened', async () => {
    const calls = stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ modifiedTime: '2026-10-15T10:30:00.000Z' })) });
    const err = (await connector().source.save!(saved(), origin()).catch((e: unknown) => e)) as Error;
    expect(err.name).toBe('SaveConflictError');
    expect(calls).toHaveLength(1);
  });

  it('writes a large file through a resumable session', async () => {
    const big = new Uint8Array(SIMPLE_UPLOAD_LIMIT + 1);
    const session = 'https://www.googleapis.com/upload/drive/v3/files/f1?uploadType=resumable&upload_id=xyz';
    const calls = stubFetch({
      [`GET ${fileUrl('f1')}`]: () => json(driveFile()),
      [`PATCH ${UPLOAD}/f1?uploadType=resumable&${Q}&supportsAllDrives=true`]: () => new Response(null, { status: 200, headers: { Location: session } }),
      [`PUT ${session}`]: () => json(driveFile({ version: '8' })),
    });
    await connector().source.save!(saved(big), origin());
    expect(calls[1].headers.get('X-Upload-Content-Length')).toBe(String(big.byteLength));
    expect(calls[2].headers.get('Authorization')).toBeNull();
  });

  it('never writes over a Google Sheet', async () => {
    const calls = stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile({ mimeType: SHEET })) });
    await expect(connector().source.save!(saved(), origin())).rejects.toThrow(/never written over/);
    expect(calls).toHaveLength(1);
  });

  it('asks to sign in again when Drive refuses the token', async () => {
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => driveError(401, 'authError', 'Invalid Credentials') });
    const { source } = connector();
    const err = (await source.save!(saved(), origin()).catch((e: unknown) => e)) as { text: unknown };
    expect(key(err.text)).toBe('signInAgain');
  });
});

describe('saving a copy', () => {
  const CREATE = `${UPLOAD}?uploadType=multipart&${Q}&supportsAllDrives=true`;

  it('writes a new .xlsx into the folder picked, starting beside the Google Sheet it came from', async () => {
    let body: Blob | undefined;
    stubFetch({
      [`GET ${fileUrl('f1')}`]: () => json(driveFile({ name: 'Budget', mimeType: SHEET, parents: ['p1'] })),
      [`POST ${CREATE}`]: (c) => ((body = c.body as Blob), json(driveFile({ id: 'f2', name: 'Budget.xlsx', parents: ['p2'] }))),
    });
    picker.result = { id: 'p2', name: 'Reports', mimeType: 'application/vnd.google-apps.folder' };
    const made = await connector().source.saveAs!(saved(BYTES, 'Budget.xlsx', origin({ name: 'Budget.xlsx', writable: false })));
    expect(picker.views[0]).toMatchObject({ id: 'folders', setSelectFolderEnabledArgs: true, setParentArgs: 'p1' });
    expect(await body!.text()).toContain(JSON.stringify({ name: 'Budget.xlsx', parents: ['p2'], mimeType: XLSX }));
    expect(made).toEqual({ source: 'google-drive', name: 'Budget.xlsx', ref: { file: 'f2' }, version: STAMP, writable: true });
  });

  it('resolves null when the Picker is closed', async () => {
    const calls = stubFetch({});
    picker.result = null;
    expect(await connector().source.saveAs!(saved())).toBeNull();
    expect(calls).toHaveLength(0);
  });
});

describe('the rest', () => {
  it('revokes access and forgets the token on signing out, asking again next time', async () => {
    stubFetch({ [`GET ${fileUrl('f1')}`]: () => json(driveFile()), [`GET ${MEDIA}`]: () => new Response(BYTES) });
    const { source, commands, toasts } = connector();
    await source.pick!();
    await commands[0].run({});
    expect(gis.revoked).toEqual(['token1']);
    expect(key(toasts.at(-1)!.text)).toBe('signedOut');
    await source.pick!();
    expect(gis.requests).toHaveLength(2);
  });

  it('names the functions only Google has in a Sheets formula', () => {
    expect(googleFunctions('SUM(QUERY(A:B,"select SUM(A)"))')).toEqual(['QUERY']);
    expect(googleFunctions('MYSTERY(A1)+SUM(B1)')).toEqual(['MYSTERY']);
  });

  it('reads the functions an export dropped, most used first', async () => {
    expect(await droppedFunctions(EXPORT)).toEqual([
      { name: 'QUERY', cells: 2 },
      { name: 'IMPORTRANGE', cells: 1 },
    ]);
  });
});
