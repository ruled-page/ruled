import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest';
import type { Command, FileOrigin, FileSource, PluginContext, PluginText, SavedWorkbook } from '@ruled/sheet/plugin';
import { CHUNK_BYTES, SIMPLE_UPLOAD_LIMIT } from '../src/onedrive/graph';
import { nextName, oneDrive, type OneDriveOptions } from '../src/onedrive';
import type { PickerOptions } from '../src/onedrive/picker';

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- each mock takes the call MSAL's method would
type Fn = Mock<(...args: any[]) => Promise<unknown>>;

// MSAL as the connector uses it: an account (or none), tokens by scope, and popups that can be blocked.
const msal = vi.hoisted(() => ({
  account: null as { tenantId: string; username: string } | null,
  config: undefined as unknown,
  silentFails: false,
  loginPopup: undefined as unknown as Fn,
  acquireTokenPopup: undefined as unknown as Fn,
  acquireTokenSilent: undefined as unknown as Fn,
  clearCache: undefined as unknown as Fn,
}));

vi.mock('@azure/msal-browser', () => {
  class InteractionRequiredAuthError extends Error {}
  class PublicClientApplication {
    constructor(config: unknown) {
      msal.config = config;
    }
    initialize = async () => undefined;
    getActiveAccount = () => msal.account;
    getAllAccounts = () => (msal.account ? [msal.account] : []);
    setActiveAccount = () => undefined;
    loginPopup = (...args: unknown[]) => msal.loginPopup(...args);
    acquireTokenPopup = (...args: unknown[]) => msal.acquireTokenPopup(...args);
    acquireTokenSilent = (request: { scopes: string[] }) => {
      if (msal.silentFails) return Promise.reject(new InteractionRequiredAuthError('interaction_required'));
      return msal.acquireTokenSilent(request);
    };
    clearCache = (...args: unknown[]) => msal.clearCache(...args);
  }
  return { PublicClientApplication, InteractionRequiredAuthError };
});

const picker = vi.hoisted(() => ({ runPicker: vi.fn() }));
vi.mock('../src/onedrive/picker', async (original) => ({ ...(await original<object>()), runPicker: picker.runPicker }));

const PERSONAL = { tenantId: '9188040d-6c67-4c5b-b112-36a304b66dad', username: 'ada@outlook.com' };
const WORK = { tenantId: '72f988bf-0000-0000-0000-000000000000', username: 'ada@contoso.com' };
const GRAPH = 'https://graph.microsoft.com/v1.0';
const ITEM = `${GRAPH}/drives/d1/items/i1`;
const BYTES = new Uint8Array([80, 75, 3, 4, 1, 2, 3]);

/** Graph's driveItem for the file, at a version. */
const driveItem = (over: Record<string, unknown> = {}) => ({
  id: 'i1',
  name: 'prices.xlsx',
  size: BYTES.byteLength,
  eTag: '"{ABC},1"',
  parentReference: { driveId: 'd1', path: '/drive/root:/Documents/Finance' },
  file: {},
  '@microsoft.graph.downloadUrl': 'https://public.dm.files.1drv.com/download?t=x',
  ...over,
});

const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });

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

/** The connector set up against a fake editor context: what it registered, and what it showed. */
function connector(options: Partial<OneDriveOptions> = {}, limits = { maxFileBytes: 50 * 1024 * 1024 }) {
  let source!: FileSource;
  const commands: Command[] = [];
  const toasts: { kind: string; text: PluginText }[] = [];
  const confirm = vi.fn(async () => true);
  const dialog = vi.fn(async (d: { render(body: HTMLElement): unknown }) => {
    d.render(fakeElement() as unknown as HTMLElement);
    return 'save';
  });
  const ctx = {
    signal: new AbortController().signal,
    limits,
    text: (text: PluginText) => (typeof text === 'string' ? text : 'key' in text ? text.key : (text as { en: string }).en),
    files: { registerSource: (s: FileSource) => ((source = s), () => undefined), open: vi.fn() },
    commands: { register: (c: Command) => (commands.push(c), () => undefined) },
    ui: { toast: (kind: string, text: PluginText) => void toasts.push({ kind, text }), confirm, dialog, registerPanel: vi.fn() },
    on: () => () => undefined,
    storage: { get: vi.fn(), set: vi.fn(), delete: vi.fn() },
  } as unknown as PluginContext;
  oneDrive({ clientId: 'app-id', ...options }).setup(ctx);
  return { source, commands, toasts, confirm, dialog };
}

/** Just enough of an element for the name dialog. */
function fakeElement(): Record<string, unknown> {
  const el: Record<string, unknown> = { style: {}, children: [] as unknown[], value: '', textContent: '' };
  el.append = (...kids: unknown[]) => (el.children as unknown[]).push(...kids);
  el.focus = () => undefined;
  el.setSelectionRange = () => undefined;
  return el;
}

/** What onSave and sources get: only the bytes and the name matter here. */
const saved = (bytes: Uint8Array = BYTES, fileName = 'prices.xlsx') => ({ bytes, fileName }) as SavedWorkbook;

const origin = (over: Partial<FileOrigin> = {}): FileOrigin => ({ source: 'onedrive', name: 'prices.xlsx', ref: { drive: 'd1', item: 'i1' }, version: '"{ABC},1"', writable: true, ...over });

let popup: { closed: boolean; close: () => void } | null;

beforeEach(() => {
  msal.account = PERSONAL;
  msal.silentFails = false;
  msal.loginPopup = vi.fn(async () => {
    msal.account = PERSONAL;
    return { account: PERSONAL };
  });
  msal.acquireTokenPopup = vi.fn(async () => ({ account: msal.account, accessToken: 'popup-token' }));
  msal.acquireTokenSilent = vi.fn(async ({ scopes }: { scopes: string[] }) => ({ accessToken: `token:${scopes.join(' ')}` }));
  msal.clearCache = vi.fn(async () => {
    msal.account = null;
  });
  popup = { closed: false, close: vi.fn() };
  vi.stubGlobal('location', { origin: 'http://localhost:5180' });
  vi.stubGlobal('window', { open: vi.fn(() => popup) });
  vi.stubGlobal('document', { createElement: () => fakeElement() });
  picker.runPicker.mockReset();
  picker.runPicker.mockResolvedValue({ id: 'i1', parentReference: { driveId: 'd1' } });
});

afterEach(() => vi.unstubAllGlobals());

describe('oneDrive()', () => {
  it('needs a client id', () => {
    expect(() => oneDrive({} as OneDriveOptions)).toThrow(/client id/);
  });

  it('is a plugin with English and German messages, a source and a sign-out command', () => {
    const plugin = oneDrive({ clientId: 'app-id' });
    expect(plugin).toMatchObject({ id: 'page.ruled.onedrive', apiVersion: 1 });
    expect(Object.keys(plugin.translations!['de'])).toEqual(Object.keys(plugin.translations!['en']));
    const { source, commands } = connector();
    expect(source).toMatchObject({ id: 'onedrive', label: { key: 'plugin.page.ruled.onedrive.oneDrive' } });
    expect(source.save).toBeTypeOf('function');
    expect(source.saveAs).toBeTypeOf('function');
    expect(commands.map((c) => [c.id, c.where])).toEqual([['signOut', 'more']]);
  });

  it('signs in with MSAL kept in sessionStorage, returning to the redirect page on this origin', async () => {
    stubFetch({});
    const { source } = connector({ authority: 'organizations' });
    source.preload!();
    await vi.waitFor(() => expect(msal.config).toBeDefined());
    expect(msal.config).toEqual({
      auth: { clientId: 'app-id', authority: 'https://login.microsoftonline.com/organizations', redirectUri: 'http://localhost:5180/onedrive-redirect.html' },
      cache: { cacheLocation: 'sessionStorage' },
    });
  });
});

describe('opening', () => {
  it('reads the item, then downloads it from its download address without the token', async () => {
    const calls = stubFetch({
      [`GET ${ITEM}`]: () => json(driveItem()),
      'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES),
    });
    const { source } = connector();
    const file = await source.pick();

    expect(new Uint8Array(file!.bytes as ArrayBuffer)).toEqual(BYTES);
    expect(file!.origin).toEqual({
      source: 'onedrive',
      name: 'prices.xlsx',
      ref: { drive: 'd1', item: 'i1' },
      version: '"{ABC},1"',
      writable: true,
      location: 'Documents/Finance/prices.xlsx',
    });
    expect(calls[0].headers.get('Authorization')).toBe('Bearer token:Files.ReadWrite.All');
    expect(calls[1].headers.get('Authorization')).toBeNull();
  });

  it('opens personal OneDrive’s picker for files, with tokens for OneDrive', async () => {
    stubFetch({ [`GET ${ITEM}`]: () => json(driveItem()), 'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES) });
    await connector().source.pick();
    const options = picker.runPicker.mock.calls[0][0] as PickerOptions;
    expect(options).toMatchObject({ win: popup, url: 'https://onedrive.live.com/picker', mode: 'files', locale: 'en-us' });
    expect(await options.token(undefined)).toBe('token:OneDrive.ReadWrite');
  });

  it('opens a work account’s picker on its own SharePoint, with tokens for the resource it asks for', async () => {
    msal.account = WORK;
    stubFetch({
      [`GET ${GRAPH}/me/drive?$select=webUrl`]: () => json({ webUrl: 'https://contoso-my.sharepoint.com/personal/ada_contoso_com/Documents' }),
      [`GET ${ITEM}`]: () => json(driveItem()),
      'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES),
    });
    await connector().source.pick();
    const options = picker.runPicker.mock.calls[0][0] as PickerOptions;
    expect(options.url).toBe('https://contoso-my.sharepoint.com/_layouts/15/FilePicker.aspx');
    expect(await options.token(undefined)).toBe('token:https://contoso-my.sharepoint.com/.default');
    expect(await options.token('https://contoso.sharepoint.com')).toBe('token:https://contoso.sharepoint.com/.default');
  });

  it('opens a SharePoint site’s picker when given one', async () => {
    msal.account = WORK;
    stubFetch({ [`GET ${ITEM}`]: () => json(driveItem()), 'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES) });
    const { source } = connector({ sharePoint: { baseUrl: 'https://contoso.sharepoint.com/sites/finance/' } });
    expect(source.label).toEqual({ key: 'plugin.page.ruled.onedrive.sharePoint' });
    await source.pick();
    expect(picker.runPicker.mock.calls[0][0]).toMatchObject({ url: 'https://contoso.sharepoint.com/sites/finance/_layouts/15/FilePicker.aspx', site: 'https://contoso.sharepoint.com/sites/finance' });
  });

  it('refuses a file larger than the editor opens, before downloading it', async () => {
    const calls = stubFetch({ [`GET ${ITEM}`]: () => json(driveItem({ size: 3 * 1024 * 1024 })) });
    const { source, toasts } = connector({}, { maxFileBytes: 2 * 1024 * 1024 });
    expect(await source.pick()).toBeNull();
    expect(calls.map((c) => c.url)).toEqual([ITEM]);
    expect(toasts).toEqual([{ kind: 'error', text: { key: 'plugin.page.ruled.onedrive.tooBig', params: { name: 'prices.xlsx', size: 3, max: 2 } } }]);
  });

  it('resolves null when the picker is closed', async () => {
    const calls = stubFetch({});
    picker.runPicker.mockResolvedValue(null);
    expect(await connector().source.pick()).toBeNull();
    expect(calls).toEqual([]);
  });

  it('signs in first; when that spent the click and the picker’s window is blocked, asks, and opens it on the answer', async () => {
    msal.account = null;
    stubFetch({ [`GET ${ITEM}`]: () => json(driveItem()), 'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES) });
    const open = vi.fn().mockReturnValueOnce(null).mockReturnValue(popup);
    vi.stubGlobal('window', { open });
    const { source, confirm } = connector();
    expect(await source.pick()).not.toBeNull();
    expect(msal.loginPopup).toHaveBeenCalledWith({ scopes: ['Files.ReadWrite.All'], prompt: 'select_account' });
    expect(confirm).toHaveBeenCalledWith(expect.objectContaining({ body: { key: 'plugin.page.ruled.onedrive.chooseFile', params: { source: source.label } } }));
    expect(open).toHaveBeenCalledTimes(2);
  });

  it('resolves null when the person closes the sign-in popup', async () => {
    msal.account = null;
    msal.loginPopup = vi.fn(async () => Promise.reject(Object.assign(new Error('closed'), { errorCode: 'user_cancelled' })));
    stubFetch({});
    expect(await connector().source.pick()).toBeNull();
    expect(picker.runPicker).not.toHaveBeenCalled();
  });

  it('closes the picker’s window when it can’t find where the picker is', async () => {
    msal.account = WORK;
    stubFetch({ [`GET ${GRAPH}/me/drive?$select=webUrl`]: () => json({ error: { code: 'unauthenticated', message: 'No' } }, 401) });
    await expect(connector().source.pick()).rejects.toMatchObject({ text: { params: { status: 401 } } });
    expect(popup!.close).toHaveBeenCalled();
  });

  it('says what Graph refused, in the editor’s words', async () => {
    stubFetch({ [`GET ${ITEM}`]: () => json({ error: { code: 'accessDenied', message: 'Access denied' } }, 403) });
    const { source } = connector();
    await expect(source.pick()).rejects.toMatchObject({
      text: { key: 'plugin.page.ruled.onedrive.refused', params: { source: source.label, status: 403, reason: 'Access denied' } },
    });
  });
});

describe('opening again by origin', () => {
  it('fetches the file as it is now, without the picker, taking its current version', async () => {
    stubFetch({
      [`GET ${ITEM}`]: () => json(driveItem({ eTag: '"{ABC},5"' })),
      'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES),
    });
    const file = await connector().source.fetch!(origin());
    expect(new Uint8Array(file!.bytes as ArrayBuffer)).toEqual(BYTES);
    expect(file!.origin).toMatchObject({ ref: { drive: 'd1', item: 'i1' }, version: '"{ABC},5"' });
    expect(picker.runPicker).not.toHaveBeenCalled();
  });

  it('signs in first, and resolves null when the person won’t', async () => {
    msal.account = null;
    msal.loginPopup = vi.fn(async () => Promise.reject(Object.assign(new Error('closed'), { errorCode: 'user_cancelled' })));
    expect(await connector().source.fetch!(origin())).toBeNull();
  });

  it('refuses an origin that doesn’t name a OneDrive file', async () => {
    await expect(connector().source.fetch!(origin({ ref: 'elsewhere' }))).rejects.toThrow(/doesn’t name a OneDrive file/);
  });
});

describe('saving back', () => {
  it('writes over the file only if it is still the version opened, and takes the new version', async () => {
    const calls = stubFetch({ [`PUT ${ITEM}/content`]: () => json(driveItem({ eTag: '"{ABC},2"' })) });
    const next = await connector().source.save!(saved(), origin({ location: 'Documents/Finance/prices.xlsx' }));
    expect(next).toEqual(origin({ version: '"{ABC},2"', location: 'Documents/Finance/prices.xlsx' }));
    expect(calls[0].headers.get('If-Match')).toBe('"{ABC},1"');
    expect(calls[0].headers.get('Content-Type')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(calls[0].body).toBe(BYTES);
  });

  it('throws SaveConflictError when the file changed since it was opened', async () => {
    stubFetch({ [`PUT ${ITEM}/content`]: () => json({ error: { code: 'resourceModified', message: 'ETag does not match' } }, 412) });
    await expect(connector().source.save!(saved(), origin())).rejects.toMatchObject({ name: 'SaveConflictError', message: '' });
  });

  it('writes a large file through an upload session, in chunks sent without the token', async () => {
    const big = new Uint8Array(SIMPLE_UPLOAD_LIMIT + CHUNK_BYTES + 5);
    const upload = 'https://my.microsoftpersonalcontent.com/up/abc';
    const calls = stubFetch({
      [`POST ${ITEM}/createUploadSession`]: () => json({ uploadUrl: upload }),
      [`PUT ${upload}`]: (call) => (call.headers.get('Content-Range')!.endsWith(`-${big.byteLength - 1}/${big.byteLength}`) ? json(driveItem({ eTag: '"{ABC},2"' }), 201) : json({}, 202)),
    });
    const next = await connector().source.save!(saved(big), origin());
    expect(next.version).toBe('"{ABC},2"');
    expect(calls[0].headers.get('If-Match')).toBe('"{ABC},1"');
    expect(JSON.parse(calls[0].body as string)).toEqual({ item: { '@microsoft.graph.conflictBehavior': 'replace' } });
    expect(calls.slice(1).map((c) => [c.headers.get('Content-Range'), c.headers.get('Authorization')])).toEqual([
      [`bytes 0-${CHUNK_BYTES - 1}/${big.byteLength}`, null],
      [`bytes ${CHUNK_BYTES}-${big.byteLength - 1}/${big.byteLength}`, null],
    ]);
  });

  it('a conflict on starting an upload session is a SaveConflictError too', async () => {
    stubFetch({ [`POST ${ITEM}/createUploadSession`]: () => json({ error: { code: 'resourceModified' } }, 412) });
    await expect(connector().source.save!(saved(new Uint8Array(SIMPLE_UPLOAD_LIMIT + 1)), origin())).rejects.toMatchObject({ name: 'SaveConflictError' });
  });

  it('cancels an upload session that fails part way, so no partial file is left', async () => {
    const upload = 'https://my.microsoftpersonalcontent.com/up/abc';
    const calls = stubFetch({
      [`POST ${ITEM}/createUploadSession`]: () => json({ uploadUrl: upload }),
      [`PUT ${upload}`]: () => json({ error: { code: 'serviceNotAvailable', message: 'Try later' } }, 503),
      [`DELETE ${upload}`]: () => new Response(null, { status: 204 }),
    });
    await expect(connector().source.save!(saved(new Uint8Array(SIMPLE_UPLOAD_LIMIT + 1)), origin())).rejects.toMatchObject({ text: { params: { status: 503, reason: 'Try later' } } });
    expect(calls.map((c) => c.method)).toEqual(['POST', 'PUT', 'DELETE']);
  });

  it('keeps the changes, saying why, when the person won’t sign in to save', async () => {
    msal.account = null;
    stubFetch({});
    const { source, confirm } = connector();
    msal.acquireTokenPopup = vi.fn().mockRejectedValue(Object.assign(new Error('blocked'), { errorCode: 'popup_window_error' }));
    confirm.mockResolvedValue(false);
    await expect(source.save!(saved(), origin())).rejects.toMatchObject({ text: { key: 'plugin.page.ruled.onedrive.signInBody' } });
  });

  it('asks to sign in again when the sign-in has expired', async () => {
    msal.silentFails = true;
    const calls = stubFetch({ [`PUT ${ITEM}/content`]: () => json(driveItem({ eTag: '"{ABC},2"' })) });
    const { source, confirm } = connector();
    msal.acquireTokenPopup = vi.fn().mockRejectedValueOnce(Object.assign(new Error('blocked'), { errorCode: 'popup_window_error' })).mockResolvedValue({ account: PERSONAL, accessToken: 'popup-token' });
    await source.save!(saved(), origin());
    expect(confirm).toHaveBeenCalledWith(expect.objectContaining({ body: { key: 'plugin.page.ruled.onedrive.signInAgain', params: { source: source.label } } }));
    expect(calls[0].headers.get('Authorization')).toBe('Bearer popup-token');
  });
});

describe('saving a copy', () => {
  const FOLDER = `${GRAPH}/drives/d1/items/f1`;

  it('writes a new file into the folder picked, never over another', async () => {
    picker.runPicker.mockResolvedValue({ id: 'f1', parentReference: { driveId: 'd1' } });
    const calls = stubFetch({
      [`PUT ${FOLDER}:/prices.xlsx:/content?@microsoft.graph.conflictBehavior=fail`]: () => json(driveItem({ id: 'i2', eTag: '"{DEF},1"', parentReference: { driveId: 'd1', path: '/drive/root:/Archive' } }), 201),
    });
    const made = await connector().source.saveAs!(saved());
    expect(made).toEqual({ source: 'onedrive', name: 'prices.xlsx', ref: { drive: 'd1', item: 'i2' }, version: '"{DEF},1"', writable: true, location: 'Archive/prices.xlsx' });
    expect(picker.runPicker.mock.calls[0][0]).toMatchObject({ mode: 'folders', pickLabel: 'plugin.page.ruled.onedrive.saveHere' });
    expect(calls).toHaveLength(1);
  });

  it('asks for another name when the folder has a file by that one', async () => {
    picker.runPicker.mockResolvedValue({ id: 'f1', parentReference: { driveId: 'd1' } });
    stubFetch({
      [`PUT ${FOLDER}:/prices.xlsx:/content?@microsoft.graph.conflictBehavior=fail`]: () => json({ error: { code: 'nameAlreadyExists', message: 'Name taken' } }, 409),
      [`PUT ${FOLDER}:/prices%20(2).xlsx:/content?@microsoft.graph.conflictBehavior=fail`]: () => json(driveItem({ id: 'i3', name: 'prices (2).xlsx' }), 201),
    });
    const { source, dialog } = connector();
    expect((await source.saveAs!(saved()))?.name).toBe('prices (2).xlsx');
    expect(dialog).toHaveBeenCalledWith(expect.objectContaining({ title: { key: 'plugin.page.ruled.onedrive.nameTakenTitle', params: undefined } }));
  });

  it('resolves null when the person cancels the new name', async () => {
    picker.runPicker.mockResolvedValue({ id: 'f1', parentReference: { driveId: 'd1' } });
    stubFetch({ [`PUT ${FOLDER}:/prices.xlsx:/content?@microsoft.graph.conflictBehavior=fail`]: () => json({ error: { code: 'nameAlreadyExists' } }, 409) });
    const { source, dialog } = connector();
    dialog.mockResolvedValue(null as unknown as string);
    expect(await source.saveAs!(saved())).toBeNull();
  });
});

describe('the rest', () => {
  it('only reads, with readOnly: no saving back or copies, and read permission only', async () => {
    stubFetch({ [`GET ${ITEM}`]: () => json(driveItem()), 'GET https://public.dm.files.1drv.com/download?t=x': () => new Response(BYTES) });
    const { source } = connector({ readOnly: true });
    expect(source.save).toBeUndefined();
    expect(source.saveAs).toBeUndefined();
    const file = await source.pick();
    expect(file!.origin.writable).toBe(false);
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith(expect.objectContaining({ scopes: ['Files.Read.All'] }));
    expect(await (picker.runPicker.mock.calls[0][0] as PickerOptions).token(undefined)).toBe('token:OneDrive.ReadOnly');
  });

  it('signs out of OneDrive in this editor', async () => {
    const { commands, toasts } = connector();
    await commands[0].run({});
    expect(msal.clearCache).toHaveBeenCalledWith({ account: PERSONAL });
    expect(toasts).toEqual([{ kind: 'info', text: { key: 'plugin.page.ruled.onedrive.signedOut', params: { source: { key: 'plugin.page.ruled.onedrive.oneDrive', params: undefined } } } }]);
  });

  it('numbers names for copies', () => {
    expect(nextName('prices.xlsx')).toBe('prices (2).xlsx');
    expect(nextName('prices (2).xlsx')).toBe('prices (3).xlsx');
    expect(nextName('notes')).toBe('notes (2)');
  });
});
