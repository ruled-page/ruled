/**
 * Google's scripts, loaded at run time: Identity Services for tokens, and
 * the Picker. Neither is on npm, so they are fetched from Google the first
 * time someone reaches for Google Drive.
 */

export const GIS_SCRIPT = 'https://accounts.google.com/gsi/client';
export const GAPI_SCRIPT = 'https://apis.google.com/js/api.js';
export const DRIVE_FILE_SCOPE = 'https://www.googleapis.com/auth/drive.file';

interface TokenResponse {
  access_token?: string;
  expires_in?: number | string;
  error?: string;
  error_description?: string;
}

interface TokenClient {
  requestAccessToken(overrides?: { prompt?: string }): void;
}

interface PickerDoc {
  id: string;
  name: string;
  mimeType: string;
  parentId?: string;
}

/** The parts of Google's globals the connector uses. */
export interface GoogleGlobals {
  accounts: {
    oauth2: {
      initTokenClient(config: {
        client_id: string;
        scope: string;
        callback: (response: TokenResponse) => void;
        error_callback?: (error: { type: string; message?: string }) => void;
      }): TokenClient;
      hasGrantedAllScopes?(response: TokenResponse, scope: string): boolean;
      revoke(token: string, done?: () => void): void;
    };
  };
  // The Picker's builder is a chain of setters; typing each one adds nothing here.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  picker: any;
}

/** Google's globals, once its scripts have run; not declared on Window, so hosts' types are left alone. */
const globals = () =>
  window as unknown as { google?: GoogleGlobals; gapi?: { load(name: string, callbacks: { callback: () => void; onerror?: () => void }): void } };

/** Adds a script to the page once, resolving when it has run. */
function script(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const el = document.createElement('script');
    el.src = src;
    el.async = true;
    el.onload = () => resolve();
    el.onerror = () => {
      el.remove();
      reject(new Error(`Couldn’t load ${src}.`));
    };
    document.head.append(el);
  });
}

/** Loads Identity Services and the Picker, unless the page has them already. */
export async function loadGoogle(): Promise<GoogleGlobals> {
  const w = globals();
  await Promise.all([
    w.google?.accounts?.oauth2 ? undefined : script(GIS_SCRIPT),
    w.google?.picker
      ? undefined
      : (w.gapi ? Promise.resolve() : script(GAPI_SCRIPT)).then(
          () => new Promise<void>((resolve, reject) => w.gapi!.load('picker', { callback: resolve, onerror: () => reject(new Error('Couldn’t load the Google Picker.')) })),
        ),
  ]);
  return w.google!;
}

/** The browser blocked Google's popup: it wasn't opened inside a click. */
export const isPopupBlocked = (err: unknown) => (err as Error)?.name === 'PopupBlocked';

/** Signing in to Google: an access token for `drive.file`, kept in memory only. */
export interface Auth {
  /** The token, while it has more than a minute left. */
  current(): string | undefined;
  /** Opens Google's popup at once, so call it inside a click. Resolves null when the person closes it; throws NotAllowed when they don't allow access. */
  request(): Promise<string | null>;
  /** Forgets the token after Drive refused it. */
  forget(): void;
  /** Revokes the access this editor was given, and forgets the token. */
  signOut(): Promise<void>;
}

export function createAuth(google: GoogleGlobals, clientId: string): Auth {
  let token: { value: string; expires: number } | undefined;
  let granted = false;
  let pending: { resolve: (token: string | null) => void; reject: (err: Error) => void } | undefined;
  const settle = (fn: (p: NonNullable<typeof pending>) => void) => {
    const p = pending;
    pending = undefined;
    if (p) fn(p);
  };

  const client = google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: DRIVE_FILE_SCOPE,
    callback(response) {
      // Declining, or leaving Drive's box unticked on the consent screen, gives no token for drive.file.
      const allowed = !!response.access_token && !response.error && (google.accounts.oauth2.hasGrantedAllScopes?.(response, DRIVE_FILE_SCOPE) ?? true);
      if (!allowed) return settle((p) => p.reject(Object.assign(new Error('Access to Google Drive wasn’t allowed.'), { name: 'NotAllowed' })));
      granted = true;
      token = { value: response.access_token!, expires: Date.now() + Number(response.expires_in ?? 3600) * 1000 };
      settle((p) => p.resolve(token!.value));
    },
    error_callback(error) {
      if (error.type === 'popup_failed_to_open') {
        settle((p) => p.reject(Object.assign(new Error('The browser blocked Google’s sign-in window.'), { name: 'PopupBlocked' })));
      } else if (error.type === 'popup_closed') {
        settle((p) => p.resolve(null));
      } else {
        settle((p) => p.reject(new Error(error.message || 'Google sign-in failed.')));
      }
    },
  });

  return {
    current: () => (token && token.expires - Date.now() > 60_000 ? token.value : undefined),
    request() {
      // A request still waiting never hears back once another starts.
      settle((p) => p.resolve(null));
      return new Promise((resolve, reject) => {
        pending = { resolve, reject };
        // After the first grant, Google skips the consent screen where it can.
        client.requestAccessToken(granted ? { prompt: '' } : {});
      });
    },
    forget() {
      token = undefined;
    },
    async signOut() {
      const value = token?.value;
      token = undefined;
      granted = false;
      if (value) await new Promise<void>((resolve) => google.accounts.oauth2.revoke(value, resolve));
    },
  };
}

/** The file or folder chosen in the Picker. */
export type Picked = PickerDoc;

export interface PickerOptions {
  token: string;
  apiKey: string;
  appId: string;
  mode: 'files' | 'folders';
  /** The Picker's language, such as "de". */
  locale: string;
  title: string;
  /** The folder to start in. */
  parent?: string;
  mimeTypes: string[];
}

/** Shows Google's Picker, a frame over the page; resolves null when the person closes it. */
export function showPicker(google: GoogleGlobals, o: PickerOptions): Promise<Picked | null> {
  const p = google.picker;
  const view =
    o.mode === 'files'
      ? new p.DocsView(p.ViewId.DOCS).setMimeTypes(o.mimeTypes.join(',')).setIncludeFolders(true).setMode(p.DocsViewMode.LIST)
      : new p.DocsView(p.ViewId.FOLDERS).setMimeTypes(o.mimeTypes.join(',')).setIncludeFolders(true).setSelectFolderEnabled(true);
  if (o.parent) view.setParent(o.parent);
  return new Promise((resolve) => {
    new p.PickerBuilder()
      .setAppId(o.appId)
      .setDeveloperKey(o.apiKey)
      .setOAuthToken(o.token)
      .setLocale(o.locale)
      .setTitle(o.title)
      .enableFeature(p.Feature.SUPPORT_DRIVES)
      .addView(view)
      .setCallback((data: Record<string, unknown>) => {
        const action = data[p.Response.ACTION];
        if (action === p.Action.PICKED) resolve((data[p.Response.DOCUMENTS] as PickerDoc[])[0] ?? null);
        else if (action === p.Action.CANCEL) resolve(null);
      })
      .build()
      .setVisible(true);
  });
}
