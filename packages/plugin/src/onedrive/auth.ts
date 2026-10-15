import type { AccountInfo, PublicClientApplication } from '@azure/msal-browser';

/** Personal Microsoft accounts all belong to this tenant. */
const CONSUMER_TENANT = '9188040d-6c67-4c5b-b112-36a304b66dad';

export interface AuthOptions {
  clientId: string;
  authority: string;
  redirectUri: string;
}

/** Signing in to Microsoft with MSAL's popups; MSAL keeps its cache in sessionStorage, so a sign-in lasts as long as the tab. */
export interface Auth {
  /** The account signed in, if any. */
  account(): AccountInfo | null;
  /** Whether the account is a personal Microsoft account, rather than a work or school one. */
  consumer(): boolean;
  /** Opens the sign-in popup at once, so call it inside a click. Resolves false when the person closes it. */
  signIn(scopes: string[]): Promise<boolean>;
  /** A token without asking the person; throws NeedsSignIn when only a popup can get one. */
  token(scopes: string[]): Promise<string>;
  /** A token through a popup, opened at once. Resolves null when the person closes it. */
  tokenPopup(scopes: string[]): Promise<string | null>;
  /** Forgets the account in this editor, without signing it out of Microsoft everywhere else. */
  signOut(): Promise<void>;
}

/** Only a popup can get the token now: the sign-in expired, or consent is needed. */
export class NeedsSignIn extends Error {
  constructor(cause: unknown) {
    super((cause as Error)?.message ?? 'Sign-in needed', { cause });
    this.name = 'NeedsSignIn';
  }
}

const code = (err: unknown) => (err as { errorCode?: string } | null)?.errorCode;

/** The browser blocked a popup: it wasn't opened inside a click. */
export const isPopupBlocked = (err: unknown) => code(err) === 'popup_window_error' || code(err) === 'empty_window_error' || (err as Error)?.name === 'PopupBlocked';

/** The person closed the popup or cancelled. */
const isCancelled = (err: unknown) => code(err) === 'user_cancelled';

/** "https://login.microsoftonline.com/common" from "common", a tenant id or a domain; a full address is used as it is. */
export const authorityUrl = (authority: string) => (authority.startsWith('https://') ? authority : `https://login.microsoftonline.com/${authority}`);

/** Loads MSAL, which is only fetched when someone first reaches for OneDrive. */
export async function loadAuth(options: AuthOptions): Promise<Auth> {
  const msal = await import('@azure/msal-browser');
  const app: PublicClientApplication = new msal.PublicClientApplication({
    auth: { clientId: options.clientId, authority: authorityUrl(options.authority), redirectUri: options.redirectUri },
    cache: { cacheLocation: 'sessionStorage' },
  });
  await app.initialize();
  const account = () => app.getActiveAccount() ?? app.getAllAccounts()[0] ?? null;

  return {
    account,
    consumer: () => account()?.tenantId === CONSUMER_TENANT,
    async signIn(scopes) {
      try {
        const result = await app.loginPopup({ scopes, prompt: 'select_account' });
        app.setActiveAccount(result.account);
        return true;
      } catch (err) {
        if (isCancelled(err)) return false;
        throw err;
      }
    },
    async token(scopes) {
      const signedIn = account();
      if (!signedIn) throw new NeedsSignIn(undefined);
      try {
        return (await app.acquireTokenSilent({ scopes, account: signedIn })).accessToken;
      } catch (err) {
        if (err instanceof msal.InteractionRequiredAuthError || code(err) === 'no_account_error') throw new NeedsSignIn(err);
        throw err;
      }
    },
    async tokenPopup(scopes) {
      try {
        const result = await app.acquireTokenPopup({ scopes, account: account() ?? undefined });
        app.setActiveAccount(result.account);
        return result.accessToken;
      } catch (err) {
        if (isCancelled(err)) return null;
        throw err;
      }
    },
    async signOut() {
      const signedIn = account();
      if (signedIn) await app.clearCache({ account: signedIn });
      app.setActiveAccount(null);
    },
  };
}
