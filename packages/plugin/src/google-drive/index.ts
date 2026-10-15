/**
 * Open workbooks from Google Drive, save them back, and save copies there:
 * `@ruled/plugin/google-drive`. Google Sheets open as an Excel copy, which
 * can only be saved as a new .xlsx; the Google Sheet is never written over.
 *
 * ```ts
 * import { googleDrive } from '@ruled/plugin/google-drive';
 * const editor = useWorkbookEditor({
 *   plugins: [googleDrive({ clientId: 'your-client-id.apps.googleusercontent.com', apiKey: 'your-api-key', appId: 'your-project-number' })],
 * });
 * ```
 */
import type { FileSource, Localized, RuledPlugin } from '@ruled/sheet/plugin';
import type { Connector } from './connector';
import { GOOGLE_DRIVE_PLUGIN_ID, GOOGLE_DRIVE_SOURCE_ID, msg, translations } from './messages';

export { GOOGLE_DRIVE_PLUGIN_ID, GOOGLE_DRIVE_SOURCE_ID };

export interface GoogleDriveOptions {
  /** The OAuth client id of a web application client in your Google Cloud project: "….apps.googleusercontent.com". */
  clientId: string;
  /** An API key from the same project, restricted to your sites and the Picker API; the Picker needs it. */
  apiKey: string;
  /**
   * Your Google Cloud project's number. The Picker needs it so that files
   * people choose there can be opened and saved with the `drive.file`
   * permission, which covers only those files.
   */
  appId: string;
  /** The name shown in the Open menu and on "Save to …". Default "Google Drive". */
  label?: Localized;
}

/** The connector, to pass in the editor's `plugins`. */
export function googleDrive(options: GoogleDriveOptions): RuledPlugin {
  for (const key of ['clientId', 'apiKey', 'appId'] as const) {
    if (!options?.[key]) throw new Error(`googleDrive() needs the ${key} from your Google Cloud project.`);
  }
  return {
    id: GOOGLE_DRIVE_PLUGIN_ID,
    apiVersion: 1,
    translations,
    setup: (ctx) => setup(ctx, options),
  };
}

/**
 * Registers the source and the sign-out command straight away, and loads
 * the rest of the connector when they are first used: as a menu holding
 * them opens, which also loads Google's scripts.
 */
function setup(ctx: Parameters<RuledPlugin['setup']>[0], o: GoogleDriveOptions): void {
  const label = o.label ?? msg('googleDrive');
  let signedIn = false;
  const onSignIn = (now: boolean) => {
    signedIn = now;
    ctx.commands.refresh();
  };
  let loading: Promise<Connector> | undefined;
  const connector = () =>
    (loading ??= import('./connector').then(
      (m) => m.connect(ctx, o, label, onSignIn),
      (err: unknown) => {
        loading = undefined;
        throw err;
      },
    ));
  const preload = () => void connector().then((c) => c.preload(), () => undefined);

  ctx.files.registerSource({
    id: GOOGLE_DRIVE_SOURCE_ID,
    label,
    preload,
    pick: () => connector().then((c) => c.pick()),
    fetch: (origin) => connector().then((c) => c.fetch(origin)),
    save: (file, origin) => connector().then((c) => c.save(file, origin)),
    saveAs: (file) => connector().then((c) => c.saveAs(file)),
  } satisfies FileSource);

  ctx.commands.register({
    id: 'signOut',
    label: msg('signOut', { source: label }),
    where: 'more',
    // Only while the editor holds a token for Drive: there is nothing to sign out of otherwise.
    visible: () => signedIn,
    preload,
    run: () => connector().then((c) => c.signOut()),
  });
}
