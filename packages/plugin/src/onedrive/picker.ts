/**
 * Microsoft's File Picker v8, in a popup: the picker page is posted the
 * configuration and a token, then talks to this page over a MessagePort,
 * asking for tokens and handing back what was picked.
 */

/** Personal OneDrive's picker; work and school accounts use their SharePoint's /_layouts/15/FilePicker.aspx. */
export const CONSUMER_PICKER = 'https://onedrive.live.com/picker';

/** The parts of a picked item the connector needs; the picker always gives these. */
export interface PickedItem {
  id: string;
  name?: string;
  parentReference: { driveId: string };
}

export interface PickerOptions {
  /** The popup to show the picker in, opened inside the click. */
  win: Window;
  /** The picker page: CONSUMER_PICKER, or a SharePoint web's FilePicker.aspx. */
  url: string;
  /** For a SharePoint site rather than the person's OneDrive: the site's address. */
  site?: string;
  mode: 'files' | 'folders';
  /** The picker's label for its pick button, in the editor's language. */
  pickLabel?: string;
  /** "en-us", "de-de": the picker's language. */
  locale: string;
  /** A token for the resource the picker names (a SharePoint origin), or for the first page load when it names none. */
  token(resource: string | undefined): Promise<string>;
  signal: AbortSignal;
}

/** The files the editor opens; the picker hides the rest. */
const FILE_TYPES = ['.xlsx', '.xlsm', '.csv', '.tsv'];

/** Resolves with the item picked, or null when the picker is closed or cancelled. */
export function runPicker(o: PickerOptions): Promise<PickedItem | null> {
  const { win } = o;
  const channelId = crypto.randomUUID();
  const config = {
    sdk: '8.0',
    entry: o.site ? { sharePoint: { byPath: { web: o.site } } } : { oneDrive: { files: {} } },
    authentication: {},
    messaging: { origin: location.origin, channelId },
    typesAndSources: o.mode === 'files' ? { mode: 'files', filters: FILE_TYPES } : { mode: 'folders' },
    selection: { mode: 'single' },
    ...(o.pickLabel ? { commands: { pick: { label: o.pickLabel } } } : {}),
  };

  return new Promise<PickedItem | null>((resolve, reject) => {
    let port: MessagePort | undefined;
    let done = false;
    const finish = (result: PickedItem | null, err?: unknown) => {
      if (done) return;
      done = true;
      window.removeEventListener('message', onWindowMessage);
      o.signal.removeEventListener('abort', onAbort);
      clearInterval(watch);
      port?.close();
      if (!win.closed) win.close();
      if (err) reject(err);
      else resolve(result);
    };
    const onAbort = () => finish(null);
    // Closing the popup with the browser's own button sends no message.
    const watch = setInterval(() => win.closed && finish(null), 500);

    const reply = (id: string, data: unknown) => port?.postMessage({ type: 'result', id, data });
    const onCommand = async (id: string, command: { command: string; resource?: string; items?: PickedItem[] }) => {
      port?.postMessage({ type: 'acknowledge', id });
      switch (command.command) {
        case 'authenticate':
          try {
            reply(id, { result: 'token', token: await o.token(command.resource) });
          } catch (err) {
            reply(id, { result: 'error', error: { code: 'unableToObtainToken', message: (err as Error)?.message ?? String(err) } });
          }
          break;
        case 'pick': {
          const item = command.items?.[0];
          reply(id, { result: 'success' });
          finish(item ?? null);
          break;
        }
        case 'close':
          finish(null);
          break;
        default:
          reply(id, { result: 'error', error: { code: 'unsupportedCommand', message: command.command } });
      }
    };
    const onPortMessage = (event: MessageEvent) => {
      const message = event.data as { type?: string; id?: string; data?: { command: string } };
      if (message?.type === 'command' && message.id && message.data) void onCommand(message.id, message.data);
    };
    // The picker says it's ready with "initialize" and a port, again whenever it reloads.
    function onWindowMessage(event: MessageEvent) {
      if (event.source !== win) return;
      const message = event.data as { type?: string; channelId?: string };
      if (message?.type !== 'initialize' || message.channelId !== channelId || !event.ports[0]) return;
      port?.close();
      port = event.ports[0];
      port.addEventListener('message', onPortMessage);
      port.start();
      port.postMessage({ type: 'activate' });
    }

    window.addEventListener('message', onWindowMessage);
    o.signal.addEventListener('abort', onAbort);

    // The picker page is loaded by posting a form into the popup, with the configuration and a token.
    o.token(undefined).then((token) => {
      if (done) return;
      const doc = win.document;
      const form = doc.createElement('form');
      form.method = 'POST';
      form.action = `${o.url}?${new URLSearchParams({ filePicker: JSON.stringify(config), locale: o.locale })}`;
      const input = doc.createElement('input');
      input.type = 'hidden';
      input.name = 'access_token';
      input.value = token;
      form.append(input);
      doc.body.append(form);
      form.submit();
    }, (err: unknown) => finish(null, err));
  });
}
