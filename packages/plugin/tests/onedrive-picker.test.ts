import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { runPicker, type PickerOptions } from '../src/onedrive/picker';

type Listener = (event: { source: unknown; data: unknown; ports: MessagePort[] }) => void;

/** The popup the picker page loads in: records the form posted into it. */
function fakePopup() {
  const posted: { action: string; method: string; inputs: Record<string, string> }[] = [];
  const win = {
    closed: false,
    close: vi.fn(() => void (win.closed = true)),
    document: {
      createElement: (tag: string) => {
        const el: Record<string, unknown> = { tag, children: [] as Record<string, unknown>[] };
        el.append = (...kids: Record<string, unknown>[]) => (el.children as unknown[]).push(...kids);
        el.submit = () =>
          posted.push({
            action: el.action as string,
            method: el.method as string,
            inputs: Object.fromEntries((el.children as Record<string, string>[]).map((i) => [i.name, i.value])),
          });
        return el;
      },
      body: { append: vi.fn() },
    },
  };
  return { win, posted };
}

let listener: Listener | undefined;

beforeEach(() => {
  listener = undefined;
  vi.stubGlobal('location', { origin: 'http://localhost:5180' });
  vi.stubGlobal('window', {
    addEventListener: (_: string, fn: Listener) => void (listener = fn),
    removeEventListener: () => void (listener = undefined),
  });
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

/** Starts the picker; once its page is posted, connects to it as the picker page would. */
async function started(options: Partial<PickerOptions> = {}) {
  const { win, posted } = fakePopup();
  const token = vi.fn(async (resource: string | undefined) => `token:${resource ?? 'first'}`);
  const abort = new AbortController();
  const result = runPicker({ win: win as unknown as Window, url: 'https://onedrive.live.com/picker', mode: 'files', locale: 'de-de', token, signal: abort.signal, ...options });
  await vi.waitFor(() => expect(posted).toHaveLength(1));
  const url = new URL(posted[0].action);
  const config = JSON.parse(url.searchParams.get('filePicker')!);
  const channel = new MessageChannel();
  const received: { type: string; id?: string; data?: unknown }[] = [];
  channel.port2.addEventListener('message', (e) => void received.push(e.data));
  channel.port2.start();
  listener!({ source: win, data: { type: 'initialize', channelId: config.messaging.channelId }, ports: [channel.port1] });
  const command = (id: string, data: unknown) => channel.port2.postMessage({ type: 'command', id, data });
  return { win, posted, url, config, token, received, command, result, abort, channel };
}

describe('the File Picker', () => {
  it('posts its configuration and a token into the popup, then activates over the port it is given', async () => {
    const { posted, url, config, received, channel } = await started();
    expect(posted[0]).toMatchObject({ method: 'POST', inputs: { access_token: 'token:first' } });
    expect(url.origin + url.pathname).toBe('https://onedrive.live.com/picker');
    expect(url.searchParams.get('locale')).toBe('de-de');
    expect(config).toMatchObject({
      sdk: '8.0',
      entry: { oneDrive: { files: {} } },
      authentication: {},
      messaging: { origin: 'http://localhost:5180' },
      typesAndSources: { mode: 'files', filters: ['.xlsx', '.xlsm', '.csv', '.tsv'] },
      selection: { mode: 'single' },
    });
    await vi.waitFor(() => expect(received).toEqual([{ type: 'activate' }]));
    channel.port2.close();
  });

  it('answers token requests, and resolves with the item picked, closing the popup', async () => {
    const { win, received, command, result } = await started();
    command('1', { command: 'authenticate', resource: 'https://contoso.sharepoint.com' });
    await vi.waitFor(() => expect(received).toContainEqual({ type: 'result', id: '1', data: { result: 'token', token: 'token:https://contoso.sharepoint.com' } }));
    expect(received).toContainEqual({ type: 'acknowledge', id: '1' });
    command('2', { command: 'pick', items: [{ id: 'i1', parentReference: { driveId: 'd1' } }] });
    expect(await result).toEqual({ id: 'i1', parentReference: { driveId: 'd1' } });
    expect(win.close).toHaveBeenCalled();
  });

  it('says when it can’t get a token, and refuses commands it doesn’t know', async () => {
    const { received, command, token, channel } = await started();
    token.mockRejectedValueOnce(new Error('consent needed'));
    command('1', { command: 'authenticate', resource: 'https://contoso.sharepoint.com' });
    command('2', { command: 'share' });
    await vi.waitFor(() => {
      expect(received).toContainEqual({ type: 'result', id: '1', data: { result: 'error', error: { code: 'unableToObtainToken', message: 'consent needed' } } });
      expect(received).toContainEqual({ type: 'result', id: '2', data: { result: 'error', error: { code: 'unsupportedCommand', message: 'share' } } });
    });
    channel.port2.close();
  });

  it('resolves null when the picker is closed, or its window is', async () => {
    const closing = await started();
    closing.command('1', { command: 'close' });
    expect(await closing.result).toBeNull();

    vi.useFakeTimers();
    const { win, posted } = fakePopup();
    const result = runPicker({ win: win as unknown as Window, url: 'https://x', mode: 'files', locale: 'en-us', token: async () => 't', signal: new AbortController().signal });
    await vi.advanceTimersByTimeAsync(0);
    expect(posted).toHaveLength(1);
    win.closed = true;
    await vi.advanceTimersByTimeAsync(600);
    expect(await result).toBeNull();
  });

  it('ends when the editor goes, and rejects when the first token can’t be had', async () => {
    const { abort, result, win } = await started();
    abort.abort();
    expect(await result).toBeNull();
    expect(win.close).toHaveBeenCalled();

    const popup = fakePopup();
    await expect(
      runPicker({ win: popup.win as unknown as Window, url: 'https://x', mode: 'files', locale: 'en-us', token: async () => Promise.reject(new Error('no')), signal: new AbortController().signal }),
    ).rejects.toThrow('no');
    expect(popup.posted).toEqual([]);
  });

  it('picks folders in a SharePoint site, with its own label on the pick button', async () => {
    const { config, channel } = await started({ mode: 'folders', site: 'https://contoso.sharepoint.com/sites/finance', pickLabel: 'Save here' });
    expect(config).toMatchObject({
      entry: { sharePoint: { byPath: { web: 'https://contoso.sharepoint.com/sites/finance' } } },
      typesAndSources: { mode: 'folders' },
      commands: { pick: { label: 'Save here' } },
    });
    channel.port2.close();
  });

  it('ignores messages from other windows and channels', async () => {
    const { win, received, channel } = await started();
    const other = new MessageChannel();
    listener!({ source: {}, data: { type: 'initialize', channelId: 'x' }, ports: [other.port1] });
    listener!({ source: win, data: { type: 'initialize', channelId: 'other' }, ports: [other.port1] });
    await vi.waitFor(() => expect(received).toEqual([{ type: 'activate' }]));
    other.port1.close();
    channel.port2.close();
  });
});
