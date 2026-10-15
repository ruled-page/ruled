/**
 * The Microsoft Graph calls the connector makes: reading a file, writing it
 * back over itself unless it changed, and writing a new one. Personal
 * OneDrive, OneDrive for work or school and SharePoint libraries all answer
 * the same /drives/{drive}/items/{item} routes.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

/** Files larger than this are written through an upload session, as Graph recommends. */
export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;

/** Upload session chunks must be multiples of 320 KiB; this is 32 of them. */
export const CHUNK_BYTES = 32 * 320 * 1024;

/** A file or folder in a drive, as the picker and the origin name it. */
export interface ItemRef {
  drive: string;
  item: string;
}

/** The parts of a Graph driveItem the connector reads. */
export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  eTag?: string;
  parentReference?: { driveId?: string; path?: string };
  file?: object;
  folder?: object;
  '@microsoft.graph.downloadUrl'?: string;
}

/** Gives a Graph access token, signing in again if it must. */
export type GetToken = () => Promise<string>;

/** A Graph response other than success: its HTTP status and Graph's own error code and message. */
export class GraphError extends Error {
  constructor(
    readonly status: number,
    readonly code: string | undefined,
    message: string,
  ) {
    super(message);
    this.name = 'GraphError';
  }
}

const itemUrl = (ref: ItemRef) => `${GRAPH}/drives/${encodeURIComponent(ref.drive)}/items/${encodeURIComponent(ref.item)}`;

async function failure(response: Response): Promise<GraphError> {
  let code: string | undefined;
  let message = response.statusText || `HTTP ${response.status}`;
  try {
    const body = (await response.json()) as { error?: { code?: string; message?: string } };
    code = body.error?.code;
    if (body.error?.message) message = body.error.message;
  } catch {
    // Not JSON: the status says enough.
  }
  return new GraphError(response.status, code, message);
}

/** A Graph request with the person's token; anything but success throws a GraphError. */
async function graph(token: GetToken, url: string, init: RequestInit = {}, signal?: AbortSignal): Promise<Response> {
  const headers = new Headers(init.headers);
  headers.set('Authorization', `Bearer ${await token()}`);
  const response = await fetch(url, { ...init, headers, signal });
  if (!response.ok) throw await failure(response);
  return response;
}

export async function getItem(token: GetToken, ref: ItemRef, signal?: AbortSignal): Promise<DriveItem> {
  return (await graph(token, itemUrl(ref), {}, signal)).json() as Promise<DriveItem>;
}

/** The person's own drive, whose web address says where the picker for a work or school account lives. */
export async function myDrive(token: GetToken, signal?: AbortSignal): Promise<{ webUrl: string }> {
  return (await graph(token, `${GRAPH}/me/drive?$select=webUrl`, {}, signal)).json() as Promise<{ webUrl: string }>;
}

/**
 * The file's bytes. The item's short-lived download address needs no token,
 * so the browser fetches it without the preflight a redirect from …/content
 * with an Authorization header would need.
 */
export async function download(token: GetToken, ref: ItemRef, item: DriveItem, signal?: AbortSignal): Promise<ArrayBuffer> {
  const url = item['@microsoft.graph.downloadUrl'];
  if (!url) return (await graph(token, `${itemUrl(ref)}/content`, {}, signal)).arrayBuffer();
  const response = await fetch(url, { signal });
  if (!response.ok) throw await failure(response);
  return response.arrayBuffer();
}

const contentType = (name: string) =>
  /\.xlsm$/i.test(name) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Writes `bytes` over the file, only if it is still at version `eTag`:
 * otherwise Graph answers 412, thrown as a GraphError with that status.
 */
export async function replace(token: GetToken, ref: ItemRef, name: string, bytes: Uint8Array, eTag: string | undefined, signal?: AbortSignal): Promise<DriveItem> {
  const ifMatch: Record<string, string> = eTag ? { 'If-Match': eTag } : {};
  if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    const response = await graph(token, `${itemUrl(ref)}/content`, { method: 'PUT', headers: { ...ifMatch, 'Content-Type': contentType(name) }, body: bytes as BodyInit }, signal);
    return response.json() as Promise<DriveItem>;
  }
  return uploadSession(token, `${itemUrl(ref)}/createUploadSession`, { '@microsoft.graph.conflictBehavior': 'replace' }, ifMatch, bytes, signal);
}

/**
 * Writes a new file called `name` into the folder. When the folder already
 * has one by that name, Graph answers 409, thrown as a GraphError with that status.
 */
export async function create(token: GetToken, folder: ItemRef, name: string, bytes: Uint8Array, signal?: AbortSignal): Promise<DriveItem> {
  const path = `${itemUrl(folder)}:/${encodeURIComponent(name)}:`;
  if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    const url = `${path}/content?@microsoft.graph.conflictBehavior=fail`;
    return (await graph(token, url, { method: 'PUT', headers: { 'Content-Type': contentType(name) }, body: bytes as BodyInit }, signal)).json() as Promise<DriveItem>;
  }
  return uploadSession(token, `${path}/createUploadSession`, { '@microsoft.graph.conflictBehavior': 'fail', name }, {}, bytes, signal);
}

/**
 * A large file, written in chunks. The session's own address carries its
 * authorisation, so the chunks are sent without the token. A session that
 * fails part way is cancelled, so no partial file is left behind.
 */
async function uploadSession(token: GetToken, url: string, item: Record<string, string>, headers: Record<string, string>, bytes: Uint8Array, signal?: AbortSignal): Promise<DriveItem> {
  const session = (await (await graph(token, url, { method: 'POST', headers: { ...headers, 'Content-Type': 'application/json' }, body: JSON.stringify({ item }) }, signal)).json()) as { uploadUrl: string };
  try {
    for (let start = 0; ; start += CHUNK_BYTES) {
      const end = Math.min(start + CHUNK_BYTES, bytes.byteLength);
      const response = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Range': `bytes ${start}-${end - 1}/${bytes.byteLength}` },
        body: bytes.subarray(start, end) as BodyInit,
        signal,
      });
      if (!response.ok) throw await failure(response);
      if (end === bytes.byteLength) return response.json() as Promise<DriveItem>;
    }
  } catch (err) {
    void fetch(session.uploadUrl, { method: 'DELETE' }).catch(() => undefined);
    throw err;
  }
}

/** "Documents/Finance" from a parent path such as "/drive/root:/Documents/Finance"; undefined for an item outside the person's folders. */
export function folderPath(item: DriveItem): string | undefined {
  const path = item.parentReference?.path;
  if (!path) return undefined;
  const inside = decodeURIComponent(path.replace(/^\/drives?(\/[^/]+)?\/root:?/, '')).replace(/^\//, '');
  return inside || undefined;
}
