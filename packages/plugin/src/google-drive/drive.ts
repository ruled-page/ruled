/** The Drive v3 REST calls the connector makes, with a token from Google Identity Services. */

export const API = 'https://www.googleapis.com/drive/v3/files';
export const UPLOAD = 'https://www.googleapis.com/upload/drive/v3/files';
export const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12';
export const GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';
export const FOLDER = 'application/vnd.google-apps.folder';

/** Uploads up to this size go in one multipart request; larger ones through a resumable session, as Drive asks. */
export const SIMPLE_UPLOAD_LIMIT = 5 * 1024 * 1024;

const FIELDS = 'id,name,mimeType,size,version,modifiedTime,parents';

/** A file's metadata, as Drive gives it for FIELDS. */
export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  /** Bytes, as a decimal string; Google's own files have none. */
  size?: string;
  /** A number that rises with every change to the file, as a decimal string. */
  version?: string;
  modifiedTime?: string;
  parents?: string[];
}

export type GetToken = () => Promise<string>;

/** A request Drive refused, with its status and Google's reason ("exportSizeLimitExceeded"). */
export class DriveError extends Error {
  constructor(
    readonly status: number,
    readonly reason: string | undefined,
    message: string,
  ) {
    super(message);
    this.name = 'DriveError';
  }
}

async function failure(res: Response): Promise<DriveError> {
  const body = (await res.json().catch(() => undefined)) as { error?: { message?: string; status?: string; errors?: { reason?: string }[] } } | undefined;
  const error = body?.error;
  return new DriveError(res.status, error?.errors?.[0]?.reason ?? error?.status, error?.message ?? res.statusText ?? `HTTP ${res.status}`);
}

/** Fetches from Drive with the token; throws DriveError unless it succeeded. */
async function call(token: GetToken, url: string, init: RequestInit, signal: AbortSignal): Promise<Response> {
  const headers = new Headers(init.headers);
  headers.set('Authorization', `Bearer ${await token()}`);
  const res = await fetch(url, { ...init, headers, signal });
  if (!res.ok) throw await failure(res);
  return res;
}

const fileUrl = (base: string, id: string, query: string) => `${base}/${encodeURIComponent(id)}?${query}&supportsAllDrives=true`;

export async function getFile(token: GetToken, id: string, signal: AbortSignal): Promise<DriveFile> {
  return (await call(token, fileUrl(API, id, `fields=${FIELDS}`), {}, signal)).json();
}

/** The bytes of a file stored in Drive, such as an .xlsx. */
export async function download(token: GetToken, id: string, signal: AbortSignal): Promise<ArrayBuffer> {
  return (await call(token, fileUrl(API, id, 'alt=media'), {}, signal)).arrayBuffer();
}

/** A Google Sheet exported as an .xlsx. Google refuses exports over 10 MB, with the reason "exportSizeLimitExceeded". */
export async function exportXlsx(token: GetToken, id: string, signal: AbortSignal): Promise<ArrayBuffer> {
  return (await call(token, `${API}/${encodeURIComponent(id)}/export?mimeType=${encodeURIComponent(XLSX)}`, {}, signal)).arrayBuffer();
}

const toBlobPart = (bytes: ArrayBuffer | Uint8Array): BlobPart => (bytes instanceof Uint8Array ? (bytes as Uint8Array<ArrayBuffer>) : bytes);

/**
 * Writes bytes as an .xlsx: over the file `id` names, or as a new file
 * with `metadata` (its name and folder). Resolves with the file as saved.
 */
export async function upload(
  token: GetToken,
  { id, metadata, bytes }: { id?: string; metadata: Record<string, unknown>; bytes: ArrayBuffer | Uint8Array },
  signal: AbortSignal,
): Promise<DriveFile> {
  const method = id ? 'PATCH' : 'POST';
  const base = id ? `${UPLOAD}/${encodeURIComponent(id)}` : UPLOAD;
  const query = `fields=${FIELDS}&supportsAllDrives=true`;
  if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    const boundary = `ruled-${Math.random().toString(36).slice(2)}`;
    const body = new Blob([
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: ${XLSX}\r\n\r\n`,
      toBlobPart(bytes),
      `\r\n--${boundary}--`,
    ]);
    const res = await call(token, `${base}?uploadType=multipart&${query}`, { method, headers: { 'Content-Type': `multipart/related; boundary=${boundary}` }, body }, signal);
    return res.json();
  }
  const started = await call(
    token,
    `${base}?uploadType=resumable&${query}`,
    {
      method,
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': XLSX, 'X-Upload-Content-Length': String(bytes.byteLength) },
      body: JSON.stringify(metadata),
    },
    signal,
  );
  const session = started.headers.get('Location');
  if (!session) throw new DriveError(started.status, undefined, 'Google Drive didn’t say where to upload the file.');
  // The session's address is its own authorisation.
  const res = await fetch(session, { method: 'PUT', headers: { 'Content-Type': XLSX }, body: new Blob([toBlobPart(bytes)]), signal });
  if (!res.ok) throw await failure(res);
  return res.json();
}
