import { deleteDraft, fileKey } from '../state/drafts';

/**
 * Deletes the draft the editor kept in this browser for a file, given the
 * file's bytes: the unsaved edits it would offer back when the file is
 * opened. For hosts that never open the file in an editor; with one open,
 * use its `discardDraft()`. Resolves once done, also when the file has no
 * draft or the browser's storage is unavailable.
 */
export async function discardDraft(bytes: ArrayBuffer | Uint8Array): Promise<void> {
  const buffer = bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes;
  let key: string;
  try {
    key = await fileKey(buffer);
  } catch {
    return;
  }
  await deleteDraft(key);
}
