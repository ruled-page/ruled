/**
 * React hooks for hosts: an editor that lives as long as the component
 * using it, and the editor's state as React state.
 */
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { StoreOptions } from '../state/store';
import { createWorkbookEditor, type EditorHandle } from './handle';
import { applyLiveOptions } from './liveOptions';

/** A workbook for `useWorkbookEditor` to open: its name, which says what kind of file it is, and its bytes. */
export interface WorkbookFile {
  name: string;
  bytes: ArrayBuffer | Uint8Array;
}

export interface WorkbookEditorHookOptions extends StoreOptions {
  /**
   * The workbook to open, once the editor is made and again whenever it's a
   * different file (another name or other bytes; a new object for the same
   * file isn't). Over unsaved changes the user is asked first, and the
   * workbook they keep stays open. A file that can't be opened says why in
   * the editor. Leave it out to open files yourself with `open()`.
   */
  file?: WorkbookFile | null;
}

/**
 * An editor for a React component: made when the component mounts and
 * destroyed, with its worker, when it unmounts. It's null on the first
 * render, before it's made, so render `<WorkbookEditor>` once it isn't.
 *
 * - `file` is opened when the editor is ready, and again when it changes.
 *
 * - Theme, language, user, schemas and custom lists follow `options` as they
 *   change, without closing the workbook. A new array or object with the same
 *   entries isn't a change, so they can be written inline.
 * - `onSave`, `onEvent`, `onError` and `onLinkClick` can be new functions on
 *   every render: the editor calls the latest. Whether there's an `onSave` at
 *   all is read when the editor is made.
 * - Every other option is read when the editor is made. To change one, give
 *   the component a new `key`, which makes a new editor.
 */
export function useWorkbookEditor(options: WorkbookEditorHookOptions = {}): EditorHandle | null {
  const [editor, setEditor] = useState<EditorHandle | null>(null);
  const latest = useRef(options);
  const applied = useRef(options);
  useEffect(() => {
    latest.current = options;
  });

  useEffect(() => {
    const { file: _file, ...first } = latest.current;
    const call = latest;
    const made = createWorkbookEditor({
      ...first,
      // The latest callbacks, so they needn't be stable.
      ...(first.onSave && { onSave: (file) => call.current.onSave?.(file) }),
      ...(first.onEvent && { onEvent: (event) => call.current.onEvent?.(event) }),
      ...(first.onError && { onError: (error) => call.current.onError?.(error) }),
      ...(first.onLinkClick && { onLinkClick: (link) => call.current.onLinkClick?.(link) }),
    });
    applied.current = first;
    setEditor(made);
    return () => {
      made.destroy();
      setEditor((now) => (now === made ? null : now));
    };
  }, []);

  useEffect(() => {
    if (!editor) return;
    applyLiveOptions(editor, applied.current, options);
    applied.current = options;
  });

  // The file, once per editor and file: not again for a new object holding the same one, or when StrictMode runs effects twice.
  const opened = useRef<{ editor: EditorHandle; file: WorkbookFile } | null>(null);
  const { file } = options;
  useEffect(() => {
    if (!editor || !file) return;
    const was = opened.current;
    if (was && was.editor === editor && was.file.name === file.name && was.file.bytes === file.bytes) return;
    opened.current = { editor, file };
    // The editor shows why a file can't be opened.
    editor.open(file.bytes, file.name).catch(() => undefined);
  }, [editor, file]);

  return editor;
}

/** What `useEditorState` gives: the editor's status, its workbook's name, and its edits. */
export interface EditorState {
  status: EditorHandle['status'];
  fileName: string;
  /** Cells that differ from the file as it was opened. */
  editedCount: number;
  /** Whether there are edits the host hasn't received through a save. */
  unsaved: boolean;
}

const NONE: EditorState = { status: 'empty', fileName: '', editedCount: 0, unsaved: false };

/** An editor's state, re-rendering the component as it changes: for a host's own unsaved badge, title or Save button. */
export function useEditorState(editor: EditorHandle | null): EditorState {
  const last = useRef(NONE);
  const subscribe = useCallback((changed: () => void) => editor?.subscribe(changed) ?? (() => {}), [editor]);
  const read = useCallback((): EditorState => {
    if (!editor) return NONE;
    const was = last.current;
    const { status, fileName, editedCount, unsaved } = editor;
    if (status === was.status && fileName === was.fileName && editedCount === was.editedCount && unsaved === was.unsaved) return was;
    return (last.current = { status, fileName, editedCount, unsaved });
  }, [editor]);
  return useSyncExternalStore(subscribe, read, read);
}
