/**
 * React hooks for hosts: an editor that lives as long as the component
 * using it, and the editor's state as React state.
 */
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { StoreOptions } from '../state/store';
import { createWorkbookEditor, type EditorHandle } from './handle';
import { applyLiveOptions } from './liveOptions';

/**
 * An editor for a React component: made when the component mounts and
 * destroyed, with its worker, when it unmounts. It's null on the first
 * render, before it's made, so render `<WorkbookEditor>` once it isn't.
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
export function useWorkbookEditor(options: StoreOptions = {}): EditorHandle | null {
  const [editor, setEditor] = useState<EditorHandle | null>(null);
  const latest = useRef(options);
  const applied = useRef(options);
  useEffect(() => {
    latest.current = options;
  });

  useEffect(() => {
    const first = latest.current;
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
