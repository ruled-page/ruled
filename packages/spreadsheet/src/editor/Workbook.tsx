import type { SaveAnswer } from '../state/store';
import { useWorkbookEditor, type WorkbookEditorHookOptions } from './hooks';
import { plainSaved, type SavedWorkbookData } from './plainData';
import { WorkbookEditor, type WorkbookEditorProps } from './WorkbookEditor';

export interface WorkbookProps extends Omit<WorkbookEditorHookOptions, 'onSave'>, Omit<WorkbookEditorProps, 'editor'> {
  /**
   * Receives the workbook when the user saves, as plain data that can be
   * sent to a server, so a Server Component can pass a Server Action here.
   * Answer `{ error }` or `{ conflict }` rather than throwing (see
   * SaveAnswer): an error thrown on a server doesn't reach the page as written.
   */
  onSave?: (file: SavedWorkbookData) => SaveAnswer | Promise<SaveAnswer>;
}

/**
 * The editor in one component whose props can come from a React Server
 * Component, as in Next.js's App Router: plain data such as the file's name
 * and bytes, the theme, features and language, and a Server Action as
 * `onSave`. It makes its editor as useWorkbookEditor does, once it's in the
 * browser; rendered on the server, and until then, it's the editor's empty
 * box. Options that are functions (onEvent, onError, onLinkClick,
 * createWorker, and schemas with code) can't come from a Server Component:
 * render it from a client component of your own to pass them.
 */
export function Workbook({ emptyState, toolbarEnd, className, onSave, ...options }: WorkbookProps) {
  const editor = useWorkbookEditor({ ...options, ...(onSave && { onSave: (saved) => onSave(plainSaved(saved)) }) });
  if (!editor) return <div className={`wbe wbe-root ${className ?? ''}`.trim()} aria-busy="true" />;
  return <WorkbookEditor editor={editor} emptyState={emptyState} toolbarEnd={toolbarEnd} className={className} />;
}
