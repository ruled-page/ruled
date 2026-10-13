import { useEffect, useId, useLayoutEffect, useRef, useState, type FormEvent, type KeyboardEvent, type RefObject } from 'react';
import { toA1 } from '@ruled/formula/refs';
import type { CellNote, NoteEntry } from '../core/types';
import type { WorkbookStore } from '../state/store';
import { commandKey, mod } from './keys';
import { noteDate } from './NoteTip';
import { useDialog } from './useDialog';
import { useT } from '../i18n/react';

type At = NonNullable<WorkbookStore['commentFor']>;

/** The control to start at, focused as the popup (or a part of it the user asked for) appears. */
function useFocusFirst<T extends HTMLElement>(): RefObject<T | null> {
  const ref = useRef<T>(null);
  useLayoutEffect(() => ref.current?.focus({ preventScroll: true }), []);
  return ref;
}

/** Ctrl+Enter (⌘Enter) in a note's or comment's text sends it, as in Excel; Enter starts a new line. */
function sendOnModEnter(e: KeyboardEvent<HTMLTextAreaElement>): void {
  if (e.key === 'Enter' && commandKey(e)) {
    e.preventDefault();
    e.currentTarget.form?.requestSubmit();
  }
}

/**
 * A cell's note, to add, change or delete, or its comment thread, to read,
 * reply to, resolve or delete: beside the cell, opened from the cell menu or
 * with Shift+F2. Where the cell's notes are read-only it says why and shows
 * them only.
 */
export function CommentPopup({ store }: { store: WorkbookStore }) {
  const t = useT();
  const at = store.commentFor;
  const view = at ? store.sheets[at.sheet] : undefined;
  const ref = useRef<HTMLDivElement>(null);
  const id = useId();
  const [pos, setPos] = useState({ left: 0, top: 0 });
  const close = () => store.set('commentFor', undefined);
  useDialog(ref, !!at, close, { modal: true });

  // Beside the cell, but on screen.
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || !at) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(at.x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(at.y, window.innerHeight - r.height - 4)) });
  }, [at]);

  useEffect(() => {
    if (!at) return;
    const away = (e: Event) => {
      if (!ref.current?.contains(e.target as Node)) store.set('commentFor', undefined);
    };
    window.addEventListener('mousedown', away, true);
    return () => window.removeEventListener('mousedown', away, true);
  }, [store, at]);

  if (!at || !view) return null;
  const note = view.noteAt(at.row, at.col);
  const reason = view.commentReadOnlyReason(at.row, at.col);
  const thread = note ? note.kind === 'thread' : at.start === 'thread';
  const cell = toA1(at.row, at.col);
  const title = note ? t(thread ? 'comments.threadTitle' : 'comments.noteTitle', { cell }) : t(thread ? 'comments.newThreadTitle' : 'comments.newNoteTitle', { cell });
  const props = { store, at, note, locked: !!reason, close };
  return (
    <div ref={ref} className="wbe-comment-popup" style={pos} role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1}>
      <h2 id={`${id}-title`} className="wbe-comment-title">
        {title}
      </h2>
      {reason && <p className="wbe-comment-meta">{t.text(reason)}</p>}
      {thread ? <Thread key={`${at.sheet}:${cell}`} {...props} dialog={ref} /> : <Note key={`${at.sheet}:${cell}`} {...props} />}
    </div>
  );
}

interface Props {
  store: WorkbookStore;
  at: At;
  note: CellNote | undefined;
  locked: boolean;
  close: () => void;
}

/** A note's text to change, or a new note's to write. */
function Note({ store, at, note, locked, close }: Props) {
  const t = useT();
  const id = useId();
  const was = note?.entries[0]?.text ?? '';
  const author = note?.entries[0]?.author;
  const [text, setText] = useState(was);
  const input = useRef<HTMLTextAreaElement>(null);
  const closeButton = useFocusFirst<HTMLButtonElement>();
  // Typing carries on from the end of the note.
  useLayoutEffect(() => {
    input.current?.focus({ preventScroll: true });
    input.current?.setSelectionRange(was.length, was.length);
  }, [was]);

  const by = author && <p className="wbe-comment-meta">{t('comments.by', { author })}</p>;
  if (locked) {
    return (
      <>
        {by}
        <p className="wbe-comment-text">{was}</p>
        <div className="wbe-modal-actions">
          <button ref={closeButton} type="button" className="wbe-btn" onClick={close}>
            {t('comments.close')}
          </button>
        </div>
      </>
    );
  }
  const changed = text.trim() !== '' && text !== was;
  const save = async (e: FormEvent) => {
    e.preventDefault();
    if (changed && (await store.setNote(at.row, at.col, text))) close();
  };
  const remove = async () => {
    if (await store.deleteComment(at.row, at.col)) close();
  };
  return (
    <form onSubmit={(e) => void save(e)}>
      {by}
      <label className="wbe-field">
        <span>{t('comments.noteLabel')}</span>
        <textarea
          ref={input}
          className="wbe-comment-input"
          value={text}
          rows={4}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={sendOnModEnter}
          aria-describedby={`${id}-hint`}
        />
      </label>
      <p id={`${id}-hint`} className="wbe-comment-meta">
        {t('comments.sendHint', { key: mod(t, 'Enter') })}
      </p>
      <div className="wbe-modal-actions">
        {note && (
          <button type="button" className="wbe-btn wbe-comment-delete" onClick={() => void remove()}>
            {t('comments.deleteNote')}
          </button>
        )}
        <span className="wbe-spacer" />
        <button type="button" className="wbe-btn" onClick={close}>
          {t('dialog.cancel')}
        </button>
        <button type="submit" className="wbe-btn wbe-primary" disabled={!changed}>
          {t('comments.saveNote')}
        </button>
      </div>
    </form>
  );
}

/** A comment thread: its comments, with the user's own to edit or delete, a reply box, and resolving or deleting it. A new thread's first comment otherwise. */
function Thread({ store, at, note, locked, close, dialog }: Props & { dialog: RefObject<HTMLDivElement | null> }) {
  const t = useT();
  const id = useId();
  const [text, setText] = useState('');
  const [editing, setEditing] = useState<string>();
  const [draft, setDraft] = useState('');
  const { row, col } = at;
  const textBox = useRef<HTMLTextAreaElement>(null);
  const reopenButton = useRef<HTMLButtonElement>(null);
  const closeButton = useRef<HTMLButtonElement>(null);
  // The reply box; or, with none, reopening a resolved thread, or closing a read-only one.
  useLayoutEffect(() => (textBox.current ?? reopenButton.current ?? closeButton.current)?.focus({ preventScroll: true }), []);
  const hint = (
    <p id={`${id}-hint`} className="wbe-comment-meta">
      {t('comments.sendHint', { key: mod(t, 'Enter') })}
    </p>
  );
  const box = (label: string, rows: number) => (
    <label className="wbe-field">
      <span>{label}</span>
      <textarea
        ref={textBox}
        className="wbe-comment-input"
        value={text}
        rows={rows}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={sendOnModEnter}
        aria-describedby={`${id}-hint`}
      />
    </label>
  );

  if (!note) {
    const start = async (e: FormEvent) => {
      e.preventDefault();
      if (text.trim() && (await store.startThread(row, col, text))) close();
    };
    return (
      <form onSubmit={(e) => void start(e)}>
        {box(t('comments.commentLabel'), 3)}
        {hint}
        <div className="wbe-modal-actions">
          <button type="button" className="wbe-btn" onClick={close}>
            {t('dialog.cancel')}
          </button>
          <button type="submit" className="wbe-btn wbe-primary" disabled={!text.trim()}>
            {t('comments.post')}
          </button>
        </div>
      </form>
    );
  }

  const reply = async (e: FormEvent) => {
    e.preventDefault();
    if (text.trim() && (await store.reply(row, col, text))) setText('');
  };
  const saveEdit = async (e: FormEvent, entry: NoteEntry) => {
    e.preventDefault();
    if (draft === entry.text || (draft.trim() && (await store.editComment(row, col, entry.id!, draft)))) setEditing(undefined);
  };
  const removeReply = async (entry: NoteEntry) => {
    // Its buttons go with it: the focus stays in the popup.
    if (await store.deleteReply(row, col, entry.id!)) dialog.current?.focus();
  };
  const removeThread = async () => {
    if (await store.deleteComment(row, col)) close();
  };
  const me = store.commentAuthor;
  return (
    <>
      {note.resolved && <p className="wbe-comment-meta">{t('note.resolved')}</p>}
      <ol className="wbe-comment-entries">
        {note.entries.map((e, i) => {
          const own = !locked && !!e.id && e.author === me;
          const textId = `${id}-${i}`;
          return (
            <li key={e.id ?? i} className="wbe-comment-entry">
              <div className="wbe-comment-head">
                <span className="wbe-comment-who">{e.author ?? t('comments.someone')}</span>
                {e.date && <span className="wbe-comment-meta">{noteDate(t, e.date)}</span>}
              </div>
              {editing !== undefined && editing === e.id ? (
                <EditEntry text={draft} onChange={setDraft} onSave={(ev) => void saveEdit(ev, e)} onCancel={() => setEditing(undefined)} />
              ) : (
                <p id={textId} className="wbe-comment-text">
                  {e.text}
                </p>
              )}
              {own && editing !== e.id && (
                <div className="wbe-comment-actions">
                  <button
                    type="button"
                    className="wbe-link"
                    aria-describedby={textId}
                    onClick={() => {
                      setEditing(e.id);
                      setDraft(e.text);
                    }}
                  >
                    {t('comments.edit')}
                  </button>
                  {i > 0 && (
                    <button type="button" className="wbe-link" aria-describedby={textId} onClick={() => void removeReply(e)}>
                      {t('comments.deleteReply')}
                    </button>
                  )}
                </div>
              )}
            </li>
          );
        })}
      </ol>
      {!locked && !note.resolved && (
        <form className="wbe-comment-reply" onSubmit={(e) => void reply(e)}>
          {box(t('comments.replyLabel'), 2)}
          {hint}
          <div className="wbe-comment-actions">
            <button type="submit" className="wbe-btn wbe-primary" disabled={!text.trim()}>
              {t('comments.reply')}
            </button>
          </div>
        </form>
      )}
      <div className="wbe-modal-actions">
        {!locked && (
          <button type="button" className="wbe-btn wbe-comment-delete" onClick={() => void removeThread()}>
            {t('comments.deleteThread')}
          </button>
        )}
        <span className="wbe-spacer" />
        {!locked && (
          <button ref={note.resolved ? reopenButton : undefined} type="button" className="wbe-btn" onClick={() => void store.resolveThread(row, col, !note.resolved)}>
            {t(note.resolved ? 'comments.reopen' : 'comments.resolve')}
          </button>
        )}
        <button ref={closeButton} type="button" className="wbe-btn" onClick={close}>
          {t('comments.close')}
        </button>
      </div>
    </>
  );
}

/** One of the user's own comments, being changed in place. */
function EditEntry({ text, onChange, onSave, onCancel }: { text: string; onChange: (text: string) => void; onSave: (e: FormEvent) => void; onCancel: () => void }) {
  const t = useT();
  const input = useFocusFirst<HTMLTextAreaElement>();
  return (
    <form onSubmit={onSave}>
      <label className="wbe-field">
        <span className="wbe-sr-only">{t('comments.editLabel')}</span>
        <textarea ref={input} className="wbe-comment-input" value={text} rows={3} onChange={(e) => onChange(e.target.value)} onKeyDown={sendOnModEnter} />
      </label>
      <div className="wbe-comment-actions">
        <button type="button" className="wbe-btn" onClick={onCancel}>
          {t('dialog.cancel')}
        </button>
        <button type="submit" className="wbe-btn wbe-primary" disabled={!text.trim()}>
          {t('comments.save')}
        </button>
      </div>
    </form>
  );
}
