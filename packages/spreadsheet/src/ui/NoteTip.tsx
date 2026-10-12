import type { CellNote } from '../core/types';
import type { Translator } from '../i18n';
import { useT } from '../i18n/react';

/** A comment's date (an ISO date such as 2024-05-01) in the editor's language, or as written if it isn't one. */
export function noteDate(t: Translator, date: string): string {
  const at = Date.parse(date);
  return Number.isNaN(at) ? date : t.date(at, { dateStyle: 'medium', timeZone: 'UTC' });
}

/** A cell's note or comment thread, shown beside the cell while the pointer is over it. */
export function NoteTip({ note, x, y }: { note: CellNote; x: number; y: number }) {
  const t = useT();
  // Keep the tip on screen when the cell is near the right or bottom edge.
  const left = Math.min(x, window.innerWidth - 296);
  const top = Math.min(y, window.innerHeight - 120);
  return (
    <div className={`wbe-note-tip wbe-${note.kind}`} style={{ left, top }} role="tooltip">
      {note.resolved && <div className="wbe-resolved">{t('note.resolved')}</div>}
      {note.entries.map((e, i) => (
        <div className="wbe-entry" key={i}>
          {e.author && (
            <span className="wbe-who">
              {e.author}
              {e.date && <span className="wbe-when">{noteDate(t, e.date)}</span>}
            </span>
          )}
          {e.text}
        </div>
      ))}
    </div>
  );
}
