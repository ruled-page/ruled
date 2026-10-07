import type { CellNote } from '../core/types';

/** A cell's note or comment thread, shown beside the cell while the pointer is over it. */
export function NoteTip({ note, x, y }: { note: CellNote; x: number; y: number }) {
  // Keep the tip on screen when the cell is near the right or bottom edge.
  const left = Math.min(x, window.innerWidth - 296);
  const top = Math.min(y, window.innerHeight - 120);
  return (
    <div className={`note-tip ${note.kind}`} style={{ left, top }} role="tooltip">
      {note.resolved && <div className="resolved">Resolved</div>}
      {note.entries.map((e, i) => (
        <div className="entry" key={i}>
          {e.author && (
            <span className="who">
              {e.author}
              {e.date && <span className="when">{e.date}</span>}
            </span>
          )}
          {e.text}
        </div>
      ))}
    </div>
  );
}
