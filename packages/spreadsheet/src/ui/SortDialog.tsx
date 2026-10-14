import { useId, useRef, useState } from 'react';
import type { SortKey } from '../core/types';
import { useT } from '../i18n/react';
import { hasColors, type ColorPick } from '../state/filterConditions';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { Select } from './Select';
import { useDialog } from './useDialog';

/** One level of the sort: the column (or row) it orders by, and how. */
interface Level {
  index: number;
  /**
   * "asc" or "desc"; a custom list by its place among the store's lists,
   * forwards ("list:2") or backwards ("rlist:2"); or a colour or icon by its
   * place among the line's, on top ("color:1:top") or at the bottom.
   */
  order: string;
}

/** Excel allows 64 levels. */
const MAX_LEVELS = 64;

const ORDER_KEYS = {
  text: ['sort.ascText', 'sort.descText'],
  number: ['sort.ascNumber', 'sort.descNumber'],
  date: ['sort.ascDate', 'sort.descDate'],
} as const;

/** A custom list as the order picker shows it: its first entries. */
const preview = (list: readonly string[]) => (list.length > 3 ? `${list.slice(0, 3).join(', ')}, …` : list.join(', '));

/**
 * Excel's Sort dialog (Data › Sort, or Custom sort… in the cell menu) for the
 * block around the selection: levels, each a column to sort by and its order
 * (A to Z, Smallest to Largest, Oldest to Newest, the reverse, a custom
 * list such as the months, or a cell colour, font colour or icon on top or
 * at the bottom), whether the block has headings, a case-sensitive sort, and
 * sorting left to right.
 */
export function SortDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const setup = store.sortDialog;
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLSelectElement>(null);
  const id = useId();
  const [header, setHeader] = useState(false);
  const [across, setAcross] = useState(false);
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [levels, setLevels] = useState<Level[]>([]);
  const [tried, setTried] = useState(false);
  // The level to start with: the active cell's column (or row), if it's in the block.
  const start = (sideways: boolean): Level[] => {
    if (!setup) return [];
    const { rect } = setup;
    const { row, col } = store.selection;
    return [{ index: sideways ? Math.min(Math.max(row, rect.r1), rect.r2) : Math.min(Math.max(col, rect.c1), rect.c2), order: 'asc' }];
  };
  const opened = useRef<typeof setup>(undefined);
  if (opened.current !== setup) {
    opened.current = setup;
    if (setup) {
      setHeader(setup.header);
      setAcross(false);
      setCaseSensitive(false);
      setLevels(start(false));
      setTried(false);
    }
  }
  const close = () => store.set('sortDialog', undefined);
  useDialog(dialog, !!setup, close, { modal: true, initial });
  if (!setup) return null;

  const lines = store.sortLines(setup.rect, header, across);
  const lists = store.fillLists;
  const lineOf = (index: number) => lines.find((l) => l.index === index) ?? lines[0];
  // A line can be sorted by several colours, one level each, but by its values once.
  const seen = new Set<number>();
  const duplicate = levels.find((l) => !l.order.startsWith('color:') && (seen.has(l.index) ? true : (seen.add(l.index), false)));
  /** The colours and icons a line has, to sort by, when it has more than one of a kind. */
  const colorPicks = (index: number): ColorPick[] => {
    const c = store.sortColors(setup.rect, index, header, across);
    if (!hasColors(c)) return [];
    return [
      ...(c.fill.length > 1 ? c.fill.map((color) => ({ kind: 'fill' as const, ...(color && { color }) })) : []),
      ...(c.font.length > 1 ? c.font.map((color) => ({ kind: 'font' as const, ...(color && { color }) })) : []),
      ...c.icons.map((icon) => ({ kind: 'icon' as const, ...icon })),
    ];
  };
  const named = (on: ColorPick) =>
    on.kind === 'icon'
      ? on.index === undefined
        ? t('sort.what.noIcon')
        : t('sort.what.icon', { n: on.index + 1 })
      : on.color
        ? t(on.kind === 'fill' ? 'sort.what.fill' : 'sort.what.font', { color: on.color })
        : t(on.kind === 'fill' ? 'sort.what.noFill' : 'sort.what.automatic');
  const error = duplicate && t('sort.duplicate', { line: t.text(lineOf(duplicate.index).label) });

  const update = (i: number, patch: Partial<Level>) => setLevels(levels.map((l, j) => (j === i ? { ...l, ...patch } : l)));
  const move = (i: number, by: number) => {
    const next = [...levels];
    [next[i], next[i + by]] = [next[i + by], next[i]];
    setLevels(next);
  };
  const unused = lines.find((l) => !levels.some((v) => v.index === l.index));
  const keyOf = (l: Level): SortKey => {
    const [kind, n, end] = l.order.split(':');
    const pick = kind === 'color' ? colorPicks(l.index)[Number(n)] : undefined;
    if (pick) return store.colorSortKey(setup.rect, l.index, header, across, pick, end === 'bottom');
    const list = n === undefined ? undefined : lists[Number(n)];
    const descending = kind === 'desc' || kind === 'rlist';
    const order = { descending, ...(list && { list }) };
    return across ? { row: l.index, ...order } : { col: l.index, ...order };
  };
  const submit = () => {
    if (error) return setTried(true);
    close();
    void store.sortBy({ rect: setup.rect, keys: levels.map(keyOf), header, caseSensitive, across });
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-modal-wide wbe-sort-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        noValidate
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          submit();
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('sort.title')}</h2>
          <div className="wbe-sort-levels">
            <div className="wbe-sort-level wbe-sort-heading" aria-hidden="true">
              <span />
              <span>{t(across ? 'sort.rowHeading' : 'sort.columnHeading')}</span>
              <span>{t('sort.order')}</span>
            </div>
            {levels.map((level, i) => {
              const name = t(i === 0 ? 'sort.sortBy' : 'sort.thenBy');
              const kind = lineOf(level.index).kind;
              return (
                <div className="wbe-sort-level" key={i} role="group" aria-label={`${name} (${i + 1})`}>
                  <span className="wbe-sort-name">{name}</span>
                  <Select
                    ref={i === 0 ? initial : undefined}
                    aria-label={`${name}: ${t(across ? 'sort.rowHeading' : 'sort.columnHeading')}`}
                    value={level.index}
                    onChange={(e) => update(i, { index: Number(e.target.value), ...(level.order.startsWith('color:') && { order: 'asc' }) })}
                    aria-invalid={tried && level === duplicate}
                  >
                    {lines.map((l) => (
                      <option key={l.index} value={l.index}>
                        {t.text(l.label)}
                      </option>
                    ))}
                  </Select>
                  <Select aria-label={`${name}: ${t('sort.order')}`} value={level.order} onChange={(e) => update(i, { order: e.target.value })}>
                    <option value="asc">{t(ORDER_KEYS[kind][0])}</option>
                    <option value="desc">{t(ORDER_KEYS[kind][1])}</option>
                    {lists.map((list, n) => [
                      <option key={`list:${n}`} value={`list:${n}`}>
                        {t('sort.list', { list: preview(list) })}
                      </option>,
                      <option key={`rlist:${n}`} value={`rlist:${n}`}>
                        {t('sort.list', { list: preview([...list].reverse()) })}
                      </option>,
                    ])}
                    {colorPicks(level.index).flatMap((on, n) =>
                      (['top', 'bottom'] as const).map((end) => (
                        <option key={`color:${n}:${end}`} value={`color:${n}:${end}`}>
                          {t(end === 'top' ? 'sort.colorOnTop' : 'sort.colorOnBottom', { what: named(on) })}
                        </option>
                      )),
                    )}
                  </Select>
                  <span className="wbe-sort-tools">
                    <button type="button" className="wbe-btn wbe-icon wbe-ghost" disabled={i === 0} onClick={() => move(i, -1)} aria-label={t('sort.moveUp', { level: i + 1 })}>
                      <Icon name="chevronUp" />
                    </button>
                    <button type="button" className="wbe-btn wbe-icon wbe-ghost" disabled={i === levels.length - 1} onClick={() => move(i, 1)} aria-label={t('sort.moveDown', { level: i + 1 })}>
                      <Icon name="chevronDown" />
                    </button>
                    <button
                      type="button"
                      className="wbe-btn wbe-icon wbe-ghost"
                      disabled={levels.length === 1}
                      onClick={() => setLevels(levels.filter((_, j) => j !== i))}
                      aria-label={t('sort.deleteLevel', { level: i + 1 })}
                    >
                      <Icon name="trash" />
                    </button>
                  </span>
                </div>
              );
            })}
          </div>
          <button
            type="button"
            className="wbe-btn wbe-sort-add"
            disabled={!unused || levels.length >= MAX_LEVELS}
            onClick={() => unused && setLevels([...levels, { index: unused.index, order: 'asc' }])}
          >
            <Icon name="plus" />
            {t('sort.addLevel')}
          </button>
          {tried && error && (
            <p className="wbe-field-error" role="alert">
              {error}
            </p>
          )}
          <fieldset className="wbe-series-group wbe-sort-options">
            <legend>{t('sort.options')}</legend>
            <label>
              <input type="checkbox" checked={header && !across} disabled={across || setup.rect.r2 <= setup.rect.r1} onChange={(e) => setHeader(e.target.checked)} />
              {t('sort.headers')}
            </label>
            <label>
              <input type="checkbox" checked={caseSensitive} onChange={(e) => setCaseSensitive(e.target.checked)} />
              {t('sort.caseSensitive')}
            </label>
            {[false, true].map((sideways) => (
              <label key={String(sideways)}>
                <input
                  type="radio"
                  name={`${id}-orientation`}
                  checked={across === sideways}
                  onChange={() => {
                    setAcross(sideways);
                    setLevels(start(sideways));
                    setTried(false);
                  }}
                />
                {t(sideways ? 'sort.leftToRight' : 'sort.topToBottom')}
              </label>
            ))}
          </fieldset>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('dialog.ok')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
