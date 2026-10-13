import { useId, useRef, useState, type DragEvent } from 'react';
import { NEEDS_BASE_FIELD, NEEDS_BASE_ITEM, SHOW_AS, type PivotFn, type PivotSetup, type PivotShowAs } from '../core/pivotTable';
import type { PivotInfo } from '../core/types';
import { useT } from '../i18n/react';
import { addField, areaFields, fieldsInUse, moveChip, moveTo, PIVOT_FNS, removeChip, removeField, setShowAs, setShown, setSummary, shownItems, VALUES_FIELD, type DataName, type PivotArea } from '../state/pivotFields';
import type { WorkbookStore } from '../state/store';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { Icon } from './icons';
import { Select } from './Select';

/** Excel's order: Filters and Columns above, Rows and Values below. */
const AREAS: PivotArea[] = ['pages', 'cols', 'rows', 'data'];

/** What is being dragged: a field from the list, or one already in an area. */
type Dragged = { field: number } | { area: PivotArea; index: number };

/**
 * Excel's PivotTable Fields pane for the pivot table the active cell is in:
 * its fields, ticked when used, and the four areas they go in. Fields are
 * dragged between the areas, or moved with each one's menu, which also
 * changes how a value is summarised; a field's items are filtered from the
 * list. Each change lays the pivot table out again as one undoable step.
 */
export function PivotFieldsBody({ store }: { store: WorkbookStore }) {
  const t = useT();
  const pivot = store.activePivot;
  if (!pivot) return <p className="wbe-report-body wbe-muted">{t('pivot.pane.none')}</p>;
  if (pivot.kept) {
    return (
      <div className="wbe-report-body">
        <p className="wbe-report-lede">{pivot.name}</p>
        <p className="wbe-muted">{t.text(pivot.kept)}</p>
      </div>
    );
  }
  return <Fields key={pivot.path} store={store} pivot={pivot} />;
}

function Fields({ store, pivot }: { store: WorkbookStore; pivot: PivotInfo }) {
  const t = useT();
  const id = useId();
  const { setup, fields } = pivot;
  const [menu, setMenu] = useState<{ x: number; y: number; area: PivotArea; index: number } | null>(null);
  const [filtering, setFiltering] = useState<number | null>(null);
  const [showing, setShowing] = useState<{ index: number; showAs: PivotShowAs } | null>(null);
  const [over, setOver] = useState<PivotArea | null>(null);
  const dragged = useRef<Dragged | null>(null);
  const used = fieldsInUse(setup);
  const name: DataName = (fn, field) => t(`pivot.dataName.${fn}`, { field: fields[field]?.name ?? '' });
  const fnFor = (field: number): PivotFn => (fields[field]?.numeric ? 'sum' : 'count');
  const change = (next: PivotSetup) => void store.setPivotFields(pivot.path, next);
  const label = (field: number) => (field === VALUES_FIELD ? t('pivot.pane.values') : (fields[field]?.name ?? ''));
  const onAxis = (field: number) => setup.rows.includes(field) || setup.cols.includes(field) || setup.pages.some((p) => p.field === field);

  const drop = (e: DragEvent, area: PivotArea, at: number) => {
    e.preventDefault();
    e.stopPropagation();
    setOver(null);
    const from = dragged.current;
    dragged.current = null;
    if (!from) return;
    if ('field' in from) change(moveTo(setup, from.field, area, at, name, fnFor(from.field)));
    else change(moveChip(setup, from, { area, at }, name, fnFor(areaFields(setup, from.area)[from.index])));
  };
  const dragOver = (e: DragEvent, area: PivotArea) => {
    if (!dragged.current) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setOver(area);
  };
  const startDrag = (e: DragEvent, what: Dragged) => {
    dragged.current = what;
    e.dataTransfer.effectAllowed = 'move';
    // Firefox drags only with data set.
    e.dataTransfer.setData('text/plain', 'field' in what ? label(what.field) : label(areaFields(setup, what.area)[what.index]));
  };

  const menuItems = (area: PivotArea, index: number): MenuItem[] => {
    const list = areaFields(setup, area);
    const field = list[index];
    const move = (to: PivotArea, at: number) => change(moveChip(setup, { area, index }, { area: to, at }, name, fnFor(field)));
    const items: MenuItem[] = [
      { label: t('pivot.menu.moveUp'), disabled: index === 0, onSelect: () => move(area, index - 1) },
      { label: t('pivot.menu.moveDown'), disabled: index === list.length - 1, onSelect: () => move(area, index + 2) },
      { label: t('pivot.menu.moveFirst'), disabled: index === 0, onSelect: () => move(area, 0) },
      { label: t('pivot.menu.moveLast'), disabled: index === list.length - 1, onSelect: () => move(area, list.length) },
      { separator: true },
      ...AREAS.filter((to) => to !== area).map((to) => ({
        label: t(`pivot.menu.moveTo.${to}`),
        disabled: field === VALUES_FIELD && (to === 'pages' || to === 'data'),
        onSelect: () => move(to, areaFields(setup, to).length),
      })),
      { separator: true },
    ];
    if (area === 'data') {
      const d = setup.data[index];
      items.push({
        label: t('pivot.menu.summarize'),
        items: PIVOT_FNS.map((fn) => ({ label: t(`pivot.fn.${fn}`), checked: d.fn === fn, onSelect: () => change(setSummary(setup, index, fn, name)) })),
      });
      items.push({
        label: t('pivot.menu.showAs'),
        items: SHOW_AS.map((kind) => ({
          label: t(`pivot.showAs.${kind}`),
          checked: (d.showAs ?? 'normal') === kind,
          // Kinds measured against another field ask which, and which item.
          onSelect: () => (NEEDS_BASE_FIELD.has(kind) ? setShowing({ index, showAs: kind }) : change(setShowAs(setup, index, kind))),
        })),
      });
    } else if (field !== VALUES_FIELD) {
      items.push({ label: t('pivot.menu.filter'), onSelect: () => setFiltering(field) });
    }
    if (field !== VALUES_FIELD) items.push({ label: t('pivot.menu.remove'), onSelect: () => change(removeChip(setup, area, index)) });
    return items;
  };

  return (
    <div className="wbe-report-body wbe-pivot-pane">
      <p className="wbe-report-lede">{pivot.name}</p>
      <h3 id={`${id}-fields`}>{t('pivot.pane.choose')}</h3>
      <ul className="wbe-pivot-fields" aria-labelledby={`${id}-fields`}>
        {fields.map((f, i) => (
          <li key={i} draggable onDragStart={(e) => startDrag(e, { field: i })} onDragEnd={() => (dragged.current = null)}>
            <label>
              <input type="checkbox" checked={used.has(i)} onChange={(e) => change(e.target.checked ? addField(setup, i, f.numeric, name) : removeField(setup, i))} />
              <span>{f.name}</span>
            </label>
            {onAxis(i) && (
              <button
                type="button"
                className={`wbe-btn wbe-icon wbe-ghost wbe-small${setup.hidden[i]?.length || setup.pages.some((p) => p.field === i && p.item !== undefined) ? ' wbe-on' : ''}`}
                aria-expanded={filtering === i}
                title={t('pivot.filter.button', { field: f.name })}
                aria-label={t('pivot.filter.button', { field: f.name })}
                onClick={() => setFiltering(filtering === i ? null : i)}
              >
                <Icon name="chevronDown" />
              </button>
            )}
            {filtering === i && <ItemFilter pivot={pivot} field={i} onDone={(next) => {
                  setFiltering(null);
                  if (next) change(next);
                }} />}
          </li>
        ))}
      </ul>
      <h3>{t('pivot.pane.drag')}</h3>
      <div className="wbe-pivot-areas">
        {AREAS.map((area) => {
          const list = areaFields(setup, area);
          return (
            <section
              key={area}
              className={`wbe-pivot-area${over === area ? ' wbe-drop' : ''}`}
              aria-labelledby={`${id}-${area}`}
              onDragOver={(e) => dragOver(e, area)}
              onDragLeave={() => setOver(null)}
              onDrop={(e) => drop(e, area, list.length)}
            >
              <h4 id={`${id}-${area}`}>{t(`pivot.area.${area}`)}</h4>
              <ul>
                {list.map((field, index) => {
                  const text = area === 'data' ? setup.data[index].name : label(field);
                  return (
                    <li key={`${field}-${index}`} draggable onDragStart={(e) => startDrag(e, { area, index })} onDragEnd={() => (dragged.current = null)} onDrop={(e) => drop(e, area, index)}>
                      <button
                        type="button"
                        className="wbe-pivot-chip"
                        aria-haspopup="menu"
                        title={t('pivot.chip.menu', { field: text })}
                        onClick={(e) => {
                          const r = e.currentTarget.getBoundingClientRect();
                          setMenu({ x: r.left, y: r.bottom, area, index });
                        }}
                      >
                        <span>{text}</span>
                        <Icon name="chevronDown" size={14} />
                      </button>
                    </li>
                  );
                })}
              </ul>
              {area === 'data' && showing && (
                <ShowAsForm
                  pivot={pivot}
                  index={showing.index}
                  showAs={showing.showAs}
                  onDone={(next) => {
                    setShowing(null);
                    if (next) change(next);
                  }}
                />
              )}
            </section>
          );
        })}
      </div>
      {menu && <ContextMenu x={menu.x} y={menu.y} label={t('pivot.chip.menu', { field: label(areaFields(setup, menu.area)[menu.index]) })} items={menuItems(menu.area, menu.index)} onClose={() => setMenu(null)} />}
    </div>
  );
}

/** A field's items, ticked where shown, as Excel's filter for a field lists them. */
function ItemFilter({ pivot, field, onDone }: { pivot: PivotInfo; field: number; onDone: (next?: PivotSetup) => void }) {
  const t = useT();
  const id = useId();
  const f = pivot.fields[field];
  const all = f.items.map((it) => it.key);
  const [shown, setShownKeys] = useState(() => new Set(shownItems(pivot.setup, field, all)));
  const every = shown.size === all.length;
  const toggle = (key: string, on: boolean) => {
    const next = new Set(shown);
    if (on) next.add(key);
    else next.delete(key);
    setShownKeys(next);
  };
  return (
    <form
      className="wbe-pivot-filter"
      aria-labelledby={`${id}-title`}
      onSubmit={(e) => {
        e.preventDefault();
        onDone(setShown(pivot.setup, field, [...shown], all));
      }}
      onKeyDown={(e) => {
        if (e.key !== 'Escape') return;
        e.stopPropagation();
        onDone();
      }}
    >
      <span id={`${id}-title`} className="wbe-small wbe-muted">
        {t('pivot.filter.title', { field: f.name })}
      </span>
      <label>
        <input type="checkbox" checked={every} ref={(el) => {
            if (el) el.indeterminate = !every && shown.size > 0;
          }} onChange={(e) => setShownKeys(new Set(e.target.checked ? all : []))} />
        <span>{t('pivot.filter.all')}</span>
      </label>
      <div className="wbe-pivot-items">
        {f.items.map((it) => (
          <label key={it.key}>
            <input type="checkbox" checked={shown.has(it.key)} onChange={(e) => toggle(it.key, e.target.checked)} />
            <span>{it.label}</span>
          </label>
        ))}
      </div>
      {f.more && <p className="wbe-small wbe-muted">{t('pivot.filter.more', { count: f.items.length })}</p>}
      <div className="wbe-modal-actions">
        <button type="button" className="wbe-btn" onClick={() => onDone()}>
          {t('dialog.cancel')}
        </button>
        <button type="submit" className="wbe-btn wbe-primary" disabled={!shown.size}>
          {t('dialog.ok')}
        </button>
      </div>
    </form>
  );
}

/** Show Values As against a base field, and an item of it where the kind needs one, as Excel's dialog for them asks. */
function ShowAsForm({ pivot, index, showAs, onDone }: { pivot: PivotInfo; index: number; showAs: PivotShowAs; onDone: (next?: PivotSetup) => void }) {
  const t = useT();
  const id = useId();
  const d = pivot.setup.data[index];
  const placed = [...pivot.setup.rows, ...pivot.setup.cols].filter((f) => f >= 0);
  const [field, setField] = useState(d.baseField ?? placed[0] ?? 0);
  const [item, setItem] = useState(d.baseItem ?? 'previous');
  const items = pivot.fields[field]?.items ?? [];
  return (
    <form
      className="wbe-pivot-filter"
      aria-labelledby={`${id}-title`}
      onSubmit={(e) => {
        e.preventDefault();
        onDone(setShowAs(pivot.setup, index, showAs, field, NEEDS_BASE_ITEM.has(showAs) ? item : 'previous'));
      }}
      onKeyDown={(e) => {
        if (e.key !== 'Escape') return;
        e.stopPropagation();
        onDone();
      }}
    >
      <span id={`${id}-title`} className="wbe-small wbe-muted">
        {t('pivot.showAs.title', { name: d.name, kind: t(`pivot.showAs.${showAs}`) })}
      </span>
      <div className="wbe-field">
        <label htmlFor={`${id}-field`}>{t('pivot.showAs.baseField')}</label>
        <Select
          id={`${id}-field`}
          value={field}
          onChange={(e) => {
            const next = Number(e.target.value);
            setField(next);
            if (!pivot.fields[next]?.items.some((it) => it.key === item)) setItem('previous');
          }}
        >
          {pivot.fields.map((f, i) => (
            <option key={i} value={i}>
              {f.name}
            </option>
          ))}
        </Select>
      </div>
      {NEEDS_BASE_ITEM.has(showAs) && (
        <div className="wbe-field">
          <label htmlFor={`${id}-item`}>{t('pivot.showAs.baseItem')}</label>
          <Select id={`${id}-item`} value={item} onChange={(e) => setItem(e.target.value)}>
            <option value="previous">{t('pivot.showAs.previous')}</option>
            <option value="next">{t('pivot.showAs.next')}</option>
            {items.map((it) => (
              <option key={it.key} value={it.key}>
                {it.label}
              </option>
            ))}
          </Select>
        </div>
      )}
      <div className="wbe-modal-actions">
        <button type="button" className="wbe-btn" onClick={() => onDone()}>
          {t('dialog.cancel')}
        </button>
        <button type="submit" className="wbe-btn wbe-primary">
          {t('dialog.ok')}
        </button>
      </div>
    </form>
  );
}
