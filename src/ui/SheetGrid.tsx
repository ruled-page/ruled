import {
  CompactSelection,
  DataEditor,
  GridCellKind,
  type DataEditorRef,
  type DrawCellCallback,
  type FillPatternEventArgs,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type Item,
  type Rectangle,
  type Theme,
} from '@glideapps/glide-data-grid';
import '@glideapps/glide-data-grid/dist/index.css';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { colToLetters } from '../core/refs';
import { CellFlag, type DisplayStyle, type Rect } from '../core/types';
import { parseClipboardText } from '../state/clipboard';
import { alignment, displayText, editText, exactText } from '../state/format';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { EnumEditor, enumOptionsOf, type EnumCell } from './EnumEditor';
import { ObjectsOverlay } from './ObjectsOverlay';

const FONT_STACK = 'Calibri, Carlito, "Aptos Narrow", "Segoe UI", system-ui, sans-serif';
const EDITED_FILL = '#FFF4CE';
/** Background for pivot tables and array formulas, which cannot be edited here. */
const LOCKED_FILL = '#F2F4F7';
/** Background for labels fixed by a schema. */
const LABEL_FILL = '#EAF2EC';

const THEME: Partial<Theme> = {
  accentColor: '#1E7145',
  accentFg: '#FFFFFF',
  accentLight: 'rgba(30, 113, 69, 0.10)',
  textDark: '#1F1F1F',
  textMedium: '#5F6368',
  textLight: '#9AA0A6',
  textHeader: '#4A4A4A',
  textHeaderSelected: '#1E7145',
  bgCell: '#FFFFFF',
  bgCellMedium: '#F7F7F7',
  bgHeader: '#F4F5F5',
  bgHeaderHasFocus: '#E3EFE8',
  bgHeaderHovered: '#EAEBEB',
  borderColor: '#E1E3E3',
  horizontalBorderColor: '#E1E3E3',
  headerFontStyle: '500 12px',
  baseFontStyle: '14.7px',
  fontFamily: FONT_STACK,
  editorFontSize: '14.7px',
  cellHorizontalPadding: 4,
  cellVerticalPadding: 2,
  lineHeight: 1.25,
};

const ROW_HEADER_THEME: Partial<Theme> = {
  bgCell: '#F4F5F5',
  textDark: '#5F6368',
  baseFontStyle: '12px',
  fontFamily: 'system-ui, -apple-system, "Segoe UI", sans-serif',
};

const themeCache = new Map<string, Partial<Theme>>();
function cellTheme(style: DisplayStyle, color: string | undefined, edited: boolean, locked = false): Partial<Theme> | undefined {
  const fill = edited ? EDITED_FILL : (style.fill ?? (locked ? LOCKED_FILL : undefined));
  const textColor = color ?? style.color;
  const plain = !fill && !textColor && !style.bold && !style.italic && Math.abs(style.fontSizePx - 14.7) < 0.2 && /^Calibri$/i.test(style.fontFamily);
  if (plain) return undefined;
  const key = `${fill}|${textColor}|${style.bold}|${style.italic}|${style.fontSizePx}|${style.fontFamily}`;
  let t = themeCache.get(key);
  if (!t) {
    // Glide spreads overrides over the base theme, so unset keys must be absent, not undefined.
    t = {
      baseFontStyle: `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px`,
      fontFamily: `"${style.fontFamily}", ${FONT_STACK}`,
    };
    if (fill) t.bgCell = fill;
    if (textColor) t.textDark = textColor;
    themeCache.set(key, t);
  }
  return t;
}

interface Props {
  store: WorkbookStore;
  view: SheetView;
  /** Changes whenever cell data or styles change, so the grid redraws. */
  dataVersion: number;
  /** Changes whenever schema issues change, so markers redraw. */
  schemaVersion: number;
  selection: Selection;
}

export function SheetGrid({ store, view, dataVersion, schemaVersion, selection }: Props) {
  const ref = useRef<DataEditorRef>(null);
  const host = useRef<HTMLDivElement>(null);
  const [colVersion, setColVersion] = useState(0);
  const [region, setRegion] = useState(0);
  const styles = store.styles;
  const date1904 = store.date1904;
  const highlight = store.highlightEdits;
  const numbers = store.numbers;
  const schema = store.schema;
  const sheetIndex = view.info.index;
  const rowHeaderWidth = Math.max(40, String(view.visibleRows[view.rowCount - 1] + 1).length * 8 + 16);

  const frozen = useMemo(() => {
    let n = 0;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) n++;
    return n;
  }, [view]);

  const columns = useMemo<GridColumn[]>(
    () => [
      { title: '', id: '#', width: rowHeaderWidth },
      ...Array.from(view.visibleCols, (c) => ({ title: colToLetters(c), id: String(c), width: view.colWidth(c) })),
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, view.colCount, colVersion, rowHeaderWidth],
  );

  const getCellContent = useCallback(
    ([gc, gr]: Item): GridCell => {
      const row = view.visibleRows[gr];
      if (gc === 0) {
        return {
          kind: GridCellKind.Text,
          data: '',
          displayData: String(row + 1),
          allowOverlay: false,
          readonly: true,
          contentAlign: 'center',
          themeOverride: ROW_HEADER_THEME,
          cursor: 'default',
        };
      }
      let col = view.visibleCols[gc - 1];
      let span: [number, number] | undefined;
      let showContent = true;
      const merge = view.mergeAt(row, col);
      if (merge) {
        let start = view.gridCol(merge.c1) + 1;
        let end = view.gridCol(merge.c2) + 1;
        for (let c = merge.c2; end <= 0 && c >= merge.c1; c--) end = view.gridCol(c) + 1;
        // Glide clips a span's text to its frozen part, so a merge that
        // crosses the freeze line is split there and its text drawn in the
        // scrolling part.
        const firstScrolling = 1 + frozen;
        if (start > 0 && start < firstScrolling && end >= firstScrolling) {
          if (gc < firstScrolling) {
            end = firstScrolling - 1;
            showContent = false;
          } else {
            start = firstScrolling;
          }
        }
        if (start > 0 && end > start) span = [start, end];
        showContent = showContent && row === merge.r1;
        col = merge.c1;
      }
      const rec = showContent ? view.get(merge ? merge.r1 : row, col) : null;
      const style = styles[rec?.style ?? 0] ?? styles[0];
      const readOnly = view.readOnlyReason(row, view.visibleCols[gc - 1]);
      const blocked = view.blockedAt(row, view.visibleCols[gc - 1]);
      const locked = blocked !== undefined && blocked.kind !== 'tableHeader';
      const label = !!schema?.isLabel(sheetIndex, row, view.visibleCols[gc - 1]);
      // Enum fields edit with a dropdown; clicking the selected cell opens it.
      const field = schema && !readOnly ? schema.fieldAt(sheetIndex, row, view.visibleCols[gc - 1])?.field : undefined;
      let options: readonly string[] | undefined = field?.type === 'enum' && field.values?.length ? field.values : undefined;
      if (!options && !readOnly && view.info.listValidations.length) {
        // Excel's own list validations get the same dropdown.
        const dv = store.listAt(sheetIndex, row, view.visibleCols[gc - 1]);
        if (dv?.inCellDropdown) options = store.listOptions(dv, row, view.visibleCols[gc - 1]);
        if (!options?.length) options = undefined;
      }
      const asEnum = (cell: EnumCell): EnumCell =>
        options ? { ...cell, enumOptions: options, activationBehaviorOverride: 'second-click' } : cell;
      if (!rec) {
        const mergedStyle = merge ? styles[view.get(merge.r1, merge.c1)?.style ?? 0] : undefined;
        const bg = mergedStyle?.fill ?? (locked ? LOCKED_FILL : label ? LABEL_FILL : undefined);
        return asEnum({
          kind: GridCellKind.Text,
          data: '',
          displayData: '',
          allowOverlay: !readOnly,
          span,
          themeOverride: bg ? { bgCell: bg } : undefined,
        });
      }
      const { text, color } = displayText(rec, style, date1904, numbers.display === 'full');
      return asEnum({
        kind: GridCellKind.Text,
        data: editText(rec, style, date1904),
        displayData: text,
        copyData: numbers.copy === 'raw' ? (exactText(rec, style, false) ?? text) : text,
        allowOverlay: !readOnly,
        contentAlign: alignment(rec, style),
        allowWrapping: style.wrap,
        span,
        themeOverride: label && !style.fill ? { ...cellTheme(style, color, false, false), bgCell: LABEL_FILL } : cellTheme(style, color, highlight && !!(rec.flags & CellFlag.Edited), locked),
      });
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, highlight, frozen, numbers, schema, schemaVersion],
  );

  const drawCell = useCallback<DrawCellCallback>(
    (args, drawContent) => {
      drawContent();
      if (args.col === 0) return;
      const row = view.visibleRows[args.row];
      const col = view.visibleCols[args.col - 1];
      const { ctx, rect } = args;
      if (args.highlighted && enumOptionsOf(args.cell)) {
        // Dropdown chevron on selected enum cells, like Excel's list validation.
        const cx = rect.x + rect.width - 10;
        const cy = rect.y + rect.height / 2 + 1;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(cx - 3.5, cy - 2);
        ctx.lineTo(cx, cy + 1.5);
        ctx.lineTo(cx + 3.5, cy - 2);
        ctx.strokeStyle = '#5F6368';
        ctx.lineWidth = 1.5;
        ctx.stroke();
        ctx.restore();
      }
      const rec = view.get(row, col);
      if (rec && rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch | CellFlag.Stale)) {
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(rect.x + 1, rect.y + 1);
        ctx.lineTo(rect.x + 8, rect.y + 1);
        ctx.lineTo(rect.x + 1, rect.y + 8);
        ctx.closePath();
        ctx.fillStyle = rec.flags & CellFlag.Stale ? '#E8A33D' : '#8A94A6';
        ctx.fill();
        ctx.restore();
      }
      if (!schema) return;
      const issues = schema.cellIssues(sheetIndex, row, col);
      const right = rect.x + rect.width - 1;
      ctx.save();
      if (issues.length) {
        ctx.beginPath();
        ctx.moveTo(right, rect.y + 1);
        ctx.lineTo(right - 8, rect.y + 1);
        ctx.lineTo(right, rect.y + 9);
        ctx.closePath();
        ctx.fillStyle = issues.some((i) => i.severity === 'error') ? '#D92D20' : '#F79009';
        ctx.fill();
      } else if (schema.isPending(sheetIndex, row, col)) {
        ctx.beginPath();
        ctx.arc(right - 4, rect.y + 5, 2.5, 0, Math.PI * 2);
        ctx.fillStyle = '#2E90FA';
        ctx.fill();
      }
      ctx.restore();
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, schema, schemaVersion, sheetIndex],
  );

  // Grid selection <-> sheet selection. Grid column 0 is the row-number column.
  const toSheetRect = useCallback(
    (r: Rectangle): Rect => {
      const x1 = Math.max(1, r.x);
      const x2 = Math.max(x1, r.x + r.width - 1);
      return {
        r1: view.visibleRows[r.y],
        r2: view.visibleRows[Math.min(view.rowCount - 1, r.y + r.height - 1)],
        c1: view.visibleCols[x1 - 1],
        c2: view.visibleCols[Math.min(view.colCount, x2) - 1],
      };
    },
    [view],
  );

  const gridSelection = useMemo<GridSelection>(() => {
    const gr = Math.max(0, view.gridRow(selection.row));
    const gc = Math.max(0, view.gridCol(selection.col)) + 1;
    const [first, ...rest] = selection.rects.map((r) => {
      const y1 = Math.max(0, view.gridRow(r.r1));
      const y2 = Math.max(y1, view.gridRow(r.r2));
      const x1 = Math.max(0, view.gridCol(r.c1)) + 1;
      const x2 = Math.max(x1, view.gridCol(r.c2) + 1);
      return { x: x1, y: y1, width: x2 - x1 + 1, height: y2 - y1 + 1 };
    });
    return {
      current: { cell: [gc, gr], range: first ?? { x: gc, y: gr, width: 1, height: 1 }, rangeStack: rest },
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty(),
    };
  }, [view, selection]);

  const onSelectionChange = useCallback(
    (sel: GridSelection) => {
      if (sel.columns.length > 0) {
        const cols = sel.columns.toArray();
        const c1 = view.visibleCols[Math.max(1, Math.min(...cols)) - 1];
        const c2 = view.visibleCols[Math.max(1, Math.max(...cols)) - 1];
        const r2 = Math.max(view.info.maxRow, 0);
        store.setSelection({ row: 0, col: c1, rects: [{ r1: 0, c1, r2, c2 }] });
        return;
      }
      if (!sel.current) return;
      const [gc, gr] = sel.current.cell;
      if (gc === 0) {
        // Clicking a row number selects the whole row.
        const rows = [sel.current.range.y, sel.current.range.y + sel.current.range.height - 1].map((y) => view.visibleRows[y]);
        store.setSelection({
          row: rows[0],
          col: view.visibleCols[0],
          rects: [{ r1: rows[0], r2: rows[1], c1: 0, c2: Math.max(view.info.maxCol, 0) }],
        });
        return;
      }
      store.setSelection({
        row: view.visibleRows[gr],
        col: view.visibleCols[gc - 1],
        rects: [sel.current.range, ...sel.current.rangeStack].map(toSheetRect),
      });
    },
    [store, view, toSheetRect],
  );

  // Remember what was copied, so pasting it back carries formulas.
  useEffect(() => {
    const onCopy = () => {
      const active = document.activeElement;
      if (active && /^(INPUT|TEXTAREA)$/.test(active.tagName) && !active.closest('.gdg-clip-region, .dvn-scroller')) return;
      const first = store.selection.rects[0];
      if (first) store.noteCopy(first);
    };
    window.addEventListener('copy', onCopy, true);
    window.addEventListener('cut', onCopy, true);
    return () => {
      window.removeEventListener('copy', onCopy, true);
      window.removeEventListener('cut', onCopy, true);
    };
  }, [store]);

  // Glide reads pastes through navigator.clipboard.read(), which makes
  // browsers ask for clipboard permission. The paste event already carries
  // the data, so handle it here instead.
  useEffect(() => {
    const onPaste = (e: ClipboardEvent) => {
      if (!host.current?.contains(document.activeElement)) return;
      const text = e.clipboardData?.getData('text/plain');
      if (!text) return;
      e.preventDefault();
      const values = parseClipboardText(text);
      const sel = store.selection;
      const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
      void store.paste({ row: rect.r1, col: rect.c1 }, values, rect);
    };
    window.addEventListener('paste', onPaste, true);
    return () => window.removeEventListener('paste', onPaste, true);
  }, [store]);

  const onDelete = useCallback(
    (sel: GridSelection) => {
      const rects: Rect[] = [];
      if (sel.columns.length > 0) {
        const cols = sel.columns.toArray().filter((c) => c > 0);
        if (cols.length) {
          rects.push({ r1: 0, r2: Math.max(view.info.maxRow, 0), c1: view.visibleCols[Math.min(...cols) - 1], c2: view.visibleCols[Math.max(...cols) - 1] });
        }
      } else if (sel.current) {
        rects.push(...[sel.current.range, ...sel.current.rangeStack].map(toSheetRect));
      }
      if (rects.length) void store.clear(rects);
      return false;
    },
    [store, view, toSheetRect],
  );

  const onFillPattern = useCallback(
    (e: FillPatternEventArgs) => {
      e.preventDefault();
      void store.fill(toSheetRect(e.patternSource), toSheetRect(e.fillDestination));
    },
    [store, toSheetRect],
  );

  const onCellEdited = useCallback(
    ([gc, gr]: Item, value: { data?: unknown }) => {
      if (gc === 0) return;
      void store.input(view.visibleRows[gr], view.visibleCols[gc - 1], String(value.data ?? ''));
    },
    [store, view],
  );

  const onCellActivated = useCallback(
    ([gc, gr]: Item) => {
      if (gc === 0) return;
      const reason = view.readOnlyReason(view.visibleRows[gr], view.visibleCols[gc - 1]);
      if (reason) store.toast('warn', reason);
    },
    [store, view],
  );

  const onColumnResize = useCallback(
    (_c: GridColumn, size: number, index: number) => {
      if (index === 0) return;
      view.colWidths.set(view.visibleCols[index - 1], size);
      setColVersion((v) => v + 1);
    },
    [view],
  );

  const rowHeight = useCallback((gr: number) => view.rowHeight(view.visibleRows[gr]), [view]);
  const provideEditor = useCallback(
    (cell: GridCell) => (enumOptionsOf(cell) ? { editor: EnumEditor, disablePadding: true } : undefined),
    [],
  );
  const theme = useMemo(
    () => (view.info.showGridLines ? THEME : { ...THEME, borderColor: 'transparent', horizontalBorderColor: 'transparent' }),
    [view],
  );

  return (
    <div className="grid-host" ref={host}>
      <DataEditor
        ref={ref}
        key={view.info.index}
        columns={columns}
        rows={view.rowCount}
        getCellContent={getCellContent}
        drawCell={drawCell}
        gridSelection={gridSelection}
        onGridSelectionChange={onSelectionChange}
        onCellEdited={onCellEdited}
        onCellActivated={onCellActivated}
        keybindings={{ paste: false }}
        provideEditor={provideEditor}
        onDelete={onDelete}
        fillHandle
        allowedFillDirections="orthogonal"
        onFillPattern={onFillPattern}
        onColumnResize={onColumnResize}
        getCellsForSelection
        rangeSelect="rect"
        columnSelect="multi"
        rowSelect="none"
        rowMarkers="none"
        freezeColumns={1 + frozen}
        rowHeight={rowHeight}
        headerHeight={24}
        smoothScrollX
        smoothScrollY
        cellActivationBehavior="double-click"
        editOnType
        theme={theme}
        width="100%"
        height="100%"
        minColumnWidth={18}
        maxColumnWidth={1200}
        onVisibleRegionChanged={() => setRegion((n) => n + 1)}
      />
      <ObjectsOverlay view={view} gridRef={ref} tick={region + colVersion} />
    </div>
  );
}
