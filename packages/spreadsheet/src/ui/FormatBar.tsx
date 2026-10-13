import { useEffect, useId, useRef, useState } from 'react';
import type { BorderEdge, BorderPlacement } from '../core/types';
import { useT } from '../i18n/react';
import { NUMBER_FORMAT_KINDS, numberFormatCode, numberFormatKind, type NumberFormatKind } from '../state/numberFormats';
import type { WorkbookStore } from '../state/store';
import { ContextMenu } from './ContextMenu';
import { Icon } from './icons';
import { useDialog } from './useDialog';
import { Select } from './Select';

const FONTS = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Excel's standard colours, then a row of its default theme's. */
const COLORS = {
  darkRed: '#C00000',
  red: '#FF0000',
  orange: '#FFC000',
  yellow: '#FFFF00',
  lightGreen: '#92D050',
  green: '#00B050',
  lightBlue: '#00B0F0',
  blue: '#0070C0',
  darkBlue: '#002060',
  purple: '#7030A0',
  white: '#FFFFFF',
  black: '#000000',
  lightGray: '#E7E6E6',
  blueGray: '#44546A',
  accentBlue: '#4472C4',
  accentOrange: '#ED7D31',
  gray: '#A5A5A5',
  gold: '#FFC000',
  skyBlue: '#5B9BD5',
  accentGreen: '#70AD47',
} as const;
const BORDERS: BorderPlacement[] = ['all', 'outside', 'inside', 'top', 'bottom', 'left', 'right', 'none'];

/** Fonts, colours, borders, alignment and number formats for the selected cells, as Excel's Home tab has them. */
export function FormatBar({ store }: { store: WorkbookStore }) {
  const t = useT();
  const view = store.view;
  // Not while an edit is in flight: changes queue behind it, and a button that disabled itself would drop the focus.
  const disabled = store.status !== 'ready' || view?.info.kind !== 'worksheet';
  const s = store.activeStyle;
  const sizePt = Math.round(((s.fontSizePx * 3) / 4) * 2) / 2;
  const kind = numberFormatKind(s.numFmt, store.locale);
  const [borders, setBorders] = useState<{ x: number; y: number } | null>(null);
  const pen = store.borderPen;
  const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike', icon: Parameters<typeof Icon>[0]['name']) => (
    <button className="wbe-btn wbe-icon" aria-pressed={s[key]} disabled={disabled} onClick={() => void store.toggleFormat(key)} title={t(`format.${key}`)} aria-label={t(`format.${key}`)}>
      <Icon name={icon} />
    </button>
  );
  const align = (value: 'left' | 'center' | 'right', icon: Parameters<typeof Icon>[0]['name']) => (
    <button
      className="wbe-btn wbe-icon"
      aria-pressed={s.hAlign === value}
      disabled={disabled}
      onClick={() => void store.formatCells({ hAlign: s.hAlign === value ? null : value })}
      title={t(`format.align.${value}`)}
      aria-label={t(`format.align.${value}`)}
    >
      <Icon name={icon} />
    </button>
  );
  // Excel's vertical alignment: bottom unless set, and pressing the one that's on goes back to bottom.
  const vAlign = s.vAlign === 'top' || s.vAlign === 'center' ? s.vAlign : 'bottom';
  const valign = (value: 'top' | 'center' | 'bottom', icon: Parameters<typeof Icon>[0]['name']) => (
    <button
      className="wbe-btn wbe-icon"
      aria-pressed={vAlign === value}
      disabled={disabled}
      onClick={() => void store.formatCells({ vAlign: vAlign === value ? null : value })}
      title={t(`format.valign.${value}`)}
      aria-label={t(`format.valign.${value}`)}
    >
      <Icon name={icon} />
    </button>
  );
  return (
    <div className="wbe-format-bar" role="group" aria-label={t('format.label')} data-shown={store.formatBarShown || undefined}>
      <Select
        className="wbe-format-font"
        value={s.fontFamily}
        disabled={disabled}
        onChange={(e) => void store.formatCells({ fontFamily: e.target.value })}
        aria-label={t('format.font')}
        title={t('format.font')}
      >
        {(FONTS.includes(s.fontFamily) ? FONTS : [s.fontFamily, ...FONTS]).map((f) => (
          <option key={f} value={f}>
            {f}
          </option>
        ))}
      </Select>
      <Select
        className="wbe-format-size"
        value={String(sizePt)}
        disabled={disabled}
        onChange={(e) => void store.formatCells({ fontSizePt: Number(e.target.value) })}
        aria-label={t('format.size')}
        title={t('format.size')}
      >
        {(SIZES.includes(sizePt) ? SIZES : [...SIZES, sizePt].sort((a, b) => a - b)).map((n) => (
          <option key={n} value={String(n)}>
            {t.number(n)}
          </option>
        ))}
      </Select>
      <div className="wbe-btn-group">
        {toggle('bold', 'bold')}
        {toggle('italic', 'italic')}
        {toggle('underline', 'underline')}
        {toggle('strike', 'strike')}
      </div>
      <ColorPicker store={store} which="color" disabled={disabled} />
      <ColorPicker store={store} which="fill" disabled={disabled} />
      <button
        className="wbe-btn wbe-icon"
        aria-haspopup="menu"
        aria-expanded={!!borders}
        disabled={disabled}
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setBorders({ x: r.left, y: r.bottom + 2 });
        }}
        title={t('format.borders')}
        aria-label={t('format.borders')}
      >
        <Icon name="borders" />
        <span className="wbe-color-bar" style={{ background: pen.color ?? 'var(--wbe-text)' }} />
      </button>
      {borders && (
        <ContextMenu
          x={borders.x}
          y={borders.y}
          label={t('format.borders')}
          onClose={() => setBorders(null)}
          items={[
            ...BORDERS.map((where) => ({ label: t(`format.border.${where}`), onSelect: () => void store.drawBorders(where) })),
            { separator: true },
            {
              label: t('format.lineColor'),
              items: [
                { label: t('format.automatic'), checked: !pen.color, onSelect: () => store.setBorderPen({ color: undefined }) },
                ...Object.entries(COLORS).map(([name, hex]) => ({
                  label: t(`format.colors.${name as keyof typeof COLORS}`),
                  checked: pen.color?.toUpperCase() === hex,
                  sample: <span className="wbe-menu-swatch" style={{ background: hex }} aria-hidden="true" />,
                  onSelect: () => store.setBorderPen({ color: hex }),
                })),
              ],
            },
            {
              label: t('format.lineStyle'),
              items: LINE_STYLES.map(([name, line]) => ({
                label: t(`format.line.${name}`),
                checked: pen.width === line.width && pen.pattern === line.pattern,
                sample: <LineSample line={line} />,
                onSelect: () => store.setBorderPen(line),
              })),
            },
          ]}
        />
      )}
      <div className="wbe-btn-group">
        {align('left', 'alignLeft')}
        {align('center', 'alignCenter')}
        {align('right', 'alignRight')}
      </div>
      <div className="wbe-btn-group">
        {valign('top', 'alignTop')}
        {valign('center', 'alignMiddle')}
        {valign('bottom', 'alignBottom')}
      </div>
      <div className="wbe-btn-group">
        <button className="wbe-btn wbe-icon" disabled={disabled || !s.indent} onClick={() => void store.stepIndent(-1)} title={t('format.indentLess')} aria-label={t('format.indentLess')}>
          <Icon name="indentLess" />
        </button>
        <button className="wbe-btn wbe-icon" disabled={disabled || s.indent >= 15} onClick={() => void store.stepIndent(1)} title={t('format.indentMore')} aria-label={t('format.indentMore')}>
          <Icon name="indentMore" />
        </button>
      </div>
      <button className="wbe-btn wbe-icon" aria-pressed={s.wrap} disabled={disabled} onClick={() => void store.formatCells({ wrap: !s.wrap })} title={t('format.wrap')} aria-label={t('format.wrap')}>
        <Icon name="wrap" />
      </button>
      {/* Excel's Insert Checkbox; pressed again, Clear Formats' taking the checkboxes off and leaving TRUE and FALSE. */}
      <button
        className="wbe-btn wbe-icon"
        aria-pressed={!!s.checkbox}
        disabled={disabled}
        onClick={() => void (s.checkbox ? store.formatCells({ checkbox: false }) : store.insertCheckboxes())}
        title={t('format.checkbox')}
        aria-label={t('format.checkbox')}
      >
        <Icon name="checkbox" />
      </button>
      <Select
        className="wbe-format-number"
        value={kind ?? 'custom'}
        disabled={disabled}
        onChange={(e) =>
          e.target.value === 'more'
            ? store.set('numberFormatOpen', true)
            : void store.formatCells({ numFmt: numberFormatCode(e.target.value as NumberFormatKind, store.locale) })
        }
        aria-label={t('format.numberFormat')}
        title={t('format.numberFormat')}
      >
        {!kind && (
          <option value="custom" disabled>
            {t('format.number.custom')}
          </option>
        )}
        {NUMBER_FORMAT_KINDS.map((k) => (
          <option key={k} value={k}>
            {t(`format.number.${k}`)}
          </option>
        ))}
        <option value="more">{t('format.number.more')}</option>
      </Select>
      <div className="wbe-btn-group">
        <button className="wbe-btn wbe-icon" disabled={disabled} onClick={() => void store.stepDecimals(1)} title={t('format.decimalsMore')} aria-label={t('format.decimalsMore')}>
          <Icon name="decimalsMore" />
        </button>
        <button className="wbe-btn wbe-icon" disabled={disabled} onClick={() => void store.stepDecimals(-1)} title={t('format.decimalsFewer')} aria-label={t('format.decimalsFewer')}>
          <Icon name="decimalsFewer" />
        </button>
      </div>
      <button className="wbe-btn wbe-icon" disabled={disabled} onClick={() => void store.formatCells({}, { clear: true })} title={t('format.clear')} aria-label={t('format.clear')}>
        <Icon name="clearFormat" />
      </button>
    </div>
  );
}

/** Excel's line styles, in the order its Line Style menu lists them. */
type LineName = 'thin' | 'hair' | 'dotted' | 'dashDotDot' | 'dashDot' | 'dashed' | 'mediumDashDotDot' | 'slantDashDot' | 'mediumDashDot' | 'mediumDashed' | 'medium' | 'thick' | 'double';

const LINE_STYLES: [LineName, Pick<BorderEdge, 'width' | 'pattern'>][] = [
  ['thin', { width: 1, pattern: 'solid' }],
  ['hair', { width: 1, pattern: 'hair' }],
  ['dotted', { width: 1, pattern: 'dotted' }],
  ['dashDotDot', { width: 1, pattern: 'dashDotDot' }],
  ['dashDot', { width: 1, pattern: 'dashDot' }],
  ['dashed', { width: 1, pattern: 'dashed' }],
  ['mediumDashDotDot', { width: 2, pattern: 'dashDotDot' }],
  ['slantDashDot', { width: 2, pattern: 'slantDashDot' }],
  ['mediumDashDot', { width: 2, pattern: 'dashDot' }],
  ['mediumDashed', { width: 2, pattern: 'dashed' }],
  ['medium', { width: 2, pattern: 'solid' }],
  ['thick', { width: 3, pattern: 'solid' }],
  ['double', { width: 3, pattern: 'double' }],
];

/** A line style as the menu shows it: a short line drawn in it. */
function LineSample({ line }: { line: Pick<BorderEdge, 'width' | 'pattern'> }) {
  const style = line.pattern === 'solid' ? 'solid' : line.pattern === 'double' ? 'double' : line.pattern === 'dotted' || line.pattern === 'hair' ? 'dotted' : 'dashed';
  return <span className="wbe-menu-line" style={{ borderTop: `${line.pattern === 'hair' ? 1 : line.width}px ${style} currentColor` }} aria-hidden="true" />;
}

/** A font or fill colour: the colour shown under its button, and a palette that opens from it. */
function ColorPicker({ store, which, disabled }: { store: WorkbookStore; which: 'color' | 'fill'; disabled: boolean }) {
  const t = useT();
  /** Where the palette shows, under its button; outside the bar, which scrolls sideways and would cut it off. */
  const [open, setOpen] = useState<{ left: number; top: number } | null>(null);
  const button = useRef<HTMLButtonElement>(null);
  const panel = useRef<HTMLDivElement>(null);
  const id = useId();
  const close = () => setOpen(null);
  useDialog(panel, !!open, close);
  const current = store.activeStyle[which];
  const label = t(which === 'color' ? 'format.textColor' : 'format.fillColor');
  const pick = (value: string | null) => {
    close();
    void store.formatCells(which === 'color' ? { color: value } : { fill: value });
  };
  // The browser's own picker reports each colour passed over; only the one chosen counts.
  const more = useRef<HTMLInputElement>(null);
  const pickRef = useRef(pick);
  pickRef.current = pick;
  useEffect(() => {
    const input = more.current;
    if (!open || !input) return;
    const chosen = () => pickRef.current(input.value.toUpperCase());
    input.addEventListener('change', chosen);
    return () => input.removeEventListener('change', chosen);
  }, [open]);
  return (
    <div className="wbe-color-picker">
      <button
        ref={button}
        className="wbe-btn wbe-icon"
        aria-haspopup="dialog"
        aria-expanded={!!open}
        aria-controls={open ? id : undefined}
        disabled={disabled}
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          setOpen((o) => (o ? null : { left: Math.max(4, Math.min(r.left, window.innerWidth - 300)), top: r.bottom + 4 }));
        }}
        title={label}
        aria-label={label}
      >
        <Icon name={which === 'color' ? 'textColor' : 'fillColor'} />
        <span className="wbe-color-bar" style={{ background: current ?? (which === 'color' ? 'var(--wbe-text)' : 'transparent') }} />
      </button>
      {open && (
        <div ref={panel} id={id} className="wbe-color-panel" role="dialog" aria-label={label} style={open}>
          <button className="wbe-btn wbe-small" onClick={() => pick(null)}>
            {t(which === 'color' ? 'format.automatic' : 'format.noFill')}
          </button>
          <div className="wbe-color-swatches">
            {Object.entries(COLORS).map(([name, hex]) => (
              <button
                key={name}
                className="wbe-color-swatch"
                style={{ background: hex }}
                aria-label={t(`format.colors.${name as keyof typeof COLORS}`)}
                title={t(`format.colors.${name as keyof typeof COLORS}`)}
                aria-pressed={current?.toUpperCase() === hex}
                onClick={() => pick(hex)}
              />
            ))}
          </div>
          <label className="wbe-field wbe-small">
            <span>{t('format.moreColors')}</span>
            <input ref={more} type="color" defaultValue={current ?? '#000000'} />
          </label>
        </div>
      )}
    </div>
  );
}
