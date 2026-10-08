import type { Theme } from '@glideapps/glide-data-grid';

/**
 * The grid's colours and fonts. The grid is drawn on a canvas, which CSS
 * can't style, so these are read from the editor's --wbe-grid-* custom
 * properties when it renders (see editor.css), keeping the stylesheet the one
 * place a host restyles everything.
 */
export interface GridPalette {
  bg: string;
  bgAlt: string;
  text: string;
  textOnFill: string;
  textMuted: string;
  textLight: string;
  line: string;
  borderAuto: string;
  freezeLine: string;
  headerBg: string;
  headerText: string;
  headerLine: string;
  headerHoverBg: string;
  headerSelectedBg: string;
  headerSelectedText: string;
  rowHeaderText: string;
  selection: string;
  selectionFill: string;
  cutFill: string;
  onSelection: string;
  editedFill: string;
  editedMark: string;
  lockedFill: string;
  labelFill: string;
  errorMark: string;
  warningMark: string;
  pendingMark: string;
  noteMark: string;
  staleMark: string;
  unsupportedMark: string;
  dropdownArrow: string;
  font: string;
  headerFont: string;
  headerFontSize: string;
  headerFontWeight: string;
}

/** The defaults, the same as editor.css declares, for when no stylesheet has loaded. */
export const DEFAULT_PALETTE: GridPalette = {
  bg: '#ffffff',
  bgAlt: '#f7f7f7',
  text: '#1f1f1f',
  textOnFill: '#1f1f1f',
  textMuted: '#5f6368',
  textLight: '#9aa0a6',
  line: '#e1e3e3',
  borderAuto: '#000000',
  freezeLine: '#b7bcc2',
  headerBg: '#f4f5f5',
  headerText: '#4a4a4a',
  headerLine: '#d3d6d6',
  headerHoverBg: '#eaebeb',
  headerSelectedBg: '#e3efe8',
  headerSelectedText: '#1e7145',
  rowHeaderText: '#5f6368',
  selection: '#1e7145',
  selectionFill: 'rgba(30, 113, 69, 0.1)',
  cutFill: 'rgba(30, 113, 69, 0.06)',
  onSelection: '#ffffff',
  editedFill: '#fff4ce',
  editedMark: '#b8860b',
  lockedFill: '#f2f4f7',
  labelFill: '#eaf2ec',
  errorMark: '#d92d20',
  warningMark: '#f79009',
  pendingMark: '#2e90fa',
  noteMark: '#7f56d9',
  staleMark: '#e8a33d',
  unsupportedMark: '#8a94a6',
  dropdownArrow: '#5f6368',
  font: "Calibri, Carlito, 'Aptos Narrow', 'Segoe UI', system-ui, sans-serif",
  headerFont: "system-ui, -apple-system, 'Segoe UI', sans-serif",
  headerFontSize: '12px',
  headerFontWeight: '500',
};

/** Custom property for each palette entry: headerSelectedBg → --wbe-grid-header-selected-bg. */
const property = (key: string) => `--wbe-grid-${key.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)}`;

/** The palette as the stylesheet sets it on `el` (or on an ancestor). */
export function readGridPalette(el: Element | null): GridPalette {
  if (!el || typeof getComputedStyle === 'undefined') return DEFAULT_PALETTE;
  const css = getComputedStyle(el);
  const out = { ...DEFAULT_PALETTE };
  for (const key of Object.keys(out) as (keyof GridPalette)[]) {
    // The selection's text colour follows the brand's.
    const name = key === 'onSelection' ? '--wbe-on-accent' : property(key);
    const value = css.getPropertyValue(name).trim();
    if (value) out[key] = value;
  }
  return out;
}

/** The colours references in a formula being edited take in turn, the same as editor.css declares. */
export const DEFAULT_REF_COLORS: readonly string[] = ['#1f5fc8', '#c42b1c', '#7a3db8', '#1b7a43', '#b0287a', '#0b7285', '#a15c00'];

/** The reference colours as the stylesheet sets them on `el`: --wbe-formula-ref1, --wbe-formula-ref2… */
export function readRefColors(el: Element | null): readonly string[] {
  if (!el || typeof getComputedStyle === 'undefined') return DEFAULT_REF_COLORS;
  const css = getComputedStyle(el);
  return DEFAULT_REF_COLORS.map((fallback, i) => css.getPropertyValue(`--wbe-formula-ref${i + 1}`).trim() || fallback);
}

/** Glide's theme from the palette. Grid lines are drawn with the cells, so Glide's own are off. */
export function glideTheme(p: GridPalette): Partial<Theme> {
  return {
    accentColor: p.selection,
    accentFg: p.onSelection,
    accentLight: p.selectionFill,
    textDark: p.text,
    textMedium: p.textMuted,
    textLight: p.textLight,
    textHeader: p.headerText,
    textHeaderSelected: p.headerSelectedText,
    bgCell: p.bg,
    bgCellMedium: p.bgAlt,
    bgHeader: p.headerBg,
    bgHeaderHasFocus: p.headerSelectedBg,
    bgHeaderHovered: p.headerHoverBg,
    borderColor: 'transparent',
    horizontalBorderColor: 'transparent',
    headerBottomBorderColor: p.headerLine,
    headerFontStyle: `${p.headerFontWeight} ${p.headerFontSize}`,
    baseFontStyle: '14.7px',
    fontFamily: p.font,
    editorFontSize: '14.7px',
    cellHorizontalPadding: 4,
    cellVerticalPadding: 2,
    lineHeight: 1.25,
  };
}
