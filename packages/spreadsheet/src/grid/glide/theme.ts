import type { Theme } from '@glideapps/glide-data-grid';
import type { GridPalette } from '../../ui/gridPalette';

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
