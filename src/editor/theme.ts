/**
 * Theming. The editor's look comes entirely from CSS custom properties on its
 * root element (.wbe), declared with their defaults in editor.css: colours,
 * fonts, type sizes, corners, shadows, and the grid's colours and fonts (the
 * grid is drawn on a canvas and reads them when it renders). Override them in
 * CSS, or pass `theme` in the options, where each token is the property's
 * name in camelCase: `gridHeaderBg` sets --wbe-grid-header-bg.
 */

/** Every theme token, as `theme` takes it. */
export const THEME_TOKENS = [
  'colorScheme',
  'accent',
  'accentStrong',
  'accentSoft',
  'accentBorder',
  'accentTint',
  'onAccent',
  'surface',
  'surfaceAlt',
  'surfaceSunken',
  'backdrop',
  'focusRing',
  'text',
  'textMuted',
  'textSubtle',
  'border',
  'borderStrong',
  'neutral',
  'success',
  'successSoft',
  'successBorder',
  'successText',
  'warning',
  'warningSoft',
  'warningBorder',
  'warningText',
  'warningAccent',
  'danger',
  'dangerSoft',
  'dangerBorder',
  'onDanger',
  'infoSoft',
  'infoText',
  'infoAccent',
  'infoBannerBg',
  'infoBannerBorder',
  'infoBannerText',
  'fieldSoft',
  'fieldText',
  'noteSoft',
  'noteText',
  'noteBorder',
  'noteTipBg',
  'noteTipBorder',
  'placeholderBg',
  'placeholderBorder',
  'placeholderText',
  'placeholderTextMuted',
  'dropzoneBorder',
  'dropzoneHoverBg',
  'gridBg',
  'gridText',
  'gridTextOnFill',
  'gridLine',
  'gridBorderAuto',
  'gridFreezeLine',
  'gridHeaderBg',
  'gridHeaderText',
  'gridHeaderLine',
  'gridHeaderHoverBg',
  'gridHeaderSelectedBg',
  'gridHeaderSelectedText',
  'gridRowHeaderText',
  'gridTextMuted',
  'gridTextLight',
  'gridBgAlt',
  'gridCutFill',
  'gridSelection',
  'gridSelectionFill',
  'gridEditedFill',
  'gridLockedFill',
  'gridLabelFill',
  'gridErrorMark',
  'gridWarningMark',
  'gridPendingMark',
  'gridNoteMark',
  'gridStaleMark',
  'gridUnsupportedMark',
  'gridDropdownArrow',
  'gridFont',
  'gridHeaderFont',
  'gridHeaderFontSize',
  'gridHeaderFontWeight',
  'chartColors',
  'font',
  'fontMono',
  'fontSize',
  'fontSizeXs',
  'fontSizeSm',
  'fontSizeMd',
  'fontSizeLg',
  'fontSizeXl',
  'iconStrokeWidth',
  'radius',
  'radiusSm',
  'radiusLg',
  'shadow',
  'shadowSm',
] as const;

export type ThemeToken = (typeof THEME_TOKENS)[number];

export type EditorTheme = Partial<Record<ThemeToken, string>>;

/** The custom property behind a token: gridHeaderBg → --wbe-grid-header-bg. */
export function themeProperty(token: ThemeToken | string): string {
  return `--wbe-${token.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)}`;
}

/**
 * Copies the theme as `from` resolves it onto `to`. Cell editors open in a
 * portal at the end of the page, outside the editor, so they get the theme
 * of whichever editor opened them this way.
 */
export function copyTheme(from: Element, to: HTMLElement): void {
  const style = getComputedStyle(from);
  for (const token of THEME_TOKENS) {
    const name = themeProperty(token);
    to.style.setProperty(name, style.getPropertyValue(name));
  }
}

/** The theme as inline custom properties for the editor's root element. */
export function themeStyle(theme: EditorTheme | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [token, value] of Object.entries(theme ?? {})) if (value) out[themeProperty(token)] = value;
  return out;
}
