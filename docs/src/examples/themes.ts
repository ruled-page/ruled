/**
 * The editor in the docs' own colours: ledger green on warm paper, and a dark
 * version on green-black, set through the editor's theme tokens.
 */
import type { EditorTheme } from '@sitepen/spreadsheet';
import { DEMO_THEMES } from '../../../examples/demo/src/themes';

const fonts: EditorTheme = {
  font: "'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif",
  fontMono: "'IBM Plex Mono', ui-monospace, monospace",
  radius: '6px',
  radiusSm: '4px',
  radiusLg: '8px',
};

export const ledgerLight: EditorTheme = {
  ...fonts,
  accent: '#0b6b4f',
  accentStrong: '#085540',
  accentSoft: '#dcefe5',
  accentBorder: '#a8d4c0',
  accentTint: 'rgba(11, 107, 79, 0.12)',
  focusRing: 'rgba(11, 107, 79, 0.35)',
  surface: '#fffdf8',
  surfaceAlt: '#f7f3ea',
  surfaceSunken: '#efe9dc',
  text: '#16211c',
  textMuted: '#4c5a52',
  border: '#e3ddcf',
  borderStrong: '#d2cbbb',
  danger: '#d63c25',
  gridBg: '#fffdf8',
  gridLine: '#e8e2d5',
  gridHeaderBg: '#f4efe4',
  gridHeaderLine: '#ddd6c6',
  gridHeaderText: '#4c5a52',
  gridHeaderSelectedBg: '#dcefe5',
  gridSelection: '#0b6b4f',
  gridSelectionFill: 'rgba(11, 107, 79, 0.08)',
};

const demoDark = DEMO_THEMES.find((t) => t.id === 'dark')?.theme ?? {};

export const ledgerDark: EditorTheme = {
  ...demoDark,
  ...fonts,
  accent: '#3ecf9a',
  accentStrong: '#62dcae',
  accentSoft: '#10392c',
  accentBorder: '#245c48',
  accentTint: 'rgba(62, 207, 154, 0.16)',
  onAccent: '#0e1411',
  surface: '#151d19',
  surfaceAlt: '#1a241f',
  surfaceSunken: '#1f2a24',
  border: '#26352d',
  borderStrong: '#33453b',
  gridBg: '#121a16',
  gridBgAlt: '#16201b',
  gridLine: '#24302a',
  gridHeaderBg: '#18221d',
  gridHeaderLine: '#2b3a32',
  gridSelection: '#3ecf9a',
};

/** The docs' theme for the page's light or dark mode. */
export function ledgerTheme(): EditorTheme {
  return document.documentElement.dataset.theme === 'light' ? ledgerLight : ledgerDark;
}

const LABELS: Record<string, string> = { indigo: 'Indigo', paper: 'Paper', contrast: 'High contrast' };

/** Themes the theming example switches between. */
export const PLAYGROUND_THEMES: { id: string; label: string; theme: () => EditorTheme }[] = [
  { id: 'ledger', label: 'Ledger (these docs)', theme: ledgerTheme },
  { id: 'default', label: 'Default', theme: () => ({}) },
  ...DEMO_THEMES.filter((t) => t.theme && t.id !== 'dark').map((t) => ({ id: t.id, label: LABELS[t.id] ?? t.id, theme: () => t.theme! })),
  { id: 'dark', label: 'Dark', theme: () => demoDark },
];
