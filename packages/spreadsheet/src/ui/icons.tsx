import { createContext, useContext, type ReactNode } from 'react';

const PATHS = {
  chart: 'M4 20V10M10 20V4M16 20v-7M22 20H2',
  image: 'M4 5h16v14H4zM4 15l4-4 4 4 3-3 5 5M15 9.5a1.5 1.5 0 1 0 0-.01',
  shape: 'M4 4h10v10H4zM14 10a6 6 0 1 1-4 9.5',
  undo: 'M9 14 4 9l5-5M4 9h10a6 6 0 0 1 0 12h-3',
  redo: 'M15 14l5-5-5-5M20 9H10a6 6 0 0 0 0 12h3',
  download: 'M12 3v12M7 10l5 5 5-5M4 21h16',
  open: 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z',
  info: 'M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 16v-5M12 8h.01',
  warn: 'M12 3 2 20h20zM12 10v4M12 17h.01',
  close: 'M6 6l12 12M18 6 6 18',
  check: 'M5 12l5 5 9-10',
  lock: 'M6 11h12v10H6zM8 11V7a4 4 0 0 1 8 0v4',
  sheet: 'M4 4h16v16H4zM4 10h16M4 15h16M10 4v16',
  edit: 'M4 20h4L19 9l-4-4L4 16zM14 6l4 4',
  formula: 'M14 4h-2a3 3 0 0 0-3 3v13M6 10h7M15 13l5 6M20 13l-5 6',
  eye: 'M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12zM12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z',
  plus: 'M12 5v14M5 12h14',
  minus: 'M5 12h14',
  schema: 'M4 5h6v5H4zM14 14h6v5h-6zM7 10v6.5h7M14 7.5h3.5V14',
  upload: 'M12 21V9M7 14l5-5 5 5M4 3h16',
  sparkle: 'M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z',
  trash: 'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13',
  note: 'M4 5h16v11H9l-5 4z',
  link: 'M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1',
  keyboard: 'M3 6h18v12H3zM7 10h.01M11 10h.01M15 10h.01M7 14h10',
  more: 'M5 12h.01M12 12h.01M19 12h.01',
  bold: 'M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z',
  italic: 'M10 4h8M6 20h8M14 4l-4 16',
  underline: 'M7 4v7a5 5 0 0 0 10 0V4M5 21h14',
  strike: 'M16.5 7A4.5 4.5 0 0 0 12 4h-.5A3.5 3.5 0 0 0 9 10M4 12h16M8 17a4 4 0 0 0 4 3h.5A3.5 3.5 0 0 0 15 14',
  textColor: 'M6 17 12 4l6 13M8.3 12h7.4',
  fillColor: 'M5 11l6-6 7 7-6 6zM11 5 8 2M5 11h13M20 15s2 2.4 2 3.6a2 2 0 0 1-4 0c0-1.2 2-3.6 2-3.6',
  borders: 'M4 4h16v16H4zM4 12h16M12 4v16',
  alignLeft: 'M4 6h16M4 10h10M4 14h16M4 18h10',
  alignCenter: 'M4 6h16M7 10h10M4 14h16M7 18h10',
  alignRight: 'M4 6h16M10 10h10M4 14h16M10 18h10',
  wrap: 'M4 6h16M4 12h13a3 3 0 0 1 0 6h-4M15 16l-2 2 2 2M4 18h5',
  checkbox: 'M5 4h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1zM8 12l3 3 5-6',
  decimalsMore: 'M3 18h.01M7 13a2.5 2.5 0 0 1 5 0v3a2.5 2.5 0 0 1-5 0zM18 20V8M15 11l3-3 3 3',
  decimalsFewer: 'M3 18h.01M7 13a2.5 2.5 0 0 1 5 0v3a2.5 2.5 0 0 1-5 0zM18 8v12M15 17l3 3 3-3',
  clearFormat: 'M4 7V4h13v3M10.5 4v16M8 20h5M15 15l5 5M20 15l-5 5',
  chevronDown: 'M6 9l6 6 6-6',
  chevronUp: 'M6 15l6-6 6 6',
  alignTop: 'M4 4h16M9 8h6v12H9z',
  alignMiddle: 'M4 12h4M16 12h4M9 6h6v12H9z',
  alignBottom: 'M4 20h16M9 4h6v12H9z',
  indentMore: 'M4 5h16M11 10h9M11 14h9M4 19h16M4 9l3 3-3 3',
  indentLess: 'M4 5h16M11 10h9M11 14h9M4 19h16M7 9l-3 3 3 3'
} satisfies Record<string, string>;

export type IconName = keyof typeof PATHS;

/** Replacement icons: SVG path data drawn on a 24×24 grid with the current text colour, or any element. */
export type IconOverrides = Partial<Record<IconName, string | ReactNode>>;

/** The host's replacement icons, provided by the editor's root. */
export const IconContext = createContext<IconOverrides | undefined>(undefined);

export function Icon({ name, size = 16 }: { name: IconName; size?: number }) {
  const custom = useContext(IconContext)?.[name];
  if (custom !== undefined && typeof custom !== 'string') {
    return (
      <span className="wbe-icon" style={{ width: size, height: size }} aria-hidden="true">
        {custom}
      </span>
    );
  }
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      style={{ strokeWidth: 'var(--wbe-icon-stroke-width, 1.8)' }}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d={typeof custom === 'string' ? custom : PATHS[name]} />
    </svg>
  );
}

/**
 * Ruled's mark, for the toolbar's brand tile: a ledger page's rules, its
 * margin line, and an R written on the bottom rule, in the tile's colours.
 */
export function BrandMark({ size = 26 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="4 4 56 56" aria-hidden="true">
      <path d="M4 23H60M4 35H60M4 47H60" stroke="currentColor" strokeWidth="2.5" opacity=".38" />
      <path d="M18 4V60" stroke="var(--wbe-danger)" strokeWidth="3.5" />
      <path transform="translate(23.86 47) scale(0.022143)" fill="currentColor" d="M1305 -909Q1305 -813 1264 -734Q1223 -655 1148 -601Q1072 -548 969 -528Q932 -522 903 -515Q873 -509 841 -505Q808 -500 761 -500Q693 -500 640 -510Q587 -519 545 -535Q503 -552 470 -575L475 -659Q507 -636 540 -624Q573 -611 609 -606Q645 -600 684 -600Q806 -600 876 -675Q946 -749 946 -900Q946 -1018 903 -1110Q859 -1202 784 -1254Q709 -1307 612 -1307H549V-170Q549 -144 559 -130Q569 -116 594 -107L648 -94Q666 -86 675 -75Q683 -63 683 -45Q683 0 632 0H111Q86 0 73 -12Q60 -24 60 -45Q60 -78 95 -94L148 -107Q173 -116 184 -130Q194 -144 194 -170V-1238Q194 -1260 183 -1272Q173 -1284 148 -1293L95 -1306Q60 -1322 60 -1355Q60 -1376 73 -1388Q86 -1400 111 -1400H621Q838 -1400 991 -1338Q1144 -1275 1224 -1165Q1305 -1054 1305 -909ZM713 -553 1048 -593 1289 -160Q1306 -131 1323 -117Q1341 -103 1370 -93Q1389 -85 1397 -74Q1405 -64 1405 -45Q1405 -24 1393 -12Q1380 0 1354 0H896Q847 0 847 -45Q847 -59 853 -67Q859 -75 871 -82L903 -92Q922 -100 927 -114Q932 -127 917 -156Z" />
    </svg>
  );
}
