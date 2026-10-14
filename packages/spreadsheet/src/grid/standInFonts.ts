/**
 * Free fonts drawn to the same widths as Excel's own, for devices without
 * them. Calibri and Cambria are Microsoft's, so a Mac, a phone or Linux
 * usually lacks them, and a stand-in of other widths wouldn't fit text to its
 * columns as Excel does. Each is a chunk of its own, fetched only when a
 * workbook uses its font on a device without it. They cover Latin letters;
 * other letters are drawn in the device's own fonts.
 */

/** A font's files, in each style the grid draws. */
export interface FaceFiles {
  regular: string;
  bold: string;
  italic: string;
  boldItalic: string;
}

/** The characters the fonts' Latin files cover. */
const LATIN = 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD';

/** Excel's fonts that have a free stand-in, by name in lower case. */
const STAND_INS: Record<string, { family: string; files: () => Promise<{ default: FaceFiles }> }> = {
  calibri: { family: 'Carlito', files: () => import('./fonts/carlito') },
  cambria: { family: 'Caladea', files: () => import('./fonts/caladea') },
  georgia: { family: 'Gelasio', files: () => import('./fonts/gelasio') },
  arial: { family: 'Arimo', files: () => import('./fonts/arimo') },
  'times new roman': { family: 'Tinos', files: () => import('./fonts/tinos') },
  'courier new': { family: 'Cousine', files: () => import('./fonts/cousine') },
};

/** Whether each font is on the device, or its stand-in is loading, ready or failed to load. */
const state = new Map<string, 'installed' | 'loading' | 'ready' | 'failed'>();
const listeners = new Set<() => void>();

/** Whether the device has a font: text asked for in it measures other than in either generic font it would otherwise fall back to. */
function installed(family: string): boolean {
  const ctx = typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
  if (!ctx) return true;
  const width = (font: string) => {
    ctx.font = `40px ${font}`;
    return ctx.measureText('mmmmmmmmmmlli 0123456789').width;
  };
  return ['monospace', 'serif'].some((generic) => width(`"${family}", ${generic}`) !== width(generic));
}

async function load(key: string, family: string, files: () => Promise<{ default: FaceFiles }>): Promise<void> {
  try {
    const { regular, bold, italic, boldItalic } = (await files()).default;
    const faces = (
      [
        [regular, 'normal', '400'],
        [bold, 'normal', '700'],
        [italic, 'italic', '400'],
        [boldItalic, 'italic', '700'],
      ] as const
    ).map(([url, style, weight]) => new FontFace(family, `url(${url})`, { style, weight, unicodeRange: LATIN }));
    for (const face of faces) document.fonts.add(face);
    await Promise.all(faces.map((face) => face.load()));
    state.set(key, 'ready');
  } catch {
    // Text stays in the device's own font.
    state.set(key, 'failed');
  }
  for (const fn of listeners) fn();
}

/**
 * The stand-in to draw a font with, when the device lacks it and there is
 * one: fetched the first time it's asked for, and drawn with once it
 * arrives (see onStandInLoaded).
 */
export function standIn(family: string): string | undefined {
  const key = family.toLowerCase();
  const font = STAND_INS[key];
  if (!font) return undefined;
  let now = state.get(key);
  if (!now) {
    now = installed(family) ? 'installed' : 'loading';
    state.set(key, now);
    if (now === 'loading') void load(key, font.family, font.files);
  }
  return now === 'installed' ? undefined : font.family;
}

/** Calls `fn` each time a stand-in has loaded, or failed to, so text measured before can be measured again; until the function returned is called. */
export function onStandInLoaded(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}
