import { afterEach, describe, expect, it, vi } from 'vitest';

/** A browser whose fonts are `installed`: text measures by the first of them its font asks for, or by the generic font. */
function browser(installed: string[], fails = false) {
  const faces: { family: string; source: string; style?: string; weight?: string; unicodeRange?: string }[] = [];
  const ctx = {
    font: '',
    measureText(text: string) {
      const families = this.font.replace(/^\d+px /, '').split(',').map((f) => f.trim().replace(/"/g, ''));
      const used = families.find((f) => installed.includes(f)) ?? families.at(-1)!;
      return { width: text.length * (used.length + 1) };
    },
  };
  vi.stubGlobal('document', { createElement: () => ({ getContext: () => ctx }), fonts: { add: (f: (typeof faces)[number]) => faces.push(f) } });
  vi.stubGlobal(
    'FontFace',
    class {
      constructor(family: string, source: string, d: { style: string; weight: string; unicodeRange: string }) {
        Object.assign(this, { family, source, ...d });
      }
      load = () => (fails ? Promise.reject(new Error('offline')) : Promise.resolve(this));
    },
  );
  return faces;
}

/** The module afresh, with nothing yet known about the device. */
const fonts = async () => {
  vi.resetModules();
  return import('../src/grid/standInFonts');
};

afterEach(() => vi.unstubAllGlobals());

describe("stand-ins for Excel's fonts", () => {
  it('draws in the font itself where the device has it, or where there is no browser to ask', async () => {
    browser(['Calibri']);
    expect((await fonts()).standIn('Calibri')).toBeUndefined();
    vi.unstubAllGlobals();
    expect((await fonts()).standIn('Calibri')).toBeUndefined();
  });

  it('fetches a free stand-in drawn to the same widths where the device lacks the font, once, and says when it has arrived', async () => {
    const faces = browser([]);
    const { standIn, onStandInLoaded } = await fonts();
    const arrived = new Promise<void>((resolve) => onStandInLoaded(resolve));
    expect(standIn('Calibri')).toBe('Carlito');
    expect(standIn('calibri')).toBe('Carlito');
    await arrived;
    expect(faces.map((f) => [f.family, f.style, f.weight])).toEqual([
      ['Carlito', 'normal', '400'],
      ['Carlito', 'normal', '700'],
      ['Carlito', 'italic', '400'],
      ['Carlito', 'italic', '700'],
    ]);
    expect(faces.every((f) => /^url\(.*carlito-latin-[47]00-(normal|italic)\.woff2/.test(f.source) && f.unicodeRange?.startsWith('U+0000-00FF'))).toBe(true);
    expect(standIn('Calibri')).toBe('Carlito');
    expect(faces).toHaveLength(4);
    expect([standIn('Cambria'), standIn('Georgia'), standIn('Arial'), standIn('Times New Roman'), standIn('Courier New')]).toEqual(['Caladea', 'Gelasio', 'Arimo', 'Tinos', 'Cousine']);
  });

  it('has none for fonts with no free match, and says so when one fails to arrive', async () => {
    browser([], true);
    const { standIn, onStandInLoaded } = await fonts();
    expect([standIn('Aptos Narrow'), standIn('Verdana')]).toEqual([undefined, undefined]);
    const heard = vi.fn();
    const stop = onStandInLoaded(heard);
    standIn('Cambria');
    await vi.waitFor(() => expect(heard).toHaveBeenCalledOnce());
    stop();
  });
});
