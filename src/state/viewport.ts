/** Phone-sized screens, where the side panel covers the sheet; the stylesheet uses the same width. */
export const NARROW = '(max-width: 720px)';

export function isNarrow(): boolean {
  return typeof matchMedia === 'function' && matchMedia(NARROW).matches;
}
