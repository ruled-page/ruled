/**
 * The grid library's image cell editor imports react-responsive-carousel, a
 * CommonJS module whose require('react') can't work in the React build. The
 * editor never shows image cells, so the carousel is never drawn.
 */
export function Carousel(): null {
  return null;
}
