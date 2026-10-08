// Parts of Glide's DOM the editor depends on. None are in Glide's API, so
// tests/glide-dom.test.ts checks that the installed Glide still has them.

/** The element each grid scrolls in. */
export const GLIDE_SCROLLER = 'dvn-scroller';

/** The box an open cell editor sits in. */
export const GLIDE_CLIP_REGION = 'gdg-clip-region';

/**
 * The id of the element Glide opens cell editors in. Glide itself looks for
 * "portal"; vite.glide.ts changes that to this, so the editor doesn't take
 * over an element of the host page's.
 */
export const PORTAL_ID = 'wbe-portal';
