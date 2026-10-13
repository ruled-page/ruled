// Parts of Glide's DOM the editor depends on. None are in Glide's API, so
// tests/glide-dom.test.ts checks that the installed Glide still has them.
// Cell editors open in the editor's own portal (PORTAL_ID in ../contract).

/** The element each grid scrolls in. */
export const GLIDE_SCROLLER = 'dvn-scroller';

/** The box an open cell editor sits in. */
export const GLIDE_CLIP_REGION = 'gdg-clip-region';

/** Clicks inside an element with this class don't close an open cell editor, as clicks elsewhere do. */
export const GLIDE_CLICK_OUTSIDE_IGNORE = 'click-outside-ignore';

