/**
 * Tests that need parts of Ruled's own grid not built yet, by title. With
 * E2E_GRID=ruled they are left out of the run until those parts are built;
 * the list has to be empty before Ruled's grid becomes the default.
 */
export const RULED_PENDING: RegExp[] = [
  // Filter buttons and checkboxes in cells.
  /opens a heading's filter menu from its button$/,
  /inserts Excel 365 checkboxes, ticks them by click and Space/,
];
