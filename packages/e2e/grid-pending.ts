/**
 * Tests that need parts of Ruled's own grid not built yet, by title. With
 * E2E_GRID=ruled they are left out of the run until those parts are built;
 * the list has to be empty before Ruled's grid becomes the default.
 */
export const RULED_PENDING: RegExp[] = [
  // Dragging the selection's border, and resizing rows and columns by their edges.
  /moves cells by dragging the selection's border/,
  /moves cells in between others with Shift held/,
  /resizes a row by dragging the edge of its number$/,
  /keeps a column's width when its edge is only pressed/,
  /fits a column to its widest value when its edge is double-clicked/,
  // Filter buttons and checkboxes in cells.
  /opens a heading's filter menu from its button$/,
  /inserts Excel 365 checkboxes, ticks them by click and Space/,
];
