import { describe, expect, it } from 'vitest';
import type { OutlineAxis } from '../src/core/types';
import { clearOutline, collapseGroup, expandGroup, outlineBarSize, outlineDepth, outlineGroups, regroup, showLevel, summaryLine } from '../src/state/outline';
import { english } from '../src/i18n';

const axis = (levels: Record<number, number>, collapsed: number[] = [], summaryAfter = true): OutlineAxis => ({ levels, collapsed, summaryAfter });

describe('outline groups', () => {
  it('are the runs of lines at each level or deeper, outer before inner', () => {
    const o = axis({ 1: 1, 2: 2, 3: 2, 4: 1, 5: 3, 7: 1, 8: 1 });
    expect(outlineGroups(o)).toEqual([
      { level: 1, start: 1, end: 5 },
      { level: 2, start: 2, end: 3 },
      { level: 2, start: 5, end: 5 },
      { level: 3, start: 5, end: 5 },
      // A gap ends a group; the lines after it start another.
      { level: 1, start: 7, end: 8 },
    ]);
    expect(outlineDepth(o)).toBe(3);
    expect([outlineBarSize(o), outlineBarSize(axis({})), outlineBarSize(undefined)]).toEqual([68, 0, 0]);
  });

  it('are summed up below or right of them, or above or left when the sheet says so', () => {
    const g = { level: 1, start: 3, end: 6 };
    expect([summaryLine(axis({}), g), summaryLine(axis({}, [], false), g)]).toEqual([7, 2]);
  });
});

describe('outline changes', () => {
  it('group a level deeper, ungroup a level out, and refuse past seven or with nothing grouped', () => {
    expect(regroup(axis({ 2: 1 }), 'rows', [1, 2], 1)).toEqual({ levels: { 1: 1, 2: 2 }, collapsed: {}, hidden: {} });
    expect(regroup(axis({ 2: 1 }), 'rows', [1, 2], -1)).toEqual({ levels: { 2: 0 }, collapsed: {}, hidden: {} });
    const refused = (r: ReturnType<typeof regroup>) => ('refused' in r ? english.text(r.refused) : '');
    expect(refused(regroup(axis({ 2: 7 }), 'rows', [2], 1))).toBe('Groups can go at most 7 levels deep, as in Excel.');
    expect(refused(regroup(axis({}), 'cols', [2], -1))).toBe("The selected columns aren't in a group.");
  });

  it('collapse a group and expand it again, with a group inside it summed up on the same line opening with it', () => {
    // Rows 1-4 at level 1, and 3-4 inside them at level 2: both are summed up on row 5.
    const o = axis({ 1: 1, 2: 1, 3: 2, 4: 2 });
    const outer = outlineGroups(o)[0];
    expect(collapseGroup(o, new Set([3, 4]), outer)).toEqual({ levels: {}, collapsed: { 5: true }, hidden: { 1: true, 2: true } });
    expect(expandGroup(axis(o.levels, [5]), new Set([1, 2, 3, 4]), outer)).toEqual({ levels: {}, collapsed: { 5: false }, hidden: { 1: false, 2: false, 3: false, 4: false } });
  });

  it('show to a level, collapsing the groups that deep or deeper, here summed up above them', () => {
    const o = axis({ 1: 1, 2: 1, 3: 2, 4: 2 }, [], false);
    // Summed up above: the outer group on row 0, the inner one on row 2.
    expect(showLevel(o, new Set(), 2)).toEqual({ levels: {}, collapsed: { 2: true }, hidden: { 3: true, 4: true } });
    expect(showLevel(o, new Set([3, 4]), 1)).toEqual({ levels: {}, collapsed: { 0: true, 2: true }, hidden: { 1: true, 2: true } });
  });

  it('clear every level and mark', () => {
    expect(clearOutline(axis({ 1: 1, 2: 2 }, [3]))).toEqual({ levels: { 1: 0, 2: 0 }, collapsed: { 3: false }, hidden: {} });
  });
});
