import { useEffect, useRef, type RefObject } from 'react';
import { clipboardNumbers, parseClipboardText } from '../state/clipboard';
import { clipboardFormats } from '../state/clipboardFormats';
import type { WorkbookStore } from '../state/store';
import type { Box } from './contract';

/**
 * What a grid does with the clipboard and a few keys while the focus is in
 * it (`host`): it writes the clipboard itself, the text and an escaped HTML
 * table, rather than letting a grid write cell text into HTML unescaped,
 * and remembers what was copied, so pasting it back carries formulas. A cut
 * only marks the cells: like Excel, pasting it moves them, and references
 * follow. Pastes are read from the paste event, which already carries the
 * data, so browsers don't ask for clipboard permission. Escape forgets a
 * cut; the menu key or Shift+F10 opens the cell menu at the active cell,
 * and Shift+F2 its note or comments. `cellBox` says where a cell is drawn;
 * `ownField` whether a text field is the grid's own, not the formula bar's
 * or a dialog's.
 */
export function useSheetEvents(store: WorkbookStore, host: RefObject<HTMLElement | null>, cellBox: (row: number, col: number) => Box | undefined, ownField: (el: Element) => boolean): void {
  const boxOf = useRef(cellBox);
  boxOf.current = cellBox;
  const own = useRef(ownField);
  own.current = ownField;
  useEffect(() => {
    const inEditorField = () => {
      const active = document.activeElement;
      return !!active && /^(INPUT|TEXTAREA)$/.test(active.tagName) && !own.current(active);
    };
    const focused = () => !!host.current?.contains(document.activeElement);
    const write = (e: ClipboardEvent, cut: boolean) => {
      if (inEditorField() || !focused()) return;
      const first = store.selection.rects[0];
      if (!first || !e.clipboardData) return;
      e.preventDefault();
      e.clipboardData.setData('text/plain', store.selectionText(first));
      e.clipboardData.setData('text/html', store.selectionHtml(first));
      store.noteCopy(first, cut);
    };
    const onCopy = (e: ClipboardEvent) => write(e, false);
    const onCut = (e: ClipboardEvent) => write(e, true);
    const onPaste = (e: ClipboardEvent) => {
      if (!focused()) return;
      const text = e.clipboardData?.getData('text/plain');
      if (!text) return;
      e.preventDefault();
      const values = parseClipboardText(text);
      // Excel's HTML has the exact values of the numbers its text shows rounded, and each cell's formats.
      const html = e.clipboardData?.getData('text/html');
      const exact = html ? clipboardNumbers(html) : undefined;
      const formats = html ? clipboardFormats(html) : undefined;
      const sel = store.selection;
      const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
      void store.paste({ row: rect.r1, col: rect.c1 }, values, rect, false, exact, formats);
    };
    const onKey = (e: KeyboardEvent) => {
      if (!focused()) return;
      if (e.key === 'Escape' && store.cutRange) store.cancelCut();
      const sel = store.selection;
      if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) {
        e.preventDefault();
        const at = boxOf.current(sel.row, sel.col);
        if (at) store.set('menu', { x: at.x + Math.min(at.width, 24), y: at.y + at.height, kind: 'cell' });
      }
      if (e.shiftKey && e.key === 'F2' && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        const at = boxOf.current(sel.row, sel.col);
        if (at) store.openComment(sel.row, sel.col, { x: at.x + at.width + 4, y: at.y });
      }
    };
    window.addEventListener('copy', onCopy, true);
    window.addEventListener('cut', onCut, true);
    window.addEventListener('paste', onPaste, true);
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('copy', onCopy, true);
      window.removeEventListener('cut', onCut, true);
      window.removeEventListener('paste', onPaste, true);
      window.removeEventListener('keydown', onKey);
    };
  }, [store, host]);
}
