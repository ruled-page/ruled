import { linkAddress } from '../core/links';
import type { CellLink } from '../state/links';
import { useT } from '../i18n/react';
import { APPLE, mod } from './keys';

/**
 * Where a cell's link goes, shown under the cell while the pointer is over
 * it, as Excel does, with its screen tip and how to follow it. The address is
 * always shown in full: the workbook is untrusted, and a screen tip or the
 * cell's text can say anything.
 */
export function LinkTip({ link, x, y, follows }: { link: CellLink; x: number; y: number; follows: boolean }) {
  const t = useT();
  const address = linkAddress(link);
  // Keep the tip on screen when the cell is near the right or bottom edge.
  const left = Math.max(4, Math.min(x, window.innerWidth - 336));
  const top = Math.min(y, window.innerHeight - 96);
  return (
    <div className="wbe-link-tip" style={{ left, top }} role="tooltip">
      {link.tooltip && <div className="wbe-link-tip-title">{link.tooltip}</div>}
      {address && <div className="wbe-link-tip-address">{address}</div>}
      {follows && <div className="wbe-link-tip-hint">{t(APPLE ? 'links.tip.followApple' : 'links.tip.follow', { key: mod(t, t('keys.enter')) })}</div>}
    </div>
  );
}
