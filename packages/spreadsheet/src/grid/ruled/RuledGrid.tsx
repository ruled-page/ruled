import { useT } from '../../i18n/react';
import type { SurfaceProps } from '../contract';

/** Ruled's own grid, which is still being built: for now it says so, and gives no handle. */
export function SheetSurface({ view }: SurfaceProps) {
  const t = useT();
  return (
    <div className="wbe-grid-host" data-testid="wbe-grid">
      <p className="wbe-grid-unbuilt" role="status">
        {t('grid.unbuilt', { name: view.info.name })}
      </p>
    </div>
  );
}
