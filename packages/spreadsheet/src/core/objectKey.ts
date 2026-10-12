import type { SheetObject } from './types';

/** An object's key, the same whatever else changes: where it is in the file, or its name for one added in the editor. */
export function objectKey(obj: SheetObject): string {
  return obj.source ? `${obj.source.drawing}#${obj.source.anchor}` : `new:${obj.name}`;
}
