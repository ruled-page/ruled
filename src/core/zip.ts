import { strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { attr, children, parseXml } from './xml';

export interface Relationship {
  id: string;
  /** Last segment of the relationship type URI, e.g. "worksheet", "chart". */
  type: string;
  /** Resolved package path (no leading slash), or the raw target when external. */
  target: string;
  external: boolean;
}

const decoder = new TextDecoder('utf-8');

/** An .xlsx/.xlsm file opened as its constituent parts. */
export class OoxmlPackage {
  private readonly files: Map<string, Uint8Array>;
  private readonly order: string[];
  private readonly textCache = new Map<string, string>();

  private constructor(files: Map<string, Uint8Array>, order: string[]) {
    this.files = files;
    this.order = order;
  }

  static open(bytes: Uint8Array): OoxmlPackage {
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
      throw new Error(
        'This file is password-protected or in the old .xls format. Only .xlsx and .xlsm files are supported.',
      );
    }
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new Error('This file is not an Excel workbook (.xlsx or .xlsm).');
    }
    const raw = unzipSync(bytes);
    const files = new Map<string, Uint8Array>();
    const order: string[] = [];
    for (const [path, data] of Object.entries(raw)) {
      if (path.endsWith('/')) continue;
      files.set(path, data);
      order.push(path);
    }
    if (!files.has('[Content_Types].xml')) {
      throw new Error('This file is not a valid Office Open XML package.');
    }
    return new OoxmlPackage(files, order);
  }

  has(path: string): boolean {
    return this.files.has(path);
  }

  list(): string[] {
    return [...this.order];
  }

  text(path: string): string | undefined {
    const cached = this.textCache.get(path);
    if (cached !== undefined) return cached;
    const data = this.files.get(path);
    if (!data) return undefined;
    const s = decoder.decode(data);
    this.textCache.set(path, s);
    return s;
  }

  setText(path: string, text: string): void {
    if (!this.files.has(path)) this.order.push(path);
    this.files.set(path, strToU8(text));
    this.textCache.set(path, text);
  }

  remove(path: string): void {
    this.files.delete(path);
    this.textCache.delete(path);
    const i = this.order.indexOf(path);
    if (i !== -1) this.order.splice(i, 1);
  }

  /** Releases decoded text for a part that will not be read again. */
  forgetText(path: string): void {
    this.textCache.delete(path);
  }

  toBytes(): Uint8Array {
    const zippable: Zippable = {};
    for (const path of this.order) {
      const data = this.files.get(path);
      if (data) zippable[path] = [data, { level: 6 }];
    }
    return zipSync(zippable);
  }

  /** A copy whose part edits do not affect this package. */
  clone(): OoxmlPackage {
    const copy = new OoxmlPackage(new Map(this.files), [...this.order]);
    for (const [k, v] of this.textCache) copy.textCache.set(k, v);
    return copy;
  }

  rels(partPath: string): Relationship[] {
    const relsPath = relsPathFor(partPath);
    const xml = this.text(relsPath);
    if (!xml) return [];
    const root = parseXml(xml);
    return children(root, 'Relationship').map((r) => {
      const external = attr(r, 'TargetMode') === 'External';
      const target = attr(r, 'Target') ?? '';
      const type = attr(r, 'Type') ?? '';
      return {
        id: attr(r, 'Id') ?? '',
        type: type.slice(type.lastIndexOf('/') + 1),
        target: external ? target : resolveTarget(partPath, target),
        external,
      };
    });
  }
}

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const base = sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1);
  const parts: string[] = [];
  for (const seg of (base + target).split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}
