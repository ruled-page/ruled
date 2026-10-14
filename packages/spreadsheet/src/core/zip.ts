import { strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { msg, TextError } from '../i18n/message';
import { attr, checkTerminated, children, parseXml } from './xml';

export interface Relationship {
  id: string;
  /** Last segment of the relationship type URI, e.g. "worksheet", "chart". */
  type: string;
  /** Resolved package path (no leading slash), or the raw target when external. */
  target: string;
  external: boolean;
}

const decoder = new TextDecoder('utf-8');


/** Limits applied before a package is decompressed, so a hostile file can't exhaust memory. */
export interface PackageLimits {
  /** Largest file accepted. Default 50 MB. */
  maxFileBytes?: number;
  /** Largest total size once unzipped. Default 500 MB. */
  maxUnzippedBytes?: number;
  /** Highest unzipped:zipped ratio for a large part, the signature of a zip bomb. Default 200. */
  maxRatio?: number;
  /** Most parts in the package. Default 20,000. */
  maxParts?: number;
}

/** Sized for a browser tab, phones included: larger workbooks can be allowed through OpenOptions.limits. */
export const DEFAULT_LIMITS: Required<PackageLimits> = {
  maxFileBytes: 50 * 1024 * 1024,
  maxUnzippedBytes: 500 * 1024 * 1024,
  maxRatio: 200,
  maxParts: 20_000,
};

/** Bytes in whole megabytes, for messages that say "50 MB". */
const mb = (n: number) => Math.round(n / 1024 / 1024);

/** An .xlsx/.xlsm file opened as its constituent parts. */
export class OoxmlPackage {
  private readonly files: Map<string, Uint8Array>;
  private readonly order: string[];
  private readonly textCache = new Map<string, string>();
  /** Parts already checked for tags that never end. */
  private readonly checked = new Set<string>();

  private constructor(files: Map<string, Uint8Array>, order: string[]) {
    this.files = files;
    this.order = order;
  }

  static open(bytes: Uint8Array, limits: PackageLimits = {}): OoxmlPackage {
    const lim = { ...DEFAULT_LIMITS, ...limits };
    if (bytes.length > lim.maxFileBytes) {
      throw new TextError(msg('core.open.tooBig', { size: mb(bytes.length), max: mb(lim.maxFileBytes) }));
    }
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
      throw new TextError(msg('core.open.oldFormat'));
    }
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new TextError(msg('core.open.notWorkbook'));
    }
    // Sizes come from the zip directory, so they are checked before anything is inflated.
    let total = 0;
    let parts = 0;
    const raw = unzipSync(bytes, {
      filter: (f) => {
        parts++;
        total += f.originalSize;
        if (parts > lim.maxParts) throw new TextError(msg('core.open.tooManyParts', { max: lim.maxParts }));
        if (total > lim.maxUnzippedBytes) throw new TextError(msg('core.open.unpacksTooBig', { max: mb(lim.maxUnzippedBytes) }));
        if (f.originalSize > 10 * 1024 * 1024 && f.originalSize / Math.max(1, f.size) > lim.maxRatio) {
          throw new TextError(msg('core.open.zipBomb', { part: f.name, ratio: Math.round(f.originalSize / Math.max(1, f.size)) }));
        }
        return true;
      },
    });
    const files = new Map<string, Uint8Array>();
    const order: string[] = [];
    for (const [path, data] of Object.entries(raw)) {
      if (path.endsWith('/')) continue;
      files.set(path, data);
      order.push(path);
    }
    if (!files.has('[Content_Types].xml')) {
      throw new TextError(msg('core.open.notPackage'));
    }
    return new OoxmlPackage(files, order);
  }

  /** A package built in memory from the text of its parts, in the order given. */
  static fromParts(parts: Record<string, string>): OoxmlPackage {
    const files = new Map<string, Uint8Array>();
    for (const [path, text] of Object.entries(parts)) files.set(path, strToU8(text));
    return new OoxmlPackage(files, Object.keys(parts));
  }

  /** Decodes a part, checking once that an XML part isn't cut off mid-tag. */
  private decode(path: string, data: Uint8Array): string {
    const s = decoder.decode(data);
    if (!this.checked.has(path) && /\.(xml|rels|vml)$/i.test(path)) {
      checkTerminated(s);
      this.checked.add(path);
    }
    return s;
  }

  has(path: string): boolean {
    return this.files.has(path);
  }

  /** A part's bytes as stored, such as a picture's. */
  bytes(path: string): Uint8Array | undefined {
    return this.files.get(path);
  }

  list(): string[] {
    return [...this.order];
  }

  text(path: string): string | undefined {
    const cached = this.textCache.get(path);
    if (cached !== undefined) return cached;
    const data = this.files.get(path);
    if (!data) return undefined;
    const s = this.decode(path, data);
    this.textCache.set(path, s);
    return s;
  }

  setText(path: string, text: string): void {
    if (!this.files.has(path)) this.order.push(path);
    this.files.set(path, strToU8(text));
    this.textCache.set(path, text);
    this.checked.add(path);
  }

  remove(path: string): void {
    this.files.delete(path);
    this.textCache.delete(path);
    const i = this.order.indexOf(path);
    if (i !== -1) this.order.splice(i, 1);
  }

  /** A part's text, decoded without keeping it, for large parts read rarely. */
  textOnce(path: string): string | undefined {
    const cached = this.textCache.get(path);
    if (cached !== undefined) return cached;
    const data = this.files.get(path);
    return data ? this.decode(path, data) : undefined;
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
