/**
 * Microsoft's Compound File Binary format (MS-CFB): a file system in a
 * file, which Office uses for .xls files and for encrypted .xlsx files.
 * Read whole, its streams by path ("\u0006DataSpaces/Version"); written as
 * version 4, with 4,096-byte sectors.
 */

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const NONE = 0xffffffff;
const MINI_CUTOFF = 4096;
const MINI_SECTOR = 64;

/** Whether bytes start as a compound file does. */
export function isCompoundFile(bytes: Uint8Array): boolean {
  return bytes.length >= 512 && SIGNATURE.every((b, i) => bytes[i] === b);
}

/** A compound file's streams, by their path from the root, storages' names joined by "/". */
export function readCompoundFile(bytes: Uint8Array): Map<string, Uint8Array> {
  if (!isCompoundFile(bytes)) throw new Error('Not a compound file');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (at: number) => view.getUint32(at, true);
  const sectorSize = 1 << view.getUint16(0x1e, true);
  // Sector n follows the header, which takes the first sector's room.
  const at = (n: number) => sectorSize + n * sectorSize;

  // The FAT's own sectors: 109 in the header, the rest in a chain of DIFAT sectors.
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) fatSectors.push(u32(0x4c + i * 4));
  let difat = u32(0x44);
  for (let guard = 0; difat !== END && difat !== FREE && guard < 1 << 20; guard++) {
    const base = at(difat);
    const per = sectorSize / 4 - 1;
    for (let i = 0; i < per; i++) fatSectors.push(u32(base + i * 4));
    difat = u32(base + per * 4);
  }
  const fat: number[] = [];
  for (const s of fatSectors.slice(0, u32(0x2c))) {
    if (s === FREE || at(s) + sectorSize > bytes.length) continue;
    for (let i = 0; i < sectorSize / 4; i++) fat.push(u32(at(s) + i * 4));
  }
  /** The sectors of a chain, in order. */
  const chain = (start: number, table: number[]) => {
    const out: number[] = [];
    for (let s = start; s !== END && s !== FREE && s < table.length && out.length <= table.length; s = table[s]) out.push(s);
    return out;
  };
  const readChain = (start: number) => {
    const sectors = chain(start, fat);
    const out = new Uint8Array(sectors.length * sectorSize);
    sectors.forEach((s, i) => out.set(bytes.subarray(at(s), Math.min(bytes.length, at(s) + sectorSize)), i * sectorSize));
    return out;
  };

  const dir = readChain(u32(0x30));
  const dirView = new DataView(dir.buffer);
  interface Entry {
    name: string;
    type: number;
    left: number;
    right: number;
    child: number;
    start: number;
    size: number;
  }
  const entries: Entry[] = [];
  for (let off = 0; off + 128 <= dir.length; off += 128) {
    const nameLength = dirView.getUint16(off + 0x40, true);
    let name = '';
    for (let i = 0; i < Math.max(0, nameLength / 2 - 1); i++) name += String.fromCharCode(dirView.getUint16(off + i * 2, true));
    entries.push({
      name,
      type: dir[off + 0x42],
      left: dirView.getUint32(off + 0x44, true),
      right: dirView.getUint32(off + 0x48, true),
      child: dirView.getUint32(off + 0x4c, true),
      start: dirView.getUint32(off + 0x74, true),
      size: dirView.getUint32(off + 0x78, true),
    });
  }
  const root = entries[0];
  const miniStream = root ? readChain(root.start) : new Uint8Array();
  const miniFat: number[] = [];
  const miniFatBytes = u32(0x3c) === END ? new Uint8Array() : readChain(u32(0x3c));
  for (let i = 0; i + 4 <= miniFatBytes.length; i += 4) miniFat.push(new DataView(miniFatBytes.buffer).getUint32(i, true));

  const out = new Map<string, Uint8Array>();
  const seen = new Set<number>();
  const walk = (index: number, prefix: string) => {
    if (index === NONE || index >= entries.length || seen.has(index)) return;
    seen.add(index);
    const e = entries[index];
    walk(e.left, prefix);
    walk(e.right, prefix);
    const path = prefix + e.name;
    if (e.type === 1) walk(e.child, `${path}/`);
    else if (e.type === 2) {
      if (e.size < MINI_CUTOFF) {
        const data = new Uint8Array(e.size);
        chain(e.start, miniFat).forEach((s, i) => data.set(miniStream.subarray(s * MINI_SECTOR, Math.min(e.size - i * MINI_SECTOR, MINI_SECTOR) + s * MINI_SECTOR), i * MINI_SECTOR));
        out.set(path, data);
      } else {
        out.set(path, readChain(e.start).subarray(0, e.size));
      }
    }
  };
  if (root) walk(root.child, '');
  return out;
}

/**
 * A compound file holding `streams`, by path: storages are made for the
 * folders the paths name. Version 4, with 4,096-byte sectors, so the FAT's
 * sectors fit in the header for files up to about 450 MB.
 */
export function writeCompoundFile(streams: Map<string, Uint8Array>): Uint8Array {
  const SECTOR = 4096;
  const PER = SECTOR / 4;

  interface Node {
    name: string;
    type: 1 | 2 | 5;
    data?: Uint8Array;
    children: Node[];
    index?: number;
    start?: number;
  }
  const root: Node = { name: 'Root Entry', type: 5, children: [] };
  for (const [path, data] of streams) {
    const parts = path.split('/');
    let at = root;
    for (const part of parts.slice(0, -1)) {
      let next = at.children.find((c) => c.name === part && c.type === 1);
      if (!next) at.children.push((next = { name: part, type: 1, children: [] }));
      at = next;
    }
    at.children.push({ name: parts[parts.length - 1], type: 2, data, children: [] });
  }
  // Entries in the order of the directory: the root, then each storage's children after it.
  const order: Node[] = [];
  const number = (n: Node) => {
    n.index = order.length;
    order.push(n);
    for (const c of n.children) number(c);
  };
  number(root);

  // Small streams go in the mini stream, in 64-byte sectors; the rest in sectors of their own.
  const miniFat: number[] = [];
  const miniParts: Uint8Array[] = [];
  for (const n of order) {
    if (n.type !== 2 || n.data!.length >= MINI_CUTOFF) continue;
    const count = Math.ceil(n.data!.length / MINI_SECTOR);
    n.start = count ? miniFat.length : END;
    for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? END : miniFat.length + 1);
    const padded = new Uint8Array(count * MINI_SECTOR);
    padded.set(n.data!);
    miniParts.push(padded);
  }
  const miniStream = concat(miniParts);
  const sectorsOf = (n: number) => Math.ceil(n / SECTOR);
  const dirSectors = sectorsOf(order.length * 128);
  const miniFatSectors = sectorsOf(miniFat.length * 4);
  const miniStreamSectors = sectorsOf(miniStream.length);
  const big = order.filter((n) => n.type === 2 && n.data!.length >= MINI_CUTOFF);
  const dataSectors = miniStreamSectors + big.reduce((s, n) => s + sectorsOf(n.data!.length), 0);
  // The FAT covers itself too.
  let fatSectors = 1;
  while (fatSectors * PER < fatSectors + dirSectors + miniFatSectors + dataSectors) fatSectors++;
  if (fatSectors > 109) throw new Error('Too large to write');
  const total = fatSectors + dirSectors + miniFatSectors + dataSectors;

  const fat = new Uint32Array(fatSectors * PER).fill(FREE);
  let next = 0;
  const run = (count: number, mark?: number) => {
    const start = next;
    for (let i = 0; i < count; i++, next++) fat[next] = mark ?? (i === count - 1 ? END : next + 1);
    return count ? start : END;
  };
  run(fatSectors, FAT_SECTOR);
  const dirStart = run(dirSectors);
  const miniFatStart = run(miniFatSectors);
  root.start = run(miniStreamSectors);
  for (const n of big) n.start = run(sectorsOf(n.data!.length));

  const out = new Uint8Array(SECTOR * (1 + total));
  const view = new DataView(out.buffer);
  out.set(SIGNATURE);
  view.setUint16(0x18, 0x3e, true);
  view.setUint16(0x1a, 4, true);
  view.setUint16(0x1c, 0xfffe, true);
  view.setUint16(0x1e, 12, true);
  view.setUint16(0x20, 6, true);
  view.setUint32(0x28, dirSectors, true);
  view.setUint32(0x2c, fatSectors, true);
  view.setUint32(0x30, dirStart, true);
  view.setUint32(0x38, MINI_CUTOFF, true);
  view.setUint32(0x3c, miniFatSectors ? miniFatStart : END, true);
  view.setUint32(0x40, miniFatSectors, true);
  view.setUint32(0x44, END, true);
  for (let i = 0; i < 109; i++) view.setUint32(0x4c + i * 4, i < fatSectors ? i : FREE, true);
  const sectorAt = (n: number) => SECTOR * (1 + n);

  for (let i = 0; i < fat.length; i++) view.setUint32(sectorAt(0) + i * 4, fat[i], true);
  // Siblings in the order MS-CFB sorts names: shorter first, then by upper-cased code units; chained to the right, all black.
  const sorted = (list: Node[]) =>
    [...list].sort((a, b) => a.name.length - b.name.length || (a.name.toUpperCase() < b.name.toUpperCase() ? -1 : a.name.toUpperCase() > b.name.toUpperCase() ? 1 : 0));
  const right = new Map<Node, number>();
  const child = new Map<Node, number>();
  for (const n of order) {
    const kids = sorted(n.children);
    if (kids.length) child.set(n, kids[0].index!);
    kids.forEach((k, i) => i + 1 < kids.length && right.set(k, kids[i + 1].index!));
  }
  const dirBase = sectorAt(dirStart);
  for (const n of order) {
    const off = dirBase + n.index! * 128;
    for (let i = 0; i < n.name.length; i++) view.setUint16(off + i * 2, n.name.charCodeAt(i), true);
    view.setUint16(off + 0x40, (n.name.length + 1) * 2, true);
    out[off + 0x42] = n.type;
    out[off + 0x43] = 1;
    view.setUint32(off + 0x44, NONE, true);
    view.setUint32(off + 0x48, right.get(n) ?? NONE, true);
    view.setUint32(off + 0x4c, child.get(n) ?? NONE, true);
    if (n.type !== 1) {
      view.setUint32(off + 0x74, n.start ?? END, true);
      view.setUint32(off + 0x78, n.type === 5 ? miniStream.length : n.data!.length, true);
    }
  }
  // Unused directory entries are empty, pointing nowhere.
  for (let off = dirBase + order.length * 128; off < dirBase + dirSectors * SECTOR; off += 128) {
    view.setUint32(off + 0x44, NONE, true);
    view.setUint32(off + 0x48, NONE, true);
    view.setUint32(off + 0x4c, NONE, true);
  }
  for (let i = 0; i < miniFatSectors * PER; i++) view.setUint32(sectorAt(miniFatStart) + i * 4, i < miniFat.length ? miniFat[i] : FREE, true);
  if (miniStream.length) out.set(miniStream, sectorAt(root.start));
  for (const n of big) out.set(n.data!, sectorAt(n.start!));
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

