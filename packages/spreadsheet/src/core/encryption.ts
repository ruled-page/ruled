/**
 * Password-protected workbooks, as Office encrypts them (MS-OFFCRYPTO): an
 * .xlsx package encrypted inside a compound file. Agile encryption, which
 * Office has written since 2010, is read and written; Standard encryption,
 * Office 2007's, is read. The hashes and the cipher are worked out here, in
 * plain code, so that opening a workbook stays synchronous and works on
 * any page, secure or not.
 */
import { isCompoundFile, readCompoundFile, writeCompoundFile } from './cfb';
import { attr, child, parseXml } from './xml';

/** Why a protected workbook can't be opened: no or a wrong password, or a kind of protection the editor can't open. */
export class EncryptionError extends Error {
  constructor(readonly reason: 'password' | 'rights' | 'unsupported') {
    super(reason);
  }
}

// ---------------------------------------------------------------------------
// Hashes

type Hash = (data: Uint8Array) => Uint8Array;

let k256: Uint32Array | undefined;
let k512: Int32Array | undefined;
/** The first 64 primes', or 80's, cube roots' fractional bits, and square roots' for starting values: worked out once rather than written out. */
function primes(n: number): bigint[] {
  const out: bigint[] = [];
  for (let p = 2; out.length < n; p++) if (out.every((q) => BigInt(p) % q !== 0n)) out.push(BigInt(p));
  return out;
}
function iroot(x: bigint, k: 2n | 3n): bigint {
  let r = BigInt(Math.floor(Number(x) ** (1 / Number(k)))) || 1n;
  for (;;) {
    const next = ((k - 1n) * r + x / r ** (k - 1n)) / k;
    if (next >= r && r ** k <= x && (r + 1n) ** k > x) return r;
    r = next;
  }
}
const frac = (p: bigint, k: 2n | 3n, bits: bigint) => iroot(p << (k * bits), k) & ((1n << bits) - 1n);

function sha256(data: Uint8Array): Uint8Array {
  if (!k256) k256 = Uint32Array.from(primes(64), (p) => Number(frac(p, 3n, 32n)));
  const k = k256;
  const h = Uint32Array.from(primes(8), (p) => Number(frac(p, 2n, 32n)));
  const w = new Uint32Array(64);
  const rotr = (x: number, n: number) => (x >>> n) | (x << (32 - n));
  for (const block of blocks(data, 64, 8)) {
    for (let i = 0; i < 16; i++) w[i] = block.getUint32(i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, hh] = h;
    for (let i = 0; i < 64; i++) {
      const t1 = (hh + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + k[i] + w[i]) >>> 0;
      const t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
      [hh, g, f, e, d, c, b, a] = [g, f, e, (d + t1) >>> 0, c, b, a, (t1 + t2) >>> 0];
    }
    [a, b, c, d, e, f, g, hh].forEach((v, i) => (h[i] = (h[i] + v) >>> 0));
  }
  return words(h, 8);
}

function sha1(data: Uint8Array): Uint8Array {
  const h = Uint32Array.of(0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0);
  const w = new Uint32Array(80);
  const rotl = (x: number, n: number) => (x << n) | (x >>> (32 - n));
  for (const block of blocks(data, 64, 8)) {
    for (let i = 0; i < 16; i++) w[i] = block.getUint32(i * 4);
    for (let i = 16; i < 80; i++) w[i] = rotl(w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16], 1);
    let [a, b, c, d, e] = h;
    for (let i = 0; i < 80; i++) {
      const [f, k] = i < 20 ? [(b & c) | (~b & d), 0x5a827999] : i < 40 ? [b ^ c ^ d, 0x6ed9eba1] : i < 60 ? [(b & c) | (b & d) | (c & d), 0x8f1bbcdc] : [b ^ c ^ d, 0xca62c1d6];
      const t = (rotl(a, 5) + f + e + k + w[i]) >>> 0;
      [e, d, c, b, a] = [d, c, rotl(b, 30) >>> 0, a, t];
    }
    [a, b, c, d, e].forEach((v, i) => (h[i] = (h[i] + v) >>> 0));
  }
  return words(h, 5);
}

/** SHA-512, or SHA-384 (other starting values, cut short), with 64-bit words as pairs of 32-bit halves. */
function sha512(data: Uint8Array, bits: 384 | 512 = 512): Uint8Array {
  if (!k512) {
    k512 = new Int32Array(160);
    primes(80).forEach((p, i) => {
      const v = frac(p, 3n, 64n);
      k512![i * 2] = Number(v >> 32n) | 0;
      k512![i * 2 + 1] = Number(v & 0xffffffffn) | 0;
    });
  }
  const k = k512;
  const start = bits === 512 ? primes(8) : primes(16).slice(8);
  const h = new Int32Array(16);
  start.forEach((p, i) => {
    const v = frac(p, 2n, 64n);
    h[i * 2] = Number(v >> 32n) | 0;
    h[i * 2 + 1] = Number(v & 0xffffffffn) | 0;
  });
  const wh = new Int32Array(80);
  const wl = new Int32Array(80);
  const v = new Int32Array(16);
  for (const block of blocks(data, 128, 16)) {
    for (let i = 0; i < 16; i++) {
      wh[i] = block.getInt32(i * 8);
      wl[i] = block.getInt32(i * 8 + 4);
    }
    for (let i = 16; i < 80; i++) {
      // σ0 = rotr 1 ^ rotr 8 ^ shr 7; σ1 = rotr 19 ^ rotr 61 ^ shr 6.
      let xh = wh[i - 15];
      let xl = wl[i - 15];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = wh[i - 2];
      xl = wl[i - 2];
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      let lo = (wl[i - 16] >>> 0) + (s0l >>> 0) + (wl[i - 7] >>> 0) + (s1l >>> 0);
      wh[i] = (wh[i - 16] + s0h + wh[i - 7] + s1h + Math.floor(lo / 0x100000000)) | 0;
      wl[i] = lo | 0;
    }
    let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = h;
    for (let i = 0; i < 80; i++) {
      // Σ1(e) = rotr 14 ^ rotr 18 ^ rotr 41; Σ0(a) = rotr 28 ^ rotr 34 ^ rotr 39.
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      let lo = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (k[i * 2 + 1] >>> 0) + (wl[i] >>> 0);
      const t1h = (hh + S1h + chh + k[i * 2] + wh[i] + Math.floor(lo / 0x100000000)) | 0;
      const t1l = lo | 0;
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      lo = (S0l >>> 0) + (((al & bl) ^ (al & cl) ^ (bl & cl)) >>> 0);
      const t2h = (S0h + ((ah & bh) ^ (ah & ch) ^ (bh & ch)) + Math.floor(lo / 0x100000000)) | 0;
      const t2l = lo | 0;
      hh = gh;
      hl = gl;
      gh = fh;
      gl = fl;
      fh = eh;
      fl = el;
      lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x100000000)) | 0;
      el = lo | 0;
      dh = ch;
      dl = cl;
      ch = bh;
      cl = bl;
      bh = ah;
      bl = al;
      lo = (t1l >>> 0) + (t2l >>> 0);
      ah = (t1h + t2h + Math.floor(lo / 0x100000000)) | 0;
      al = lo | 0;
    }
    v.set([ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl]);
    for (let i = 0; i < 16; i += 2) {
      const lo = (h[i + 1] >>> 0) + (v[i + 1] >>> 0);
      h[i] = (h[i] + v[i] + Math.floor(lo / 0x100000000)) | 0;
      h[i + 1] = lo | 0;
    }
  }
  return words(h, bits / 32);
}

/** A message padded as the SHA family pads it, in blocks of `size` bytes, with its length in the last `lengthBytes`. */
function* blocks(data: Uint8Array, size: number, lengthBytes: number): Generator<DataView> {
  const total = Math.ceil((data.length + 1 + lengthBytes) / size) * size;
  const tail = new Uint8Array(total - Math.floor(data.length / size) * size);
  tail.set(data.subarray(Math.floor(data.length / size) * size));
  tail[data.length % size] = 0x80;
  const tv = new DataView(tail.buffer);
  const bitLength = data.length * 8;
  tv.setUint32(tail.length - 8, Math.floor(bitLength / 0x100000000));
  tv.setUint32(tail.length - 4, bitLength >>> 0);
  const whole = Math.floor(data.length / size);
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  for (let i = 0; i < whole; i++) yield new DataView(dv.buffer, dv.byteOffset + i * size, size);
  for (let i = 0; i < tail.length; i += size) yield new DataView(tail.buffer, i, size);
}

function words(h: Uint32Array | Int32Array, count: number): Uint8Array {
  const out = new Uint8Array(count * 4);
  const view = new DataView(out.buffer);
  for (let i = 0; i < count; i++) view.setUint32(i * 4, h[i] >>> 0);
  return out;
}

const HASHES: Record<string, { hash: Hash; block: number }> = {
  SHA1: { hash: sha1, block: 64 },
  'SHA-1': { hash: sha1, block: 64 },
  SHA256: { hash: sha256, block: 64 },
  SHA384: { hash: (d) => sha512(d, 384), block: 128 },
  SHA512: { hash: (d) => sha512(d), block: 128 },
};

function hmac(alg: { hash: Hash; block: number }, key: Uint8Array, data: Uint8Array): Uint8Array {
  const k = new Uint8Array(alg.block);
  k.set(key.length > alg.block ? alg.hash(key) : key);
  const inner = alg.hash(join(k.map((b) => b ^ 0x36), data));
  return alg.hash(join(k.map((b) => b ^ 0x5c), inner));
}

// ---------------------------------------------------------------------------
// AES

let tables: { s: Uint8Array; si: Uint8Array; te: Uint32Array[]; td: Uint32Array[] } | undefined;
/** The S-box and the round tables, worked out from the field's arithmetic. */
function aesTables() {
  if (tables) return tables;
  const s = new Uint8Array(256);
  const si = new Uint8Array(256);
  const xt = (a: number) => ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 0xff;
  const mul = (a: number, b: number) => {
    let r = 0;
    for (; b; b >>= 1, a = xt(a)) if (b & 1) r ^= a;
    return r;
  };
  for (let i = 0; i < 256; i++) {
    let inv = 0;
    if (i) for (inv = 1; mul(i, inv) !== 1; inv++);
    let x = inv;
    let r = inv;
    for (let j = 0; j < 4; j++) {
      x = ((x << 1) | (x >> 7)) & 0xff;
      r ^= x;
    }
    s[i] = r ^ 0x63;
    si[s[i]] = i;
  }
  const te = [0, 1, 2, 3].map(() => new Uint32Array(256));
  const td = [0, 1, 2, 3].map(() => new Uint32Array(256));
  for (let i = 0; i < 256; i++) {
    const a = s[i];
    const e = ((mul(a, 2) << 24) | (a << 16) | (a << 8) | mul(a, 3)) >>> 0;
    const b = si[i];
    const d = ((mul(b, 14) << 24) | (mul(b, 9) << 16) | (mul(b, 13) << 8) | mul(b, 11)) >>> 0;
    for (let t = 0; t < 4; t++) {
      te[t][i] = ((e >>> (8 * t)) | (e << (32 - 8 * t))) >>> 0;
      td[t][i] = ((d >>> (8 * t)) | (d << (32 - 8 * t))) >>> 0;
    }
  }
  return (tables = { s, si, te, td });
}

/** AES with a 128, 192 or 256-bit key: its round keys for encrypting and for decrypting. */
function aesKey(key: Uint8Array): { rounds: number; ek: Uint32Array; dk: Uint32Array } {
  const { s, td } = aesTables();
  const nk = key.length / 4;
  if (![4, 6, 8].includes(nk)) throw new EncryptionError('unsupported');
  const rounds = nk + 6;
  const ek = new Uint32Array(4 * (rounds + 1));
  const kv = new DataView(key.buffer, key.byteOffset, key.byteLength);
  for (let i = 0; i < nk; i++) ek[i] = kv.getUint32(i * 4);
  const sub = (w: number) => ((s[w >>> 24] << 24) | (s[(w >>> 16) & 255] << 16) | (s[(w >>> 8) & 255] << 8) | s[w & 255]) >>> 0;
  for (let i = nk, rcon = 1; i < ek.length; i++) {
    let t = ek[i - 1];
    if (i % nk === 0) {
      t = (sub(((t << 8) | (t >>> 24)) >>> 0) ^ (rcon << 24)) >>> 0;
      rcon = ((rcon << 1) ^ (rcon & 0x80 ? 0x1b : 0)) & 0xff;
    } else if (nk > 6 && i % nk === 4) t = sub(t);
    ek[i] = (ek[i - nk] ^ t) >>> 0;
  }
  // The equivalent inverse cipher's keys: the rounds' in reverse, mixed back for the middle rounds.
  const dk = new Uint32Array(ek.length);
  for (let r = 0; r <= rounds; r++) {
    for (let j = 0; j < 4; j++) {
      const w = ek[(rounds - r) * 4 + j];
      dk[r * 4 + j] = r === 0 || r === rounds ? w : (td[0][s[w >>> 24]] ^ td[1][s[(w >>> 16) & 255]] ^ td[2][s[(w >>> 8) & 255]] ^ td[3][s[w & 255]]) >>> 0;
    }
  }
  return { rounds, ek, dk };
}

function aesBlock(k: ReturnType<typeof aesKey>, data: Uint8Array, at: number, out: Uint8Array, decrypt: boolean): void {
  const { s, si, te, td } = aesTables();
  const rk = decrypt ? k.dk : k.ek;
  const [t0, t1, t2, t3] = decrypt ? td : te;
  const box = decrypt ? si : s;
  const dv = new DataView(data.buffer, data.byteOffset + at, 16);
  let a = dv.getUint32(0) ^ rk[0];
  let b = dv.getUint32(4) ^ rk[1];
  let c = dv.getUint32(8) ^ rk[2];
  let d = dv.getUint32(12) ^ rk[3];
  // Encrypting takes each column's bytes from the columns to its right, decrypting from those to its left.
  const step = (w: number, x: number, y: number, z: number, n: number) => t0[w >>> 24] ^ t1[(x >>> 16) & 255] ^ t2[(y >>> 8) & 255] ^ t3[z & 255] ^ rk[n];
  for (let r = 1; r < k.rounds; r++) {
    const n = r * 4;
    [a, b, c, d] = decrypt
      ? [step(a, d, c, b, n), step(b, a, d, c, n + 1), step(c, b, a, d, n + 2), step(d, c, b, a, n + 3)]
      : [step(a, b, c, d, n), step(b, c, d, a, n + 1), step(c, d, a, b, n + 2), step(d, a, b, c, n + 3)];
  }
  const n = k.rounds * 4;
  const last = (w: number, x: number, y: number, z: number, i: number) => ((box[w >>> 24] << 24) | (box[(x >>> 16) & 255] << 16) | (box[(y >>> 8) & 255] << 8) | box[z & 255]) ^ rk[n + i];
  const ov = new DataView(out.buffer, out.byteOffset + at, 16);
  const result = decrypt ? [last(a, d, c, b, 0), last(b, a, d, c, 1), last(c, b, a, d, 2), last(d, c, b, a, 3)] : [last(a, b, c, d, 0), last(b, c, d, a, 1), last(c, d, a, b, 2), last(d, a, b, c, 3)];
  result.forEach((w, i) => ov.setUint32(i * 4, w >>> 0));
}

/** AES in CBC mode, or ECB without an IV, over data a whole number of blocks long; no padding is added or taken off. */
function aes(key: Uint8Array, data: Uint8Array, decrypt: boolean, iv?: Uint8Array): Uint8Array {
  const k = aesKey(key);
  const out = new Uint8Array(data.length - (data.length % 16));
  let prev = iv ? Uint8Array.from(iv.subarray(0, 16)) : undefined;
  const block = new Uint8Array(16);
  for (let at = 0; at < out.length; at += 16) {
    if (decrypt) {
      aesBlock(k, data, at, out, true);
      if (prev) for (let i = 0; i < 16; i++) out[at + i] ^= prev[i];
      prev = prev && data.slice(at, at + 16);
    } else {
      block.set(data.subarray(at, at + 16));
      if (prev) for (let i = 0; i < 16; i++) block[i] ^= prev[i];
      aesBlock(k, block, 0, out.subarray(at, at + 16), false);
      prev = prev && out.slice(at, at + 16);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Office's encryption

/** What kind of file a compound file is, as far as opening it goes. */
export function protectionOf(bytes: Uint8Array): 'none' | 'password' | 'rights' | 'other' {
  if (!isCompoundFile(bytes)) return 'none';
  let streams: Map<string, Uint8Array>;
  try {
    streams = readCompoundFile(bytes);
  } catch {
    return 'other';
  }
  // Rights management (sensitivity labels, IRM) encrypts with keys from Microsoft's service.
  if ([...streams.keys()].some((p) => /DRMEncrypted/i.test(p))) return 'rights';
  return streams.has('EncryptionInfo') && streams.has('EncryptedPackage') ? 'password' : 'other';
}

/** The workbook package inside a password-protected file; EncryptionError('password') for a wrong password. */
export function decryptWorkbook(file: Uint8Array, password: string): Uint8Array {
  const streams = readCompoundFile(file);
  const info = streams.get('EncryptionInfo');
  const pkg = streams.get('EncryptedPackage');
  if (!info || !pkg || info.length < 8 || pkg.length < 8) throw new EncryptionError('unsupported');
  const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const [major, minor] = [view.getUint16(0, true), view.getUint16(2, true)];
  if (major === 4 && minor === 4) return agileDecrypt(info, pkg, password);
  if ((major === 3 || major === 4) && minor === 2) return standardDecrypt(info, pkg, password);
  throw new EncryptionError('unsupported');
}

const BLOCK_VERIFIER_INPUT = Uint8Array.of(0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79);
const BLOCK_VERIFIER_HASH = Uint8Array.of(0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e);
const BLOCK_KEY = Uint8Array.of(0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6);
const BLOCK_HMAC_KEY = Uint8Array.of(0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6);
const BLOCK_HMAC_VALUE = Uint8Array.of(0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33);
const SEGMENT = 4096;

/** The password's hash, salted and hashed again `spin` times. */
function passwordHash(hash: Hash, salt: Uint8Array, password: string, spin: number): Uint8Array {
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    pw[i * 2] = password.charCodeAt(i) & 255;
    pw[i * 2 + 1] = password.charCodeAt(i) >> 8;
  }
  let h = hash(join(salt, pw));
  const buf = new Uint8Array(4 + h.length);
  const bv = new DataView(buf.buffer);
  for (let i = 0; i < spin; i++) {
    bv.setUint32(0, i, true);
    buf.set(h, 4);
    h = hash(buf);
  }
  return h;
}

/** Bytes cut or padded (with `pad`) to a length. */
function fit(bytes: Uint8Array, length: number, pad = 0x36): Uint8Array {
  const out = new Uint8Array(length).fill(pad);
  out.set(bytes.subarray(0, length));
  return out;
}

/** The setting of an Agile key: its cipher, hash, salt and sizes. */
interface AgileParams {
  hash: { hash: Hash; block: number };
  keyBytes: number;
  blockSize: number;
  salt: Uint8Array;
}

function agileParams(node: ReturnType<typeof parseXml> | undefined): AgileParams {
  const hash = HASHES[(attr(node, 'hashAlgorithm') ?? '').toUpperCase()];
  if (!node || !hash || attr(node, 'cipherAlgorithm') !== 'AES' || attr(node, 'cipherChaining') !== 'ChainingModeCBC') throw new EncryptionError('unsupported');
  return { hash, keyBytes: Number(attr(node, 'keyBits')) / 8, blockSize: Number(attr(node, 'blockSize') ?? 16), salt: base64(attr(node, 'saltValue') ?? '') };
}

/** The package from an Agile-encrypted file: the password unlocks the package's key, which decrypts it in 4,096-byte segments. */
function agileDecrypt(info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
  const xml = parseXml(new TextDecoder().decode(info.subarray(8)));
  const keyData = agileParams(child(xml, 'keyData'));
  const encryptor = [...(child(child(xml, 'keyEncryptors'), 'keyEncryptor')?.children ?? [])].find((c) => c.name === 'encryptedKey');
  if (!encryptor) throw new EncryptionError('unsupported');
  const p = agileParams(encryptor);
  const h = passwordHash(p.hash.hash, p.salt, password, Number(attr(encryptor, 'spinCount') ?? 100_000));
  const keyFor = (block: Uint8Array) => fit(p.hash.hash(join(h, block)), p.keyBytes);
  const iv = fit(p.salt, p.blockSize);
  const input = aes(keyFor(BLOCK_VERIFIER_INPUT), base64(attr(encryptor, 'encryptedVerifierHashInput') ?? ''), true, iv).subarray(0, p.salt.length);
  const expected = aes(keyFor(BLOCK_VERIFIER_HASH), base64(attr(encryptor, 'encryptedVerifierHashValue') ?? ''), true, iv);
  const got = p.hash.hash(input);
  if (!got.every((b, i) => b === expected[i])) throw new EncryptionError('password');
  const key = aes(keyFor(BLOCK_KEY), base64(attr(encryptor, 'encryptedKeyValue') ?? ''), true, iv).subarray(0, keyData.keyBytes);
  return segments(pkg, key, keyData, true);
}

/** A package encrypted or decrypted segment by segment, each with its own IV, after its 8-byte size. */
function segments(pkg: Uint8Array, key: Uint8Array, keyData: AgileParams, decrypt: boolean): Uint8Array {
  const pv = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
  const size = decrypt ? pv.getUint32(0, true) + pv.getUint32(4, true) * 0x100000000 : pkg.length;
  const body = decrypt ? pkg.subarray(8) : pkg;
  const parts: Uint8Array[] = [];
  const index = new Uint8Array(4);
  for (let i = 0, at = 0; at < body.length; i++, at += SEGMENT) {
    new DataView(index.buffer).setUint32(0, i, true);
    const iv = fit(keyData.hash.hash(join(keyData.salt, index)), keyData.blockSize);
    const chunk = body.subarray(at, at + SEGMENT);
    parts.push(aes(key, decrypt ? chunk : fit(chunk, Math.ceil(chunk.length / 16) * 16, 0), decrypt, iv));
  }
  if (!decrypt) {
    const head = new Uint8Array(8);
    new DataView(head.buffer).setUint32(0, size % 0x100000000, true);
    new DataView(head.buffer).setUint32(4, Math.floor(size / 0x100000000), true);
    return join(head, ...parts);
  }
  return join(...parts).subarray(0, size);
}

/** The package from a Standard-encrypted file (Office 2007): AES in ECB mode, with a key from SHA-1. */
function standardDecrypt(info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
  const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const flags = view.getUint32(4, true);
  const headerSize = view.getUint32(8, true);
  const algorithm = view.getUint32(12 + 8, true);
  const keyBits = view.getUint32(12 + 16, true);
  // Only AES (fAES), with SHA-1; RC4 is refused.
  if (!(flags & 0x20) || ![0x660e, 0x660f, 0x6610].includes(algorithm)) throw new EncryptionError('unsupported');
  const v = 12 + headerSize;
  const saltSize = view.getUint32(v, true);
  const salt = info.subarray(v + 4, v + 4 + saltSize);
  const encryptedVerifier = info.subarray(v + 4 + saltSize, v + 20 + saltSize);
  const hashSize = view.getUint32(v + 20 + saltSize, true);
  const encryptedHash = info.subarray(v + 24 + saltSize, v + 24 + saltSize + 32);
  let h = passwordHash(sha1, salt, password, 50_000);
  h = sha1(join(h, new Uint8Array(4)));
  const x = (fill: number) => sha1(new Uint8Array(64).fill(fill).map((b, i) => b ^ (h[i] ?? 0)));
  const key = join(x(0x36), x(0x5c)).subarray(0, keyBits / 8);
  const verifier = aes(key, encryptedVerifier, true);
  const hash = aes(key, encryptedHash, true).subarray(0, hashSize);
  if (!sha1(verifier).every((b, i) => b === hash[i])) throw new EncryptionError('password');
  const pv = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
  const size = pv.getUint32(0, true) + pv.getUint32(4, true) * 0x100000000;
  const body = pkg.subarray(8);
  return aes(key, body.subarray(0, body.length - (body.length % 16)), true).subarray(0, size);
}

/** Random bytes, from the platform's generator. */
const random = (n: number) => crypto.getRandomValues(new Uint8Array(n));

/**
 * A workbook package encrypted with a password as Office does it: Agile
 * encryption, AES-256 and SHA-512 with 100,000 rounds, with the HMAC that
 * proves it unchanged, in a compound file with the data spaces Office's
 * own files have.
 */
export function encryptWorkbook(zip: Uint8Array, password: string, spinCount = 100_000): Uint8Array {
  const hash = HASHES.SHA512;
  const keyData: AgileParams = { hash, keyBytes: 32, blockSize: 16, salt: random(16) };
  const keySalt = random(16);
  const key = random(32);
  const h = passwordHash(hash.hash, keySalt, password, spinCount);
  const keyFor = (block: Uint8Array) => fit(hash.hash(join(h, block)), 32);
  const verifier = random(16);
  const encryptedInput = aes(keyFor(BLOCK_VERIFIER_INPUT), verifier, false, keySalt);
  const encryptedHash = aes(keyFor(BLOCK_VERIFIER_HASH), hash.hash(verifier), false, keySalt);
  const encryptedKey = aes(keyFor(BLOCK_KEY), key, false, keySalt);
  const pkg = segments(zip, key, keyData, false);
  const hmacKey = random(64);
  const integrityIv = (block: Uint8Array) => fit(hash.hash(join(keyData.salt, block)), 16);
  const encryptedHmacKey = aes(key, hmacKey, false, integrityIv(BLOCK_HMAC_KEY));
  const encryptedHmacValue = aes(key, hmac(hash, hmacKey, pkg), false, integrityIv(BLOCK_HMAC_VALUE));
  const params = (salt: Uint8Array) =>
    `saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${toBase64(salt)}"`;
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
    `<keyData ${params(keyData.salt)}/>` +
    `<dataIntegrity encryptedHmacKey="${toBase64(encryptedHmacKey)}" encryptedHmacValue="${toBase64(encryptedHmacValue)}"/>` +
    '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    `<p:encryptedKey spinCount="${spinCount}" ${params(keySalt)} encryptedVerifierHashInput="${toBase64(encryptedInput)}" encryptedVerifierHashValue="${toBase64(encryptedHash)}" encryptedKeyValue="${toBase64(encryptedKey)}"/>` +
    '</keyEncryptor></keyEncryptors></encryption>';
  const info = join(Uint8Array.of(4, 0, 4, 0, 0x40, 0, 0, 0), new TextEncoder().encode(xml));
  return writeCompoundFile(
    new Map([
      ['\u0006DataSpaces/Version', dataSpaces.version()],
      ['\u0006DataSpaces/DataSpaceMap', dataSpaces.map()],
      ['\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', dataSpaces.info()],
      ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', dataSpaces.transform()],
      ['EncryptionInfo', info],
      ['EncryptedPackage', pkg],
    ]),
  );
}

/** The data spaces streams of an encrypted Office file, which say the package is encrypted with Office's own transform (MS-OFFCRYPTO 2.1). */
const dataSpaces = {
  version: () => join(text('Microsoft.Container.DataSpaces'), u32(1), u32(1), u32(1)),
  map: () => {
    const entry = join(u32(1), u32(0), text('EncryptedPackage'), text('StrongEncryptionDataSpace'));
    return join(u32(8), u32(1), u32(entry.length + 4), entry);
  },
  info: () => join(u32(8), u32(1), text('StrongEncryptionTransform')),
  transform: () => {
    const id = text('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
    return join(u32(8 + id.length), u32(1), id, text('Microsoft.Container.EncryptionTransform'), u32(1), u32(1), u32(1), u32(0), u32(0), u32(0), u32(4));
  },
};

/** A string as MS-OFFCRYPTO's length-prefixed UTF-16, padded to four bytes. */
function text(s: string): Uint8Array {
  const out = new Uint8Array(4 + Math.ceil((s.length * 2) / 4) * 4);
  const view = new DataView(out.buffer);
  view.setUint32(0, s.length * 2, true);
  for (let i = 0; i < s.length; i++) view.setUint16(4 + i * 2, s.charCodeAt(i), true);
  return out;
}

const u32 = (n: number) => {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, n, true);
  return out;
};

function join(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function base64(s: string): Uint8Array {
  const bin = atob(s.replace(/\s+/g, ''));
  return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}

function toBase64(bytes: Uint8Array): string {
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
}

/** For tests: the hashes and cipher, checked against other implementations. */
export const primitives = { sha1, sha256, sha384: (d: Uint8Array) => sha512(d, 384), sha512: (d: Uint8Array) => sha512(d), hmac512: (k: Uint8Array, d: Uint8Array) => hmac(HASHES.SHA512, k, d), aes };
