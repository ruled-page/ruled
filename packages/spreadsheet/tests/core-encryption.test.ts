import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readCompoundFile, writeCompoundFile } from '../src/core/cfb';
import { decryptWorkbook, encryptWorkbook, EncryptionError, primitives, protectionOf } from '../src/core/encryption';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import { loadedValue } from './helpers';

const corpus = (name: string) => new Uint8Array(readFileSync(new URL(`./corpus/${name}`, import.meta.url)));
/** tests/corpus/outline-password.xlsx: outline.xlsx encrypted by msoffcrypto-tool with this password. */
const PASSWORD = 'Ruled-2026';

/**
 * A workbook encrypted as Office 2007 did, written here from Microsoft's
 * description of Standard encryption with Node's own SHA-1 and AES, so
 * decrypting it checks the editor's code against another implementation.
 * `algorithm` 0x6801 is RC4, which the editor refuses.
 */
function standardEncrypted(zip: Uint8Array, password: string, algorithm = 0x660e): Uint8Array {
  const u32 = (...values: number[]) => Buffer.from(new Uint32Array(values).buffer);
  const sha1 = (...parts: Uint8Array[]) => createHash('sha1').update(Buffer.concat(parts)).digest();
  const ecb = (key: Buffer, data: Uint8Array) => {
    const c = createCipheriv('aes-128-ecb', key, null).setAutoPadding(false);
    return Buffer.concat([c.update(data), c.final()]);
  };
  const salt = randomBytes(16);
  let h = sha1(salt, Buffer.from(password, 'utf16le'));
  for (let i = 0; i < 50_000; i++) h = sha1(u32(i), h);
  h = sha1(h, u32(0));
  const x1 = sha1(Buffer.alloc(64, 0x36).map((b, i) => b ^ (h[i] ?? 0)));
  const x2 = sha1(Buffer.alloc(64, 0x5c).map((b, i) => b ^ (h[i] ?? 0)));
  const key = Buffer.concat([x1, x2]).subarray(0, 16);
  const verifier = randomBytes(16);
  const csp = Buffer.from('Microsoft Enhanced RSA and AES Cryptographic Provider\0', 'utf16le');
  // Flags (fCryptoAPI | fAES), SizeExtra, AlgID, AlgIDHash (SHA-1), KeySize, ProviderType, two reserved, then the provider's name.
  const header = Buffer.concat([u32(0x24, 0, algorithm, 0x8004, 128, 0x18, 0, 0), csp]);
  const info = Buffer.concat([
    Buffer.from(new Uint16Array([3, 2]).buffer),
    u32(0x24, header.length),
    header,
    u32(16),
    salt,
    ecb(key, verifier),
    u32(20),
    ecb(key, Buffer.concat([sha1(verifier), Buffer.alloc(12)])),
  ]);
  const padded = Buffer.concat([zip, Buffer.alloc((16 - (zip.length % 16)) % 16)]);
  const pkg = Buffer.concat([u32(zip.length, 0), ecb(key, padded)]);
  return writeCompoundFile(new Map([['EncryptionInfo', new Uint8Array(info)], ['EncryptedPackage', new Uint8Array(pkg)]]));
}

describe('hashes and cipher', () => {
  it('match Node’s for messages of every length around the block sizes', () => {
    for (const n of [0, 1, 55, 56, 63, 64, 65, 111, 112, 127, 128, 129, 1000, 70_000]) {
      const data = new Uint8Array(randomBytes(n));
      for (const [name, fn] of [['sha1', primitives.sha1], ['sha256', primitives.sha256], ['sha384', primitives.sha384], ['sha512', primitives.sha512]] as const) {
        expect(Buffer.from(fn(data)).toString('hex'), `${name} of ${n} bytes`).toBe(createHash(name).update(data).digest('hex'));
      }
      const key = new Uint8Array(randomBytes(64));
      expect(Buffer.from(primitives.hmac512(key, data)).toString('hex')).toBe(createHmac('sha512', key).update(data).digest('hex'));
    }
  });

  it('encrypts and decrypts with AES as Node does, in CBC and ECB, with 128, 192 and 256-bit keys', () => {
    for (const bits of [128, 192, 256]) {
      const key = new Uint8Array(randomBytes(bits / 8));
      const iv = new Uint8Array(randomBytes(16));
      const data = new Uint8Array(randomBytes(16 * 37));
      const node = (mode: string, decrypt: boolean, input: Uint8Array) => {
        const c = (decrypt ? createDecipheriv : createCipheriv)(`aes-${bits}-${mode}`, key, mode === 'ecb' ? null : iv).setAutoPadding(false);
        return new Uint8Array(Buffer.concat([c.update(input), c.final()]));
      };
      expect(primitives.aes(key, data, false, iv)).toEqual(node('cbc', false, data));
      expect(primitives.aes(key, data, true, iv)).toEqual(node('cbc', true, data));
      expect(primitives.aes(key, data, false)).toEqual(node('ecb', false, data));
      expect(primitives.aes(key, data, true)).toEqual(node('ecb', true, data));
    }
  });
});

describe('password-protected workbooks', () => {
  it('open with the password, as another tool encrypted them, and refuse a wrong one', () => {
    const file = corpus('outline-password.xlsx');
    expect(protectionOf(file)).toBe('password');
    expect(protectionOf(corpus('outline.xlsx'))).toBe('none');
    expect(decryptWorkbook(file, PASSWORD)).toEqual(corpus('outline.xlsx'));
    expect(() => decryptWorkbook(file, 'wrong')).toThrow(new EncryptionError('password'));
  });

  it('are encrypted as Office encrypts them and open again with the same password', () => {
    const zip = corpus('outline.xlsx');
    const file = encryptWorkbook(zip, 'pässwörd ✓', 1000);
    expect(protectionOf(file)).toBe('password');
    const streams = readCompoundFile(file);
    expect([...streams.keys()].sort()).toEqual([
      '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace',
      '\u0006DataSpaces/DataSpaceMap',
      '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary',
      '\u0006DataSpaces/Version',
      'EncryptedPackage',
      'EncryptionInfo',
    ]);
    // The data spaces as another tool writes them, byte for byte.
    const theirs = readCompoundFile(corpus('outline-password.xlsx'));
    for (const path of [...theirs.keys()].filter((p) => p.startsWith('\u0006'))) expect(streams.get(path), path).toEqual(theirs.get(path));
    expect(decryptWorkbook(file, 'pässwörd ✓')).toEqual(zip);
    expect(() => decryptWorkbook(file, 'password')).toThrow(new EncryptionError('password'));
  });

  it('open when encrypted as Office 2007 did (Standard encryption), and refuse a wrong password or RC4', () => {
    const zip = corpus('outline.xlsx');
    const file = standardEncrypted(zip, 'secret');
    expect(protectionOf(file)).toBe('password');
    expect(decryptWorkbook(file, 'secret')).toEqual(zip);
    expect(() => decryptWorkbook(file, 'Secret')).toThrow(new EncryptionError('password'));
    expect(() => decryptWorkbook(standardEncrypted(zip, 'secret', 0x6801), 'secret')).toThrow(new EncryptionError('unsupported'));
  });

  it('keep streams of every size in a compound file, small ones in its mini stream', () => {
    const streams = new Map([1, 63, 64, 4095, 4096, 4097, 100_000].map((n, i) => [`${i % 2 ? 'Dir/' : ''}s${n}`, new Uint8Array(randomBytes(n))]));
    expect(readCompoundFile(writeCompoundFile(streams))).toEqual(streams);
  });
});

describe('opening a password-protected workbook', () => {
  const failure = (fn: () => unknown) => {
    try {
      fn();
    } catch (e) {
      return english.text((e as { text?: Parameters<typeof english.text>[0] }).text ?? String(e));
    }
    return 'opened';
  };

  it('asks for its password, refuses a wrong one, and opens with the right one, saying so in the report', () => {
    const file = corpus('outline-password.xlsx');
    expect(failure(() => WorkbookSession.open(file, 'o.xlsx'))).toBe('This workbook is password-protected. Enter its password to open it.');
    expect(failure(() => WorkbookSession.open(file, 'o.xlsx', 'en-US', { password: 'ruled-2026' }))).toBe("That password isn't right. Passwords are case-sensitive.");
    const s = WorkbookSession.open(file, 'o.xlsx', 'en-US', { password: PASSWORD });
    expect(s.encrypted).toBe(true);
    expect(s.model.sheets.map((m) => m.info.name)).toContain('Report');
    expect(s.model.findings[0]).toMatchObject({ kind: 'encryption' });
  });

  it('saves it encrypted with the same password, and unchanged as it was', () => {
    const file = corpus('outline-password.xlsx');
    const s = WorkbookSession.open(file, 'o.xlsx', 'en-US', { password: PASSWORD });
    expect(s.save()).toBe(file);
    s.apply({ type: 'input', sheet: 0, row: 0, col: 0, text: 'Changed' });
    const saved = s.save();
    expect(protectionOf(saved)).toBe('password');
    const again = WorkbookSession.open(saved, 'o.xlsx', 'en-US', { password: PASSWORD });
    expect(loadedValue(again, again.model.sheets[0].info.name, 'A1')).toBe('Changed');
  });

  it('refuses rights-managed files and old .xls files, each with its own reason', () => {
    const rights = writeCompoundFile(new Map([['\u0006DataSpaces/TransformInfo/DRMEncryptedTransform/\u0006Primary', new Uint8Array(8)], ['EncryptedPackage', new Uint8Array(16)]]));
    expect(failure(() => WorkbookSession.open(rights, 'r.xlsx'))).toMatch(/sensitivity label or rights management/);
    const xls = writeCompoundFile(new Map([['Workbook', new Uint8Array(5000)]]));
    expect(failure(() => WorkbookSession.open(xls, 'old.xlsx'))).toMatch(/old \.xls format/);
  });
});
