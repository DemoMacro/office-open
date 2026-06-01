/**
 * Native ZIP writer with automatic fallback to fflate.
 *
 * In Node.js: uses `node:zlib` for DEFLATE compression and CRC-32,
 * providing ~2-3x faster DEFLATE than fflate's pure-JS implementation.
 * In browsers: falls back to fflate (deflate-raw in CompressionStream
 * is only available in Chromium, not Firefox/Safari).
 *
 * @module
 */

import type { Zippable, ZipOptions } from "fflate";

// ── CRC-32 lookup table ──

const CRC_TABLE = new Int32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = c & 1 ? (c >>> 1) ^ -306674912 : c >>> 1;
  CRC_TABLE[i] = c;
}

function computeCrc32(data: Uint8Array): number {
  let crc = -1;
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return ~crc;
}

// ── Native zlib detection (Node.js ESM) ──

type DeflateFn = (data: Uint8Array, level: number) => Uint8Array;
type AsyncDeflateFn = (data: Uint8Array, level: number) => Promise<Uint8Array>;
type Crc32Fn = (data: Uint8Array) => number;

let _nativeDeflate: DeflateFn | undefined;
let _nativeDeflateAsync: AsyncDeflateFn | undefined;
let _nativeCrc32: Crc32Fn | undefined;

try {
  // top-level await — ESM standard.
  // In Node.js: resolves to zlib module.
  // In browsers: import("node:zlib") throws → caught → fflate fallback.
  const zlib = await import("node:zlib");
  _nativeDeflate = (data: Uint8Array, level: number): Uint8Array =>
    zlib.deflateRawSync(data, { level });
  _nativeDeflateAsync = (data: Uint8Array, level: number): Promise<Uint8Array> =>
    new Promise<Uint8Array>((resolve, reject) =>
      zlib.deflateRaw(data, { level }, (err, result) => {
        if (err) reject(err);
        else resolve(result);
      }),
    );
  _nativeCrc32 = zlib.crc32 ? (data: Uint8Array) => zlib.crc32(data) : computeCrc32;
} catch {
  // Browser or Node.js without zlib — fflate fallback
}

export const hasNativeDeflate = (): boolean => _nativeDeflate !== undefined;

// ── LE write helpers ──

function wU16(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >> 8) & 0xff;
}

function wU32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >> 8) & 0xff;
  b[o + 2] = (v >> 16) & 0xff;
  b[o + 3] = (v >> 24) & 0xff;
}

// ── Per-entry metadata ──

interface Entry {
  filename: Uint8Array;
  data: Uint8Array;
  uncompressedSize: number;
  crc: number;
  method: number; // 0 = STORE, 8 = DEFLATE
  localOffset: number;
}

// ── Phase 1 helpers ──

function resolveEntryData(
  raw: Zippable[string],
  defaultLevel: number,
): { data: Uint8Array; level: number } {
  if (Array.isArray(raw)) {
    return { data: raw[0] as Uint8Array, level: (raw[1] as ZipOptions).level ?? defaultLevel };
  }
  return { data: raw as Uint8Array, level: defaultLevel };
}

function compressOne(
  data: Uint8Array,
  level: number,
  method: number,
  deflate: DeflateFn,
): { compressed: Uint8Array; method: number } {
  if (level === 0) return { compressed: data, method: 0 };
  const compressed = deflate(data, level);
  if (compressed.length >= data.length) return { compressed: data, method: 0 };
  return { compressed, method: 8 };
}

async function compressOneAsync(
  data: Uint8Array,
  level: number,
  deflate: AsyncDeflateFn,
): Promise<{ compressed: Uint8Array; method: number }> {
  if (level === 0) return { compressed: data, method: 0 };
  const compressed = await deflate(data, level);
  if (compressed.length >= data.length) return { compressed: data, method: 0 };
  return { compressed, method: 8 };
}

// ── Phase 2+3: calculate sizes + write ZIP buffer ──

function writeZipBuffer(entries: Entry[]): Uint8Array {
  let totalSize = 0;
  for (let i = 0; i < entries.length; i++) {
    totalSize += 30 + entries[i].filename.length + entries[i].data.length;
  }
  const cdOffset = totalSize;
  let cdSize = 0;
  for (let i = 0; i < entries.length; i++) {
    cdSize += 46 + entries[i].filename.length;
  }
  totalSize += cdSize + 22;

  const buf = new Uint8Array(totalSize);
  let offset = 0;

  // Local file headers + data
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    e.localOffset = offset;
    wU32(buf, offset, 0x04034b50); // signature
    wU16(buf, offset + 4, 20); // version needed
    wU16(buf, offset + 6, 0); // flags
    wU16(buf, offset + 8, e.method); // compression method
    wU16(buf, offset + 10, 0); // mod time
    wU16(buf, offset + 12, 0); // mod date
    wU32(buf, offset + 14, e.crc);
    wU32(buf, offset + 18, e.data.length); // compressed size
    wU32(buf, offset + 22, e.uncompressedSize);
    wU16(buf, offset + 26, e.filename.length);
    wU16(buf, offset + 28, 0); // extra field length
    offset += 30;

    buf.set(e.filename, offset);
    offset += e.filename.length;

    buf.set(e.data, offset);
    offset += e.data.length;
  }

  // Central directory
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    wU32(buf, offset, 0x02014b50); // signature
    wU16(buf, offset + 4, 20); // version made by
    wU16(buf, offset + 6, 20); // version needed
    wU16(buf, offset + 8, 0); // flags
    wU16(buf, offset + 10, e.method);
    wU16(buf, offset + 12, 0); // mod time
    wU16(buf, offset + 14, 0); // mod date
    wU32(buf, offset + 16, e.crc);
    wU32(buf, offset + 20, e.data.length);
    wU32(buf, offset + 24, e.uncompressedSize);
    wU16(buf, offset + 28, e.filename.length);
    wU16(buf, offset + 30, 0); // extra field length
    wU16(buf, offset + 32, 0); // comment length
    wU16(buf, offset + 34, 0); // disk number start
    wU16(buf, offset + 36, 0); // internal attrs
    wU32(buf, offset + 38, 0); // external attrs
    wU32(buf, offset + 42, e.localOffset);
    offset += 46;

    buf.set(e.filename, offset);
    offset += e.filename.length;
  }

  // End of Central Directory Record
  wU32(buf, offset, 0x06054b50);
  wU16(buf, offset + 4, 0); // disk number
  wU16(buf, offset + 6, 0); // CD disk
  wU16(buf, offset + 8, entries.length);
  wU16(buf, offset + 10, entries.length);
  wU32(buf, offset + 12, cdSize);
  wU32(buf, offset + 16, cdOffset);
  wU16(buf, offset + 20, 0); // comment length

  return buf;
}

// ── Public API ──

const textEncoder = new TextEncoder();

export function nativeZip(files: Zippable, level: number = 6): Uint8Array {
  if (!_nativeDeflate) throw new Error("Native deflate not available");
  const crc = _nativeCrc32 ?? computeCrc32;
  const keys = Object.keys(files);
  const entries: Entry[] = Array.from({ length: keys.length });

  for (let i = 0; i < keys.length; i++) {
    const { data, level: entryLevel } = resolveEntryData(files[keys[i]], level);
    const c = crc(data);
    const { compressed, method } = compressOne(data, entryLevel, 8, _nativeDeflate);
    entries[i] = {
      filename: textEncoder.encode(keys[i]),
      data: compressed,
      uncompressedSize: data.length,
      crc: c,
      method,
      localOffset: 0,
    };
  }

  return writeZipBuffer(entries);
}

export async function nativeZipAsync(files: Zippable, level: number = 6): Promise<Uint8Array> {
  if (!_nativeDeflateAsync) throw new Error("Native async deflate not available");
  const crc = _nativeCrc32 ?? computeCrc32;
  const keys = Object.keys(files);
  const entries: Entry[] = Array.from({ length: keys.length });

  for (let i = 0; i < keys.length; i++) {
    const { data, level: entryLevel } = resolveEntryData(files[keys[i]], level);
    const c = crc(data);
    const { compressed, method } = await compressOneAsync(data, entryLevel, _nativeDeflateAsync);
    entries[i] = {
      filename: textEncoder.encode(keys[i]),
      data: compressed,
      uncompressedSize: data.length,
      crc: c,
      method,
      localOffset: 0,
    };
  }

  return writeZipBuffer(entries);
}
