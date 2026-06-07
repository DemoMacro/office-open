/**
 * Shared packer utilities for OOXML document generation.
 *
 * @module
 */

import {
  type AsyncZippable,
  type ZipOptions,
  type Zippable,
  AsyncZipDeflate,
  Zip,
  ZipPassThrough,
  zip,
  zipSync,
} from "fflate";

import { convertOutput } from "./output-type";
import type { OutputByType, OutputType } from "./output-type";
import { hasNativeDeflate, nativeZip, nativeZipAsync } from "./zip-native";

export type { Zippable, ZipOptions } from "fflate";
export { strFromU8, unzipSync } from "fflate";
export { toUint8Array } from "undio";
export type { DataType } from "undio";

export interface XmlifyedFile {
  readonly path: string;
  readonly data: string | Uint8Array;
}

/** Default DEFLATE level for XML entries (SuperFast, matching MS Office). */
export const ZIP_DEFLATE_LEVEL = 1;

/** Default level for media entries (STORE — no compression). */
export const ZIP_STORED_LEVEL = 0;

/** Compression options for ZIP output (zlib levels 0-9, matching fflate). */
export interface CompressionOptions {
  /** DEFLATE level for XML files. Default: 1 (SuperFast, matching MS Office). */
  readonly xml?: number;
  /** DEFLATE level for media files. Default: 0 (STORE, no compression). */
  readonly media?: number;
}

/** Options for Packer output methods. */
export interface PackerOptions<T extends OutputType = "nodebuffer"> {
  /** Output format. Defaults to `"nodebuffer"` (Node.js Buffer). */
  readonly type?: T;
  /** Custom XML/ZIP file overrides. */
  readonly overrides?: readonly XmlifyedFile[];
  /** Compression levels for ZIP entries. */
  readonly compression?: CompressionOptions;
}

/**
 * Asynchronously compress files and convert to the requested output format.
 *
 * Uses fflate Web Workers for non-blocking DEFLATE compression.
 * XML entries use DEFLATE by default; media entries should explicitly set
 * `{ level: ZIP_STORED_LEVEL }` to avoid redundant compression.
 */
export const zipAndConvert = async <T extends OutputType>(
  files: Zippable,
  type: T,
  mimeType: string,
  level: number = ZIP_DEFLATE_LEVEL,
): Promise<OutputByType[T]> => {
  const zipped = hasNativeDeflate()
    ? await nativeZipAsync(files, level)
    : await new Promise<Uint8Array>((resolve, reject) => {
        zip(
          files as AsyncZippable,
          { level: level as ZipOptions["level"], consume: true },
          (err, data) => {
            if (err) reject(err);
            else resolve(data);
          },
        );
      });
  return convertOutput(zipped, type, mimeType);
};

/**
 * Synchronously compress files and convert to the requested output format.
 *
 * Uses synchronous DEFLATE compression for maximum throughput.
 * Blocks the event loop — prefer {@link zipAndConvert} in server contexts.
 */
export const zipSyncAndConvert = <T extends OutputType>(
  files: Zippable,
  type: T,
  mimeType: string,
  level: number = ZIP_DEFLATE_LEVEL,
): OutputByType[T] => {
  const zipped = hasNativeDeflate()
    ? nativeZip(files, level)
    : zipSync(files, { level: level as ZipOptions["level"] });
  return convertOutput(zipped, type, mimeType);
};

/**
 * Create a `ReadableStream<Uint8Array>` from compressed file entries.
 *
 * Uses fflate's `AsyncZipDeflate` for non-blocking DEFLATE compression.
 * `STORED` entries (media) pass through synchronously.
 * Works in both Node.js and browsers (Web Streams API).
 */
export const createZipStream = (
  files: Zippable,
  defaultLevel: number = ZIP_DEFLATE_LEVEL,
): ReadableStream<Uint8Array> => {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      try {
        const zip = new Zip((err, chunk, _final) => {
          if (err) {
            controller.error(err);
            return;
          }
          controller.enqueue(chunk);
          if (_final) {
            controller.close();
          }
        });

        for (const [name, data] of Object.entries(files)) {
          const raw = Array.isArray(data) ? (data[0] as Uint8Array) : (data as Uint8Array);
          const level = Array.isArray(data)
            ? ((data[1] as ZipOptions).level ?? defaultLevel)
            : defaultLevel;
          const entry =
            level === ZIP_STORED_LEVEL
              ? new ZipPassThrough(name)
              : new AsyncZipDeflate(name, { level: level as ZipOptions["level"] });
          zip.add(entry);
          entry.push(raw, true);
        }

        zip.end();
      } catch (err) {
        controller.error(err instanceof Error ? err : new Error(String(err)));
      }
    },
  });
};

// ── Factory function ──

/**
 * Compile function provided by each package to convert a file object into a Zippable map.
 */
export type CompileFn<TFile> = (
  file: TFile,
  overrides?: readonly XmlifyedFile[],
  mediaLevel?: number,
) => Zippable;

/**
 * Packer interface returned by {@link createPacker}.
 *
 * Async methods use fflate Web Workers for non-blocking compression.
 * Sync methods use synchronous compression for maximum throughput in
 * CLI scripts and build tools.
 */
export interface Packer<TFile> {
  /** Compile file to Zippable map (synchronous). */
  compile: CompileFn<TFile>;

  /** Generic async output — returns the requested OutputType. */
  pack<T extends OutputType = "nodebuffer">(
    file: TFile,
    options?: PackerOptions<T>,
  ): Promise<OutputByType[T]>;
  /** Generic sync output — returns the requested OutputType. */
  packSync<T extends OutputType = "nodebuffer">(
    file: TFile,
    options?: PackerOptions<T>,
  ): OutputByType[T];

  /** Async → `Promise<Uint8Array>` (like `Response.bytes()`). */
  toBytes(file: TFile, options?: PackerOptions): Promise<Uint8Array>;
  /** Sync → `Uint8Array`. */
  toBytesSync(file: TFile, options?: PackerOptions): Uint8Array;

  /** Async → `Promise<string>` (raw ZIP content as string). */
  toString(file: TFile, options?: PackerOptions): Promise<string>;
  /** Sync → `string`. */
  toStringSync(file: TFile, options?: PackerOptions): string;

  /** Async → `Promise<Buffer>` (Node.js). */
  toBuffer(file: TFile, options?: PackerOptions): Promise<Buffer>;
  /** Sync → `Buffer` (Node.js). */
  toBufferSync(file: TFile, options?: PackerOptions): Buffer;

  /** Async → `Promise<string>` (base64-encoded). */
  toBase64String(file: TFile, options?: PackerOptions): Promise<string>;
  /** Sync → `string` (base64-encoded). */
  toBase64StringSync(file: TFile, options?: PackerOptions): string;

  /** Async → `Promise<Blob>` (browser). */
  toBlob(file: TFile, options?: PackerOptions): Promise<Blob>;
  /** Sync → `Blob`. */
  toBlobSync(file: TFile, options?: PackerOptions): Blob;

  /** Async → `Promise<ArrayBuffer>`. */
  toArrayBuffer(file: TFile, options?: PackerOptions): Promise<ArrayBuffer>;
  /** Sync → `ArrayBuffer`. */
  toArrayBufferSync(file: TFile, options?: PackerOptions): ArrayBuffer;

  /** Streaming output via `ReadableStream<Uint8Array>` (cross-platform, uses Web Workers). */
  toStream(file: TFile, options?: PackerOptions): ReadableStream<Uint8Array>;
}

/**
 * Create a Packer object with all output format methods.
 *
 * Centralises the ZIP → convert pipeline and the streaming implementation
 * so that each OOXML package only needs to provide a `compile` function and
 * a MIME type.
 */
export const createPacker = <TFile>(options: {
  readonly compile: CompileFn<TFile>;
  readonly mimeType: string;
}): Packer<TFile> => {
  const { compile, mimeType } = options;

  const pack = async <T extends OutputType = "nodebuffer">(
    file: TFile,
    opts?: PackerOptions<T>,
  ): Promise<OutputByType[T]> => {
    const type = opts?.type ?? ("nodebuffer" as T);
    const files = compile(
      file,
      opts?.overrides ?? [],
      opts?.compression?.media ?? ZIP_STORED_LEVEL,
    );
    return zipAndConvert(files, type, mimeType, opts?.compression?.xml ?? ZIP_DEFLATE_LEVEL);
  };

  const toBytes = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "uint8array" });
  const toString = (file: TFile, opts?: PackerOptions) => pack(file, { ...opts, type: "string" });
  const toBuffer = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "nodebuffer" });
  const toBase64String = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "base64" });
  const toBlob = (file: TFile, opts?: PackerOptions) => pack(file, { ...opts, type: "blob" });
  const toArrayBuffer = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "arraybuffer" });

  // ── Sync methods (zipSync, maximum throughput) ──

  const packSync = <T extends OutputType = "nodebuffer">(
    file: TFile,
    opts?: PackerOptions<T>,
  ): OutputByType[T] => {
    const type = opts?.type ?? ("nodebuffer" as T);
    const files = compile(
      file,
      opts?.overrides ?? [],
      opts?.compression?.media ?? ZIP_STORED_LEVEL,
    );
    return zipSyncAndConvert(files, type, mimeType, opts?.compression?.xml ?? ZIP_DEFLATE_LEVEL);
  };

  const toBytesSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "uint8array" });
  const toStringSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "string" });
  const toBufferSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "nodebuffer" });
  const toBase64StringSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "base64" });
  const toBlobSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "blob" });
  const toArrayBufferSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "arraybuffer" });

  // ── Stream ──

  const toStream = (file: TFile, opts?: PackerOptions) => {
    const mediaLevel = opts?.compression?.media ?? ZIP_STORED_LEVEL;
    let files: Zippable;
    try {
      files = compile(file, opts?.overrides ?? [], mediaLevel);
    } catch (err) {
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.error(err instanceof Error ? err : new Error(String(err)));
        },
      });
    }
    return createZipStream(files, opts?.compression?.xml ?? ZIP_DEFLATE_LEVEL);
  };

  return {
    compile,
    pack,
    packSync,
    toBytes,
    toBytesSync,
    toString,
    toStringSync,
    toBuffer,
    toBufferSync,
    toBase64String,
    toBase64StringSync,
    toBlob,
    toBlobSync,
    toArrayBuffer,
    toArrayBufferSync,
    toStream,
  };
};
