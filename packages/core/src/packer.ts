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

export type { Zippable, ZipOptions } from "fflate";
export { strFromU8, unzipSync } from "fflate";

export interface XmlifyedFile {
  readonly path: string;
  readonly data: string | Uint8Array;
}

/** Default DEFLATE compression level matching Microsoft Office behavior. */
export const ZIP_DEFLATE_LEVEL = 6;

/** STORE level for already-compressed media formats (JPEG, PNG, GIF, etc.). */
export const ZIP_STORED_LEVEL = 0;

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
  const zipped = await new Promise<Uint8Array>((resolve, reject) => {
    zip(files as AsyncZippable, { level: level as ZipOptions["level"] }, (err, data) => {
      if (err) {
        reject(err);
      } else {
        resolve(data);
      }
    });
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
  const zipped = zipSync(files, { level: level as ZipOptions["level"] });
  return convertOutput(zipped, type, mimeType);
};

/**
 * Create a `ReadableStream<Uint8Array>` from compressed file entries.
 *
 * Uses fflate's `AsyncZipDeflate` for non-blocking DEFLATE compression.
 * `STORED` entries (media) pass through synchronously.
 * Works in both Node.js and browsers (Web Streams API).
 */
export const createZipStream = (files: Zippable): ReadableStream<Uint8Array> => {
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
            ? ((data[1] as ZipOptions).level ?? ZIP_DEFLATE_LEVEL)
            : ZIP_DEFLATE_LEVEL;
          const entry =
            level === ZIP_STORED_LEVEL
              ? new ZipPassThrough(name)
              : new AsyncZipDeflate(name, { level });
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
export type CompileFn<TFile> = (file: TFile, overrides?: readonly XmlifyedFile[]) => Zippable;

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
  pack<T extends OutputType>(
    file: TFile,
    type: T,
    overrides?: readonly XmlifyedFile[],
  ): Promise<OutputByType[T]>;
  /** Generic sync output — returns the requested OutputType. */
  packSync<T extends OutputType>(
    file: TFile,
    type: T,
    overrides?: readonly XmlifyedFile[],
  ): OutputByType[T];

  /** Async → `Promise<Uint8Array>` (like `Response.bytes()`). */
  toBytes(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<Uint8Array>;
  /** Sync → `Uint8Array`. */
  toBytesSync(file: TFile, overrides?: readonly XmlifyedFile[]): Uint8Array;

  /** Async → `Promise<string>` (raw ZIP content as string). */
  toString(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<string>;
  /** Sync → `string`. */
  toStringSync(file: TFile, overrides?: readonly XmlifyedFile[]): string;

  /** Async → `Promise<Buffer>` (Node.js). */
  toBuffer(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<Buffer>;
  /** Sync → `Buffer` (Node.js). */
  toBufferSync(file: TFile, overrides?: readonly XmlifyedFile[]): Buffer;

  /** Async → `Promise<string>` (base64-encoded). */
  toBase64String(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<string>;
  /** Sync → `string` (base64-encoded). */
  toBase64StringSync(file: TFile, overrides?: readonly XmlifyedFile[]): string;

  /** Async → `Promise<Blob>` (browser). */
  toBlob(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<Blob>;
  /** Sync → `Blob`. */
  toBlobSync(file: TFile, overrides?: readonly XmlifyedFile[]): Blob;

  /** Async → `Promise<ArrayBuffer>`. */
  toArrayBuffer(file: TFile, overrides?: readonly XmlifyedFile[]): Promise<ArrayBuffer>;
  /** Sync → `ArrayBuffer`. */
  toArrayBufferSync(file: TFile, overrides?: readonly XmlifyedFile[]): ArrayBuffer;

  /** Streaming output via `ReadableStream<Uint8Array>` (cross-platform, uses Web Workers). */
  toStream(file: TFile, overrides?: readonly XmlifyedFile[]): ReadableStream<Uint8Array>;
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

  const compileFiles = (file: TFile, overrides: readonly XmlifyedFile[] = []): Zippable =>
    compile(file, overrides);

  // ── Async methods (fflate Web Workers) ──

  const pack = async <T extends OutputType>(
    file: TFile,
    type: T,
    overrides?: readonly XmlifyedFile[],
  ): Promise<OutputByType[T]> => {
    const files = compileFiles(file, overrides);
    return zipAndConvert(files, type, mimeType);
  };

  const toBytes = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "uint8array", overrides);

  const toString = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "string", overrides);

  const toBuffer = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "nodebuffer", overrides);

  const toBase64String = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "base64", overrides);

  const toBlob = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "blob", overrides);

  const toArrayBuffer = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    pack(file, "arraybuffer", overrides);

  // ── Sync methods (zipSync, maximum throughput) ──

  const packSync = <T extends OutputType>(
    file: TFile,
    type: T,
    overrides?: readonly XmlifyedFile[],
  ): OutputByType[T] => {
    const files = compileFiles(file, overrides);
    return zipSyncAndConvert(files, type, mimeType);
  };

  const toBytesSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "uint8array", overrides);

  const toStringSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "string", overrides);

  const toBufferSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "nodebuffer", overrides);

  const toBase64StringSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "base64", overrides);

  const toBlobSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "blob", overrides);

  const toArrayBufferSync = (file: TFile, overrides?: readonly XmlifyedFile[]) =>
    packSync(file, "arraybuffer", overrides);

  // ── Stream ──

  const toStream = (file: TFile, overrides?: readonly XmlifyedFile[]) => {
    let files: Zippable;
    try {
      files = compileFiles(file, overrides);
    } catch (err) {
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.error(err instanceof Error ? err : new Error(String(err)));
        },
      });
    }
    return createZipStream(files);
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
