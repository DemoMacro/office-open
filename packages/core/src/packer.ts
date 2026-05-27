/**
 * Shared packer utilities for OOXML document generation.
 *
 * @module
 */

import { Readable } from "stream";

import {
  type ZipOptions,
  type Zippable,
  AsyncZipDeflate,
  Zip,
  ZipPassThrough,
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

export const PrettifyType = {
  NONE: "",
  WITH_2_BLANKS: "  ",
  WITH_4_BLANKS: "    ",
  WITH_TAB: "\t",
} as const;

export const convertPrettifyType = (
  prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
): (typeof PrettifyType)[keyof typeof PrettifyType] | undefined =>
  prettify === true ? PrettifyType.WITH_2_BLANKS : prettify === false ? undefined : prettify;

/** Default DEFLATE compression level matching Microsoft Office behavior. */
export const ZIP_DEFLATE_LEVEL = 6;

/** STORE level for already-compressed media formats (JPEG, PNG, GIF, etc.). */
export const ZIP_STORED_LEVEL = 0;

/**
 * Compress files and convert to the requested output format.
 *
 * Uses synchronous DEFLATE compression for maximum throughput.
 * XML entries use DEFLATE by default; media entries should explicitly set
 * `{ level: ZIP_STORED_LEVEL }` to avoid redundant compression.
 */
export const zipAndConvert = <T extends OutputType>(
  files: Zippable,
  type: T,
  mimeType: string,
  level: number = ZIP_DEFLATE_LEVEL,
): OutputByType[T] => {
  const zipped = zipSync(files, { level: level as ZipOptions["level"] });
  return convertOutput(zipped, type, mimeType);
};

/**
 * Create a Node.js Readable stream from compressed file entries.
 *
 * Uses fflate's AsyncZipDeflate for non-blocking DEFLATE compression via
 * Web Workers. STORED entries (media) pass through synchronously.
 * Chunks are emitted asynchronously as each file's compression completes.
 */
export const createZipStream = (files: Zippable): Readable => {
  const stream = new Readable({ read() {} });

  try {
    const zip = new Zip((err, chunk, final) => {
      if (err) {
        stream.destroy(err);
        return;
      }
      if (!stream.destroyed) {
        stream.push(Buffer.from(chunk));
      }
      if (final) {
        stream.push(null);
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
    stream.destroy(err instanceof Error ? err : new Error(String(err)));
  }

  return stream;
};

// ── PrettifyType alias for convenience method signatures ──

type PrettifyValue = (typeof PrettifyType)[keyof typeof PrettifyType];

// ── Factory function ──

/**
 * Compile function provided by each package to convert a file object into a Zippable map.
 */
export type CompileFn<TFile> = (
  file: TFile,
  prettify?: PrettifyValue,
  overrides?: readonly XmlifyedFile[],
) => Zippable;

/**
 * Packer interface returned by {@link createPacker}.
 */
export interface Packer<TFile> {
  compile: CompileFn<TFile>;
  pack<T extends OutputType>(
    file: TFile,
    type: T,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<OutputByType[T]>;
  toString(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<string>;
  toBuffer(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<Buffer>;
  toBase64String(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<string>;
  toBlob(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<Blob>;
  toArrayBuffer(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Promise<ArrayBuffer>;
  toStream(
    file: TFile,
    prettify?: boolean | PrettifyValue,
    overrides?: readonly XmlifyedFile[],
  ): Readable;
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

  const pack = async <T extends OutputType>(
    file: TFile,
    type: T,
    prettify?: boolean | PrettifyValue,
    overrides: readonly XmlifyedFile[] = [],
  ): Promise<OutputByType[T]> => {
    const files = compile(file, convertPrettifyType(prettify), overrides);
    return zipAndConvert(files, type, mimeType);
  };

  return {
    compile,
    pack,

    toString: (file, prettify, overrides) =>
      Promise.resolve(pack(file, "string", prettify, overrides)),

    toBuffer: (file, prettify, overrides) =>
      Promise.resolve(pack(file, "nodebuffer", prettify, overrides)),

    toBase64String: (file, prettify, overrides) =>
      Promise.resolve(pack(file, "base64", prettify, overrides)),

    toBlob: (file, prettify, overrides) => Promise.resolve(pack(file, "blob", prettify, overrides)),

    toArrayBuffer: (file, prettify, overrides) =>
      Promise.resolve(pack(file, "arraybuffer", prettify, overrides)),

    toStream: (file, prettify, overrides) => {
      let files: Zippable;
      try {
        files = compile(file, convertPrettifyType(prettify), overrides);
      } catch (err) {
        const errorStream = new Readable({ read() {} });
        errorStream.destroy(err instanceof Error ? err : new Error(String(err)));
        return errorStream;
      }
      return createZipStream(files);
    },
  };
};
