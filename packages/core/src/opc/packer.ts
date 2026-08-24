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
  ZipDeflate,
  ZipPassThrough,
  zip,
  zipSync,
} from "fflate";

import { convertOutput } from "./output";
import type { OutputByType, OutputType } from "./output";
import {
  hasNativeDeflate,
  isBunRuntime,
  NativeZipDeflate,
  nativeZip,
  nativeZipAsync,
  nativeZipStream,
  prefersMainThreadDeflate,
} from "./zip-native";

export type { Zippable, ZipOptions } from "fflate";
export { strFromU8, unzipSync, zipSync } from "fflate";

export interface XmlifyedFile {
  path: string;
  data: string | Uint8Array;
}

/** Default DEFLATE level for XML entries (SuperFast, matching MS Office). */
export const ZIP_DEFLATE_LEVEL = 1;

/** Default DEFLATE level for compressible media (EMF/WMF/BMP/TIFF/SVG). MS
 *  Office uses CompressionOption.Normal (~zlib 6); SuperFast (1) inflates EMF
 *  output ~15% (measured against a real Office-generated package). */
export const ZIP_MEDIA_LEVEL = 6;

/** Level for already-compressed media — STORE (no compression). */
export const ZIP_STORED_LEVEL = 0;

/**
 * Media formats already compressed internally (DEFLATE for PNG, DCT for JPEG,
 * LZW for GIF; video/audio containers carry their own codec streams).
 * Re-compressing via zip DEFLATE wastes CPU and often inflates the data, so
 * MS Office STORE-s these (CompressionOption.NotCompressed → zip method 0).
 * Everything else (EMF/WMF/BMP/TIFF/SVG/WAV…) is compressible → DEFLATE —
 * measured on a real Office package, EMF at zlib level 6 matches Office
 * output where SuperFast inflates ~15%.
 */
const PRECOMPRESSED_MEDIA_EXT = new Set([
  "jpg",
  "jpeg",
  "png",
  "gif",
  "mp4",
  "mov",
  "wmv",
  "mpg",
  "mpeg",
  "mp3",
  "wma",
  "aac",
]);

/**
 * Resolve the ZIP level for a media entry by file-name extension, matching MS
 * Office: already-compressed raster formats → STORE (0), everything else →
 * DEFLATE (`mediaLevel`, default Normal). A `compression.media` override
 * therefore applies only to compressible formats, never forcing DEFLATE onto
 * pre-compressed assets.
 */
export const levelForMediaName = (fileName: string, mediaLevel: number): number => {
  const dot = fileName.lastIndexOf(".");
  const ext = dot < 0 ? "" : fileName.slice(dot + 1).toLowerCase();
  return PRECOMPRESSED_MEDIA_EXT.has(ext) ? ZIP_STORED_LEVEL : mediaLevel;
};

/** Compression options for ZIP output (zlib levels 0-9, matching fflate). */
export interface CompressionOptions {
  /** DEFLATE level for XML files. Default: 1 (SuperFast, matching MS Office). */
  xml?: number;
  /**
   * DEFLATE level for compressible media (EMF/WMF/BMP/TIFF/…). Already-compressed
   * formats (PNG/JPEG/GIF) are always STOREd regardless, matching MS Office.
   * Default: 6 (Normal, matching MS Office for EMF/WMF).
   */
  media?: number;
}

/** Options for Packer output methods. */
export interface PackerOptions<T extends OutputType = "nodebuffer"> {
  /** Output format. Defaults to `"nodebuffer"` (Node.js Buffer). */
  type?: T;
  /** Custom XML/ZIP file overrides. */
  overrides?: XmlifyedFile[];
  /** Compression levels for ZIP entries. */
  compression?: CompressionOptions;
}

/**
 * Asynchronously compress files and convert to the requested output format.
 *
 * Where a native zlib resolved (Node/Bun), entries deflate in parallel on the
 * libuv thread pool; elsewhere fflate compresses off the main thread.
 * XML entries use DEFLATE level 1 (SuperFast) by default. Media entries are
 * split by type, matching MS Office: already-compressed formats (PNG/JPEG/GIF)
 * are STOREd, everything else uses the `media` level (default Normal).
 * Set `{ media: ZIP_STORED_LEVEL }` to STORE all compressible media too.
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
  // Node's native zlib compresses the entries in parallel on the libuv pool —
  // vastly faster than fflate's per-entry workers, whose spawn cost dominates
  // small packages (~20 vs ~2700 ops/s on a 12-part stream, measured). Bun
  // stays on fflate main-thread deflate (see isBunRuntime), browsers/Deno on
  // the worker paths below.
  if (hasNativeDeflate() && !isBunRuntime()) {
    return nativeZipStream(files, defaultLevel);
  }
  // Remaining runtimes: Bun deflates on the main thread (its worker teardown
  // lags creation, accumulating cost across back-to-back generations — see
  // prefersMainThreadDeflate); browsers/Deno keep AsyncZipDeflate, where web
  // workers are cheap to spawn and deflate off the main thread for real.
  const mainThread = prefersMainThreadDeflate();
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
          const workerPath = level !== ZIP_STORED_LEVEL && !mainThread;
          const entry =
            level === ZIP_STORED_LEVEL
              ? new ZipPassThrough(name)
              : workerPath
                ? new AsyncZipDeflate(name, { level: level as ZipOptions["level"] })
                : new ZipDeflate(name, { level: level as ZipOptions["level"] });
          zip.add(entry);
          // AsyncZipDeflate transfers each pushed buffer to its worker
          // (postMessage move semantics), detaching the caller's view —
          // regenerating from the same files would push an already-detached
          // buffer and throw DataCloneError. Only the worker entry needs the
          // copy; pass-through and main-thread entries keep zero-copy.
          entry.push(workerPath ? raw.slice() : raw, true);
        }

        zip.end();
      } catch (err) {
        controller.error(err instanceof Error ? err : new Error(String(err)));
      }
    },
  });
};

// ── Streaming ZIP writer ──

/**
 * Incremental sink for one ZIP part. Feed chunks as they are produced; the
 * entry's local header is emitted on {@link ZipStreamWriter.addPart} and the
 * trailing CRC/sizes on {@link end} — the part's full content never has to
 * exist in memory at once.
 */
export interface ZipPartSink {
  /** Append a chunk to the part. */
  push(chunk: Uint8Array): void;
  /** Finalize the part, optionally with a final chunk. Must be called before adding the next one. */
  end(lastChunk?: Uint8Array): void;
}

/**
 * Incremental ZIP archive writer over fflate's streaming `Zip` — parts are
 * added and completed one at a time; compressed output flows to `ondata` as
 * it is produced. This is the constant-memory counterpart to
 * {@link createZipStream} (which requires the complete `Zippable` up front):
 * serialize each part chunk-wise into the returned sink instead of building
 * the full part string.
 *
 * Parts must be added in order — `[Content_Types].xml` first per OPC. The
 * whole archive finishes when {@link ZipStreamWriter.end} is called.
 */
export class ZipStreamWriter {
  private readonly zip: Zip;
  private current: AsyncZipDeflate | NativeZipDeflate | ZipDeflate | ZipPassThrough | undefined;
  // Node streams each part through a native zlib DeflateRaw on the libuv pool
  // (no spawn, compression overlaps with chunk production). Bun deflates on
  // the main thread; browsers/Deno use fflate's per-entry workers.
  private readonly nativeEntry = hasNativeDeflate() && !isBunRuntime();
  private readonly mainThread = prefersMainThreadDeflate();

  constructor(
    ondata: (err: Error | null, chunk: Uint8Array, final: boolean) => void,
    private readonly defaultLevel: number = ZIP_DEFLATE_LEVEL,
  ) {
    this.zip = new Zip(ondata);
  }

  /** Add a part and return its incremental sink. `end()` the previous first. */
  addPart(name: string, level: number = this.defaultLevel): ZipPartSink {
    if (this.current) throw new Error(`ZipStreamWriter: previous part not finalized`);
    const deflatePath = level !== ZIP_STORED_LEVEL;
    const workerPath = deflatePath && !this.mainThread;
    const entry = !deflatePath
      ? new ZipPassThrough(name)
      : this.nativeEntry
        ? new NativeZipDeflate(name, { level: level as ZipOptions["level"] })
        : workerPath
          ? new AsyncZipDeflate(name, { level: level as ZipOptions["level"] })
          : new ZipDeflate(name, { level: level as ZipOptions["level"] });
    // AsyncZipDeflate hands chunks to a worker; ondata delivery is asynchronous,
    // so ordering between parts is preserved by fflate's internal queue. The
    // worker transfer-detaches every chunk it is handed, so sinks receive
    // copies — a caller-held buffer must survive repeated generations.
    this.zip.add(entry);
    this.current = entry;
    return {
      push: (chunk) => entry.push(workerPath ? chunk.slice() : chunk, false),
      end: (lastChunk?: Uint8Array) => {
        entry.push((workerPath ? lastChunk?.slice() : lastChunk) ?? new Uint8Array(0), true);
        this.current = undefined;
      },
    };
  }

  /** Finish the archive (central directory + EOCD). */
  end(): void {
    if (this.current) throw new Error(`ZipStreamWriter: unclosed part at end()`);
    this.zip.end();
  }
}

// ── Factory function ──

/**
 * Compile function provided by each package to convert a file object into a Zippable map.
 */
export type CompileFn<TFile> = (
  file: TFile,
  overrides?: XmlifyedFile[],
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
  toBase64(file: TFile, options?: PackerOptions): Promise<string>;
  /** Sync → `string` (base64-encoded). */
  toBase64Sync(file: TFile, options?: PackerOptions): string;

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
  compile: CompileFn<TFile>;
  mimeType: string;
}): Packer<TFile> => {
  const { compile, mimeType } = options;

  const pack = async <T extends OutputType = "nodebuffer">(
    file: TFile,
    opts?: PackerOptions<T>,
  ): Promise<OutputByType[T]> => {
    const type = opts?.type ?? ("nodebuffer" as T);
    const files = compile(file, opts?.overrides ?? [], opts?.compression?.media ?? ZIP_MEDIA_LEVEL);
    return zipAndConvert(files, type, mimeType, opts?.compression?.xml ?? ZIP_DEFLATE_LEVEL);
  };

  const toBytes = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "uint8array" });
  const toString = (file: TFile, opts?: PackerOptions) => pack(file, { ...opts, type: "string" });
  const toBuffer = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "nodebuffer" });
  const toBase64 = (file: TFile, opts?: PackerOptions) => pack(file, { ...opts, type: "base64" });
  const toBlob = (file: TFile, opts?: PackerOptions) => pack(file, { ...opts, type: "blob" });
  const toArrayBuffer = (file: TFile, opts?: PackerOptions) =>
    pack(file, { ...opts, type: "arraybuffer" });

  // ── Sync methods (zipSync, maximum throughput) ──

  const packSync = <T extends OutputType = "nodebuffer">(
    file: TFile,
    opts?: PackerOptions<T>,
  ): OutputByType[T] => {
    const type = opts?.type ?? ("nodebuffer" as T);
    const files = compile(file, opts?.overrides ?? [], opts?.compression?.media ?? ZIP_MEDIA_LEVEL);
    return zipSyncAndConvert(files, type, mimeType, opts?.compression?.xml ?? ZIP_DEFLATE_LEVEL);
  };

  const toBytesSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "uint8array" });
  const toStringSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "string" });
  const toBufferSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "nodebuffer" });
  const toBase64Sync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "base64" });
  const toBlobSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "blob" });
  const toArrayBufferSync = (file: TFile, opts?: PackerOptions) =>
    packSync(file, { ...opts, type: "arraybuffer" });

  // ── Stream ──

  const toStream = (file: TFile, opts?: PackerOptions) => {
    const mediaLevel = opts?.compression?.media ?? ZIP_MEDIA_LEVEL;
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
    toBase64,
    toBase64Sync,
    toBlob,
    toBlobSync,
    toArrayBuffer,
    toArrayBufferSync,
    toStream,
  };
};
