import { xml2js, js2xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { unzipSync, zipSync, strFromU8, strToU8, type Zippable } from "fflate";

const XML_PARSE_OPTIONS = {
  nativeTypeAttributes: true,
  captureSpacesBetweenElements: true,
};

/**
 * Parsed OOXML archive backed by an unzipped ZIP map.
 *
 * Provides unstorage-style API (get/set/getRaw/setRaw/remove/has/keys)
 * for reading and modifying individual parts, then serializing back to a ZIP buffer.
 */
export class ParsedArchive {
  private readonly zip: Map<string, Uint8Array>;
  private readonly modified = new Map<string, Uint8Array>();
  private readonly wrapperCache = new Map<string, Element>();

  public constructor(data: Uint8Array) {
    this.zip = new Map(Object.entries(unzipSync(data)));
  }

  /** Read an XML part as an Element tree. */
  public get(path: string): Element | undefined {
    // Check modified first
    const modData = this.modified.get(path);
    if (modData) {
      const wrapper = xml2js(strFromU8(modData), XML_PARSE_OPTIONS) as Element;
      this.wrapperCache.set(path, wrapper);
      return wrapper.elements?.find((e) => e.type === "element");
    }

    const data = this.zip.get(path);
    if (data === undefined) return undefined;

    // Try cache
    const cached = this.wrapperCache.get(path);
    if (cached) return cached.elements?.find((e) => e.type === "element");

    // Parse and cache
    const wrapper = xml2js(strFromU8(data), XML_PARSE_OPTIONS) as Element;
    this.wrapperCache.set(path, wrapper);
    return wrapper.elements?.find((e) => e.type === "element");
  }

  /** Write an XML part (Element → XML string). */
  public set(path: string, element: Element): void {
    const wrapper = this.wrapperCache.get(path);
    const doc: Element = wrapper
      ? { ...wrapper, elements: [{ ...element, type: "element" as const }] }
      : { elements: [{ ...element, type: "element" as const }] };
    const xml = js2xml(doc);
    this.modified.set(path, strToU8(xml));
  }

  /** Read raw binary data (images, media, etc.). */
  public getRaw(path: string): Uint8Array | undefined {
    return this.modified.get(path) ?? this.zip.get(path);
  }

  /** Write raw binary data. */
  public setRaw(path: string, data: Uint8Array): void {
    this.modified.set(path, data);
    this.wrapperCache.delete(path);
  }

  /** Remove a part. Returns true if it existed. */
  public remove(path: string): boolean {
    this.wrapperCache.delete(path);
    return this.modified.delete(path) || this.zip.delete(path);
  }

  /** Check if a part exists. */
  public has(path: string): boolean {
    return this.modified.has(path) || this.zip.has(path);
  }

  /** List all paths matching an optional prefix. */
  public keys(prefix?: string): string[] {
    const all = new Set<string>();
    for (const key of this.zip.keys()) {
      if (!prefix || key.startsWith(prefix)) all.add(key);
    }
    for (const key of this.modified.keys()) {
      if (!prefix || key.startsWith(prefix)) all.add(key);
    }
    return [...all];
  }

  /** Serialize back to a ZIP buffer, merging original zip + modifications. */
  public save(): Uint8Array {
    const files: Zippable = {};
    for (const [path, data] of this.zip) {
      if (!this.modified.has(path)) {
        files[path] = isMediaPath(path) ? [data, { level: MEDIA_STORED_LEVEL }] : data;
      }
    }
    for (const [path, data] of this.modified) {
      files[path] = isMediaPath(path) ? [data, { level: MEDIA_STORED_LEVEL }] : data;
    }
    return zipSync(files);
  }
}

/** Set of already-compressed media file extensions (lowercase, with dot prefix). */
const MEDIA_EXTENSIONS = new Set([
  ".png",
  ".jpg",
  ".jpeg",
  ".gif",
  ".wmf",
  ".emf",
  ".tif",
  ".tiff",
  ".avi",
  ".mp4",
  ".mp3",
  ".wav",
]);

/** Check if a file path has a media extension that should use STORE (no compression). */
function isMediaPath(path: string): boolean {
  const dot = path.lastIndexOf(".");
  if (dot === -1) return false;
  return MEDIA_EXTENSIONS.has(path.slice(dot).toLowerCase());
}

/** STORE level for already-compressed media formats. */
const MEDIA_STORED_LEVEL = 0;

/** Parse an OOXML archive (.docx, .pptx, .xlsx) into a ParsedArchive. */
export function parseArchive(data: Uint8Array): ParsedArchive {
  return new ParsedArchive(data);
}
