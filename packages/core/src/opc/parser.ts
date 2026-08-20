import { parse, stringify } from "@office-open/xml";
import type { Element, ParseOptions } from "@office-open/xml";
import { unzipSync, zipSync, strFromU8, strToU8, type ZipOptions, type Zippable } from "fflate";

import { OOXML_CANONICAL_PREFIXES } from "./namespaces";
import { levelForMediaName, ZIP_MEDIA_LEVEL } from "./packer";
import { hasNativeInflate, nativeUnzip } from "./zip-native";

const XML_PARSE_OPTIONS = {
  nativeTypeAttributes: true,
  captureSpacesBetweenElements: true,
  // Normalize namespace prefixes to the canonical ones the library itself
  // emits — sources binding the Word namespace to ns0: or the spreadsheet
  // namespace to x: parse into canonical names, so descriptor matching and
  // registry paths see the elements they address.
  normalizeNamespaces: OOXML_CANONICAL_PREFIXES,
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
    // Native inflate is the fast path; on any failure (unsupported ZIP variant
    // or corruption) fall back to fflate unzipSync, the reference implementation.
    let unzipped: Record<string, Uint8Array>;
    try {
      unzipped = hasNativeInflate() ? nativeUnzip(data) : unzipSync(data);
    } catch {
      unzipped = unzipSync(data);
    }
    this.zip = new Map(Object.entries(unzipped));
  }

  /**
   * Read an XML part as an Element tree. `parseOptions` extends the default
   * XML parse options for this part (e.g. `deferElements` to capture a hot
   * container's inner XML verbatim). Callers must use consistent options per
   * path — the wrapper cache is keyed by path only.
   */
  public get(path: string, parseOptions?: ParseOptions): Element | undefined {
    const opts = parseOptions ? { ...XML_PARSE_OPTIONS, ...parseOptions } : XML_PARSE_OPTIONS;
    // Check modified first
    const modData = this.modified.get(path);
    if (modData) {
      const wrapper = parse(strFromU8(modData), opts) as Element;
      this.wrapperCache.set(path, wrapper);
      return wrapper.elements?.find((e) => e.type === "element");
    }

    const data = this.zip.get(path);
    if (data === undefined) return undefined;

    // Try cache
    const cached = this.wrapperCache.get(path);
    if (cached) return cached.elements?.find((e) => e.type === "element");

    // Parse and cache
    const wrapper = parse(strFromU8(data), opts) as Element;
    this.wrapperCache.set(path, wrapper);
    return wrapper.elements?.find((e) => e.type === "element");
  }

  /** Write an XML part (Element → XML string). */
  public set(path: string, element: Element): void {
    const wrapper = this.wrapperCache.get(path);
    const doc: Element = wrapper
      ? { ...wrapper, elements: [{ ...element, type: "element" as const }] }
      : { elements: [{ ...element, type: "element" as const }] };
    const xml = stringify(doc);
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
        files[path] = [
          data,
          { level: levelForMediaName(path, ZIP_MEDIA_LEVEL) as ZipOptions["level"] },
        ];
      }
    }
    for (const [path, data] of this.modified) {
      files[path] = [
        data,
        { level: levelForMediaName(path, ZIP_MEDIA_LEVEL) as ZipOptions["level"] },
      ];
    }
    return zipSync(files);
  }
}

/** Parse an OOXML archive (.docx, .pptx, .xlsx) into a ParsedArchive. */
export function parseArchive(data: Uint8Array): ParsedArchive {
  return new ParsedArchive(data);
}
