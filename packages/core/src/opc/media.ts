/**
 * Content-deduplicated media collection for OOXML packages.
 *
 * Stores image entries keyed by file name. `addMedia` deduplicates by raw byte
 * content — byte-identical images referenced N times share one file (matching
 * Office's normalized output), while per-image metadata (transformation,
 * extent, fallback) stays with each reference in its own drawing XML and is
 * passed via the `build` callback so it never participates in the dedup key.
 *
 * @module
 */

/**
 * Minimum fields every media entry carries. Package-specific entry types
 * extend this with their own transformation/extent/fallback fields.
 */
export interface BaseMediaEntry {
  fileName: string;
  data: Uint8Array;
  type: string;
}

/**
 * Shared media collection for docx/pptx/xlsx. Generic over the package's entry
 * type so each package keeps its own metadata while sharing the registration +
 * dedup logic.
 */
export class Media<T extends BaseMediaEntry> {
  private readonly map = new Map<string, T>();
  private counter = 0;

  /**
   * Register media, reusing the existing entry when the bytes already exist.
   * Returns the canonical entry (shared across all identical references) —
   * callers read `entry.fileName` for placeholders/relationship targets or use
   * the whole entry (e.g. drawing XML). The `build` callback constructs the
   * package-specific entry from the allocated file name.
   *
   * Pass `fileName` to pin the name (round-trip scenarios preserving a source
   * file name); omit it to allocate the next sequential `imageN.ext`.
   */
  public addMedia(
    data: Uint8Array,
    type: string,
    build: (fileName: string) => T,
    fileName?: string,
  ): T {
    const existingKey = this.findByContent(data);
    if (existingKey) return this.map.get(existingKey)!;
    const resolvedName = fileName ?? `image${++this.counter}.${type}`;
    const entry = build(resolvedName);
    this.map.set(resolvedName, entry);
    return entry;
  }

  /** Find the key of an existing entry with byte-identical content. */
  private findByContent(data: Uint8Array): string | undefined {
    for (const [key, entry] of this.map) {
      const existing = entry.data;
      if (existing.length !== data.length) continue;
      let match = true;
      for (let i = 0; i < existing.length; i++) {
        if (existing[i] !== data[i]) {
          match = false;
          break;
        }
      }
      if (match) return key;
    }
    return undefined;
  }

  /** All registered media entries. */
  public get array(): T[] {
    return [...this.map.values()];
  }
}
