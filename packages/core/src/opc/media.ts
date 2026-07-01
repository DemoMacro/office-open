/**
 * Content-deduplicated media collection for OOXML packages.
 *
 * Stores image entries keyed by file name. `addMedia` deduplicates by raw byte
 * content — byte-identical images referenced N times share one file, while
 * per-image metadata (transformation, extent, fallback) stays with each
 * reference in its own drawing XML and is passed via the `build` callback so it
 * never participates in the dedup key.
 *
 * Lookup is O(1) amortized: a `WeakMap` memoizes the resolved entry per input
 * `Uint8Array` (the hot path — a single document reuses the same buffer object
 * for repeated references), and an FNV-1a hash index resolves cross-object
 * duplicates (e.g. a freshly-decoded buffer carrying bytes already seen) with a
 * one-time byte verification to rule out hash collisions.
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
 * FNV-1a hash is a pure function of the bytes, so memoize it across
 * {@link Media} instances — repeated `generate()` calls that reuse the same
 * buffer objects (the common case) skip recomputation entirely.
 */
const hashKeyCache = new WeakMap<Uint8Array, string>();

/**
 * Shared media collection for docx/pptx/xlsx. Generic over the package's entry
 * type so each package keeps its own metadata while sharing the registration +
 * dedup logic.
 */
export class Media<T extends BaseMediaEntry> {
  private readonly map = new Map<string, T>();
  /** FNV-1a content key ("len:hash") -> canonical file name, for cross-object dedup. */
  private readonly byContent = new Map<string, string>();
  /** Memoized result per input buffer — the repeated-reference hot path. */
  private readonly verified = new WeakMap<Uint8Array, string>();
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
    // Hot path: the same buffer object referenced repeatedly resolves instantly.
    // The cached entry must match the requested type — identical bytes registered
    // under a different type (e.g. EMF bytes also carried as WMF in an
    // mc:AlternateContent fallback) are distinct resources with different
    // content-types and must NOT collapse into one part.
    const verifiedName = this.verified.get(data);
    if (verifiedName !== undefined) {
      const existing = this.map.get(verifiedName)!;
      if (existing.type === type) return existing;
    }

    // Type-scoped content key: same bytes under different types stay separate.
    const key = `${type}:${this.contentKey(data)}`;
    const existingName = this.byContent.get(key);
    if (existingName !== undefined) {
      const existing = this.map.get(existingName)!;
      // Hash hit — confirm bytes + type match to exclude a collision, then memoize.
      if (existing.type === type && this.byteEqual(existing.data, data)) {
        this.verified.set(data, existingName);
        return existing;
      }
    }

    // A pinned name may already be taken by different bytes (e.g. a
    // counter-allocated name colliding with a source basename, or two round-trip
    // paths resolving to the same name). Re-allocate rather than overwrite —
    // overwriting would silently drop the prior media's bytes.
    const resolvedName = fileName ?? this.nextName(type);
    const finalName = this.map.has(resolvedName) ? this.nextName(type) : resolvedName;
    const entry = build(finalName);
    this.map.set(finalName, entry);
    this.byContent.set(key, finalName);
    this.verified.set(data, finalName);
    return entry;
  }

  /** Next sequential `imageN.${type}` that isn't already in use. */
  private nextName(type: string): string {
    let name = `image${++this.counter}.${type}`;
    while (this.map.has(name)) name = `image${++this.counter}.${type}`;
    return name;
  }

  /** "length:hash" content key for `data` (FNV-1a over the raw bytes), memoized. */
  private contentKey(data: Uint8Array): string {
    const cached = hashKeyCache.get(data);
    if (cached !== undefined) return cached;
    let hash = 0x811c9dc5;
    for (let i = 0; i < data.length; i++) {
      hash ^= data[i];
      hash = Math.imul(hash, 0x01000193);
    }
    const key = `${data.length}:${(hash >>> 0).toString(16)}`;
    hashKeyCache.set(data, key);
    return key;
  }

  /** Structural equality with early exit on the first mismatch. */
  private byteEqual(a: Uint8Array, b: Uint8Array): boolean {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  }

  /** All registered media entries. */
  public get array(): T[] {
    return [...this.map.values()];
  }
}
