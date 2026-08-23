/**
 * OLE embedding collection shared by the format packages.
 *
 * Manages OLE object embeddings (word/embeddings/oleObjectN.bin,
 * ppt/embeddings/oleObjectN.bin) referenced by OLE object elements. Mirrors
 * the Media collection pattern but targets binary OLE container parts.
 *
 * @module
 */

/** OLE embedding data stored under the package's embeddings/ directory. */
export interface EmbeddingData {
  /** File name within embeddings/ (e.g. "oleObject1.bin"). */
  fileName: string;
  /** Raw OLE container bytes. */
  data: Uint8Array;
  /** OLE program id (e.g. "Excel.Sheet.12") — informational only. */
  progId?: string;
  /**
   * Source relationship kind (round-trip): "package" for native-format
   * embeddings (xlsx/docx), "oleObject" for OLE compound binaries. Controls
   * the re-emitted relationship type; undefined (fresh authoring) means
   * oleObject.
   */
  relationshipType?: "oleObject" | "package";
}

/**
 * Collects OLE embeddings allocated during document generation. Each embedding
 * is stored under a sequential `oleObjectN.bin` name, mirroring MS Office's
 * numbering so output is deterministic and diffable.
 */
export class EmbeddingCollection {
  private map = new Map<string, EmbeddingData>();
  private counter = 0;
  /** Cached `array` snapshot — invalidated on add. */
  private cachedArray: EmbeddingData[] | undefined;

  /** Allocate the next sequential embedding file name (oleObject1.bin, …). */
  public nextEmbeddingName(): string {
    // Skip names already taken (e.g. round-tripped source basenames) — a
    // counter-allocated name colliding with a pinned one would overwrite it.
    let name = `oleObject${++this.counter}.bin`;
    while (this.map.has(name)) name = `oleObject${++this.counter}.bin`;
    return name;
  }

  /**
   * Register embedding bytes and return the stored entry. The requested name
   * is used when free or already claimed by identical bytes (two objects
   * sharing one embedded part); a name claimed by different bytes falls back
   * to a fresh sequential name — overwriting would silently drop the other
   * object's payload.
   */
  public addEmbedding(
    data: Uint8Array,
    requestedName?: string,
    progId?: string,
    relationshipType?: EmbeddingData["relationshipType"],
  ): EmbeddingData {
    const requested = requestedName ?? this.nextEmbeddingName();
    const extras = {
      ...(progId !== undefined ? { progId } : {}),
      ...(relationshipType !== undefined ? { relationshipType } : {}),
    };
    const existing = this.map.get(requested);
    if (existing) {
      if (this.byteEqual(existing.data, data)) return existing;
      const fallbackName = this.nextEmbeddingName();
      const entry: EmbeddingData = { fileName: fallbackName, data, ...extras };
      this.map.set(fallbackName, entry);
      this.cachedArray = undefined;
      return entry;
    }
    const entry: EmbeddingData = { fileName: requested, data, ...extras };
    this.map.set(requested, entry);
    this.cachedArray = undefined;
    return entry;
  }

  /** All registered embeddings in insertion order (snapshot; stable between adds). */
  public get array(): EmbeddingData[] {
    if (this.cachedArray === undefined) this.cachedArray = [...this.map.values()];
    return this.cachedArray;
  }

  private byteEqual(a: Uint8Array, b: Uint8Array): boolean {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }
}
