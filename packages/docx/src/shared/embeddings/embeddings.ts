/**
 * Embeddings module for WordprocessingML documents.
 *
 * Manages OLE object embeddings (word/embeddings/oleObjectN.bin) referenced by
 * w:object elements. Mirrors the Media collection pattern but targets binary
 * OLE container parts instead of images.
 *
 * @module
 */

/** OLE embedding data stored under word/embeddings/. */
export interface EmbeddingData {
  /** File name within word/embeddings/ (e.g. "oleObject1.bin"). */
  fileName: string;
  /** Raw OLE container bytes. */
  data: Uint8Array;
  /** OLE program id (e.g. "Excel.Sheet.12") — informational only. */
  progId?: string;
}

/**
 * Collects OLE embeddings allocated during document generation. Each embedding
 * is stored under a sequential `oleObjectN.bin` name in word/embeddings/,
 * mirroring MS Office's numbering so output is deterministic and diffable.
 */
export class EmbeddingCollection {
  private map = new Map<string, EmbeddingData>();
  private counter = 0;

  /** Allocate the next sequential embedding file name (oleObject1.bin, …). */
  public nextEmbeddingName(): string {
    return `oleObject${++this.counter}.bin`;
  }

  /** Register an embedding under a unique key. */
  public addEmbedding(key: string, data: EmbeddingData): void {
    this.map.set(key, data);
  }

  /** All registered embeddings in insertion order. */
  public get array(): EmbeddingData[] {
    return [...this.map.values()];
  }
}
