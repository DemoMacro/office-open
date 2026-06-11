/**
 * Sub-document collection module for managing sub-document parts.
 *
 * @module
 */

/**
 * Stores sub-document data for later serialization by the compiler.
 */
export interface SubDocData {
  /** Raw document data (.docx bytes) */
  data: Uint8Array;
  /** Part sub-path within word/ (e.g., "subdocs/subdoc1.docx") */
  path: string;
}

/**
 * Manages sub-document parts in a document.
 */
export class SubDocCollection {
  private map: Map<string, SubDocData>;

  public constructor() {
    this.map = new Map<string, SubDocData>();
  }

  public addSubDoc(key: string, data: SubDocData): void {
    this.map.set(key, data);
  }

  public get array(): SubDocData[] {
    return [...this.map.values()];
  }
}
