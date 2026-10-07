/**
 * SmartArt data and collection for document generation.
 *
 * @module
 */

import type { DataType } from "../util/data-type";
import type { ColorDefinitionOptions } from "./color-definition";
import type { LayoutDefinitionOptions } from "./layout-definition";
import type { StyleDefinitionOptions } from "./style-definition";

/**
 * Verbatim source parts for byte-exact SmartArt round-trip. When set on a
 * `SmartArtData` entry, the compiler re-emits these bytes under the
 * diagram part names instead of rebuilding from the modeled XML.
 */
export interface SmartArtRawParts {
  /** diagrams/dataN.xml source bytes. */
  data?: DataType;
  /** diagrams/layoutN.xml source bytes. */
  layout?: DataType;
  /** diagrams/quickStyleN.xml source bytes. */
  style?: DataType;
  /** diagrams/colorsN.xml source bytes. */
  color?: DataType;
  /**
   * diagrams/drawingN.xml source bytes — the pre-rendered dsp:drawing
   * snapshot (MS-ODRAWXML 2008 extension) cached beside the data model.
   * Falls back to an empty spTree shell when absent.
   */
  drawing?: DataType;
  /** Source drawing part file name (for sparse packages such as drawing10.xml). */
  drawingFileName?: string;
  /** Images referenced by the data part's own rels (dgm:pt blipFill art). */
  media?: { fileName: string; data: DataType }[];
  /** Verbatim rels XML of the data part (its targets resolve beside the data part). */
  dataRels?: DataType;
}

/** Source package paths preserved by round-trip; compiler-only. */
export interface SmartArtSourcePaths {
  data?: string;
  layout?: string;
  quickStyle?: string;
  colors?: string;
  drawing?: string;
  dataRels?: string;
}

export interface SmartArtData {
  key: string;
  dataModelXml: string;
  /** Built-in layout id ("process1") or a full custom layout definition. */
  layout: string | LayoutDefinitionOptions;
  /** Built-in quick-style id ("simple1") or a full custom style definition. */
  style: string | StyleDefinitionOptions;
  /** Built-in color-transform id ("accent1_2") or a full custom definition. */
  color: string | ColorDefinitionOptions;
  /**
   * Round-trip verbatim parts; emitted in place of the modeled XML. The
   * modeled fields stay populated so the Options remain readable and dropping
   * raw rebuilds from them.
   */
  raw?: SmartArtRawParts;
  /** Source part locations for sparse paths; round-trip only. */
  sourcePaths?: SmartArtSourcePaths;
}

/**
 * Last URN segment of a custom definition's uniqueId — the id embedded in the
 * data model's doc-point type ids. Falls back to "custom" when unset.
 */
export function definitionId(definition: { uniqueId?: string }): string {
  return definition.uniqueId?.split("/").pop() || "custom";
}

/**
 * Manages SmartArt parts in a document.
 */
export class SmartArtCollection {
  private map: Map<string, SmartArtData>;
  /** Cached `array` snapshot — invalidated on add. */
  private cachedArray: SmartArtData[] | undefined;

  public constructor() {
    this.map = new Map<string, SmartArtData>();
  }

  public addSmartArt(key: string, data: SmartArtData): void {
    this.map.set(key, data);
    this.cachedArray = undefined;
  }

  public get array(): SmartArtData[] {
    if (this.cachedArray === undefined) this.cachedArray = [...this.map.values()];
    return this.cachedArray;
  }
}
