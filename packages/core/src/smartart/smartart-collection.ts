/**
 * SmartArt data and collection for document generation.
 *
 * @module
 */

import type { ColorDefinitionOptions } from "./color-definition";
import type { LayoutDefinitionOptions } from "./layout-definition";
import type { StyleDefinitionOptions } from "./style-definition";

export interface SmartArtData {
  key: string;
  dataModelXml: string;
  /** Built-in layout id ("process1") or a full custom layout definition. */
  layout: string | LayoutDefinitionOptions;
  /** Built-in quick-style id ("simple1") or a full custom style definition. */
  style: string | StyleDefinitionOptions;
  /** Built-in color-transform id ("accent1_2") or a full custom definition. */
  color: string | ColorDefinitionOptions;
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

  public constructor() {
    this.map = new Map<string, SmartArtData>();
  }

  public addSmartArt(key: string, data: SmartArtData): void {
    this.map.set(key, data);
  }

  public get array(): SmartArtData[] {
    return [...this.map.values()];
  }
}
