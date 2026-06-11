/**
 * XLSX compile context — write and read contexts for the descriptor pipeline.
 *
 * @module
 */

import { ChartCollection, Relationships, type RelationshipType } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import { ContentTypes } from "@parts/content-types";
import { Media } from "@parts/media";
import { SharedStrings } from "@parts/shared-strings";
import { Styles } from "@parts/styles";
import type { DxfOptions } from "@parts/styles";
import type { PivotCacheReference } from "@parts/workbook";

import type { XlsxDocument } from "./parse";

// ── Write Context ──

/**
 * XLSX-specific write context.
 *
 * Holds mutable state that accumulates during the compile phase:
 * shared strings, styles, media, charts, content types, and relationships.
 */
export class XlsxWriteContext implements WriteContext {
  sharedStrings = new SharedStrings();
  styles = new Styles();
  media = new Media();
  charts = new ChartCollection();
  contentTypes = new ContentTypes();
  workbookRels = new Relationships();
  pivotCacheRefs: PivotCacheReference[] = [];

  // ── WriteContext stubs (core interface) ──

  public addRelationship(type: RelationshipType, target: string, _mode?: string): string {
    const id = this.workbookRels.relationshipCount + 1;
    this.workbookRels.addRelationship(id, type, target);
    return `rId${id}`;
  }

  public addMedia(_data: Uint8Array, _type: string): string {
    // Stub — XLSX media registration goes through Media.addImage() in compiler.
    return "";
  }

  /**
   * Register a differential format and return its dxfId.
   */
  public registerDxf(opts: DxfOptions): number {
    return this.styles.registerDxf(opts);
  }
}

// ── Read Context ──

/**
 * XLSX-specific read context.
 *
 * Wraps an {@link XlsxDocument} to implement the core {@link ReadContext}
 * interface used by the descriptor parse pipeline.
 */
export class XlsxReadContext implements ReadContext {
  /** Parsed shared strings for resolving cell values. */
  public readonly sharedStrings: string[];

  constructor(
    private xlsx: XlsxDocument,
    sharedStrings?: string[],
  ) {
    this.sharedStrings = sharedStrings ?? [];
  }

  public resolveRelationship(rId: string): string | undefined {
    const wbRels = this.xlsx.doc.get("xl/_rels/workbook.xml.rels");
    if (!wbRels?.elements) return undefined;
    for (const child of wbRels.elements) {
      if (child.name !== "Relationship") continue;
      if (child.attributes?.["Id"] === rId) {
        const target = child.attributes["Target"] as string | undefined;
        if (!target) return undefined;
        return target.startsWith("/") ? target.slice(1) : `xl/${target}`;
      }
    }
    return undefined;
  }

  public getPart(path: string): Element | undefined {
    return this.xlsx.doc.get(path);
  }

  public getRaw(path: string): Uint8Array | undefined {
    return this.xlsx.doc.getRaw(path);
  }
}
