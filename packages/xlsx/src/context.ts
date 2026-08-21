/**
 * XLSX compile context — write and read contexts for the descriptor pipeline.
 *
 * @module
 */

import {
  ChartCollection,
  partPathToRelsPath,
  Relationships,
  resolveRelationshipTarget,
  type RelationshipType,
} from "@office-open/core";
import type { HyperlinkTarget, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import { SharedStrings } from "@parts/shared-strings";
import { Styles } from "@parts/styles";
import type { DxfOptions } from "@parts/styles";
import type { PivotCacheReference } from "@parts/workbook";
import type { RichTextOptions } from "@parts/worksheet";
import { Media, type MediaData } from "@shared/media";

import type { XlsxDocument } from "./parse";

// ── Write Context ──

/**
 * XLSX-specific write context.
 *
 * Holds mutable state that accumulates during the compile phase:
 * shared strings, styles, media, charts, content types, and relationships.
 */
export interface HyperlinkEntry {
  key: string;
  url: string;
  tooltip?: string;
}

export class XlsxWriteContext implements WriteContext {
  sharedStrings = new SharedStrings();
  styles = new Styles();
  media = new Media<MediaData>();
  charts = new ChartCollection();
  workbookRels = new Relationships();
  pivotCacheRefs: PivotCacheReference[] = [];
  private _hyperlinks = new Map<string, HyperlinkEntry>();

  // ── WriteContext interface (core descriptor pipeline) ──

  public addRelationship(type: RelationshipType, target: string, _mode?: string): string {
    return `rId${this.workbookRels.add(type, target)}`;
  }

  public addMedia(data: Uint8Array, type: string): string {
    // Reached by DrawingML blip fills (drawing shape picture fill) via fillDesc.
    // Image anchors still register through ctx.media directly in the compiler
    // because they carry pixel dimensions; both land in the same collection.
    const entry = this.media.addMedia(data, type, (fileName) => ({
      fileName,
      type,
      data,
      width: 0,
      height: 0,
    }));
    return `{${entry.fileName}}`;
  }

  public addHyperlink(key: string, target: HyperlinkTarget): void {
    // XLSX has no slide concept; stringify is strict, so reject the slide leg
    // instead of silently dropping it.
    if (target.slide !== undefined) {
      throw new Error("xlsx text hyperlinks cannot target slides — use url instead");
    }
    this._hyperlinks.set(key, { key, url: target.url ?? "", tooltip: target.tooltip });
  }

  /**
   * Register a differential format and return its dxfId.
   */
  public registerDxf(opts: DxfOptions): number {
    return this.styles.registerDxf(opts);
  }

  /** DrawingML text hyperlinks registered by drawing shape runs (placeholder key → entry). */
  public get hyperlinks(): HyperlinkEntry[] {
    return [...this._hyperlinks.values()];
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
  /**
   * Parsed shared-string entries for resolving cell values. Rich-text entries
   * flow through as RichTextOptions objects so cells keep their structure
   * (cell.value accepts both shapes).
   */
  public readonly sharedStrings: (string | RichTextOptions)[];

  /**
   * Path of the part currently being parsed. rId numbering is per-part, so a
   * descriptor resolving a relationship mid-parse (a theme blip fill's
   * r:embed) must hit that part's own .rels, not the workbook's.
   */
  public currentPart = "xl/workbook.xml";
  private readonly partRels = new Map<string, Map<string, string>>();

  constructor(
    private xlsx: XlsxDocument,
    sharedStrings?: (string | RichTextOptions)[],
  ) {
    this.sharedStrings = sharedStrings ?? [];
    this.loadPartRels();
  }

  /** Index every `x/_rels/y.xml.rels` as y → (rId → resolved target path). */
  private loadPartRels(): void {
    for (const relsPath of this.xlsx.doc.keys()) {
      const idx = relsPath.lastIndexOf("/_rels/");
      if (idx < 0 || !relsPath.endsWith(".rels")) continue;
      const partPath = `${relsPath.slice(0, idx)}/${relsPath.slice(idx + 7, -5)}`;
      const relsEl = this.xlsx.doc.get(relsPath);
      if (!relsEl?.elements) continue;
      const byId = new Map<string, string>();
      for (const child of relsEl.elements) {
        if (child.name !== "Relationship") continue;
        const id = child.attributes?.["Id"] as string | undefined;
        const target = child.attributes?.["Target"] as string | undefined;
        if (id && target) byId.set(id, resolveRelationshipTarget(partPath, target));
      }
      this.partRels.set(partPath, byId);
    }
  }

  /**
   * Run `fn` with `currentPart` temporarily set to `partPath`, restoring the
   * previous value afterwards — a sub-part parse (the theme) resolves its
   * relationship ids against that part's own rels.
   */
  public withPart<T>(partPath: string, fn: () => T): T {
    const prev = this.currentPart;
    this.currentPart = partPath;
    try {
      return fn();
    } finally {
      this.currentPart = prev;
    }
  }

  public resolveRelationship(rId: string): string | undefined {
    const scoped = this.partRels.get(this.currentPart)?.get(rId);
    if (scoped !== undefined) return scoped;
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

  /**
   * Resolve a relationship rId from a worksheet-level rels file.
   * Worksheet rels paths: `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`
   */
  public resolveWorksheetRel(wsPath: string, rId: string): string | undefined {
    return this.partRels.get(wsPath)?.get(rId);
  }

  /**
   * Get all relationships from a worksheet rels file matching a type fragment.
   * e.g. `getWorksheetRelsByType(path, "/comments")` returns all comment relationships.
   */
  public getWorksheetRelsByType(
    wsPath: string,
    typeFragment: string,
  ): Array<{ rId: string; target: string }> {
    const relsPath = partPathToRelsPath(wsPath);
    const rels = this.xlsx.doc.get(relsPath);
    if (!rels?.elements) return [];
    const result: Array<{ rId: string; target: string }> = [];
    for (const child of rels.elements) {
      if (child.name !== "Relationship") continue;
      const type = child.attributes?.["Type"] as string | undefined;
      if (!type || !type.includes(typeFragment)) continue;
      const rId = child.attributes?.["Id"] as string | undefined;
      const target = child.attributes?.["Target"] as string | undefined;
      if (rId && target) {
        result.push({ rId, target: resolveRelationshipTarget(wsPath, target) });
      }
    }
    return result;
  }

  public getPart(path: string): Element | undefined {
    return this.xlsx.doc.get(path);
  }

  public getRaw(path: string): Uint8Array | undefined {
    return this.xlsx.doc.getRaw(path);
  }
}
