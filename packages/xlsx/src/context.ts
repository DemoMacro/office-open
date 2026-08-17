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
import { Styles, builtinNumFmtCode } from "@parts/styles";
import type { DxfOptions, StyleOptions, StylesParseResult } from "@parts/styles";
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
  /** Parsed styles (fonts, fills, borders, cellXfs). Set by parseWorkbook(). */
  public parsedStyles?: StylesParseResult;

  constructor(
    private xlsx: XlsxDocument,
    sharedStrings?: (string | RichTextOptions)[],
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

  /**
   * Resolve a relationship rId from a worksheet-level rels file.
   * Worksheet rels paths: `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`
   */
  public resolveWorksheetRel(wsPath: string, rId: string): string | undefined {
    const relsPath = partPathToRelsPath(wsPath);
    const rels = this.xlsx.doc.get(relsPath);
    if (!rels?.elements) return undefined;
    for (const child of rels.elements) {
      if (child.name !== "Relationship") continue;
      if (child.attributes?.["Id"] === rId) {
        const target = child.attributes["Target"] as string | undefined;
        if (!target) return undefined;
        return resolveRelationshipTarget(wsPath, target);
      }
    }
    return undefined;
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

  /**
   * Resolve a cell style index to a StyleOptions object by looking up
   * the parsed cellXfs table and substituting font/fill/border/numFmt indices
   * with their resolved values.
   */
  public resolveStyle(styleIndex: number): StyleOptions | undefined {
    const ps = this.parsedStyles;
    if (!ps) return undefined;
    const { cellXfs, fonts, fills, borders, customNumFmtById } = ps;
    if (!cellXfs || styleIndex >= cellXfs.length) return undefined;
    const xf = cellXfs[styleIndex];
    if (!xf) return undefined;
    const result: StyleOptions = {};

    const fontId = xf.fontId;
    if (fontId !== undefined && fonts && fontId < fonts.length) result.font = fonts[fontId];
    const fillId = xf.fillId;
    if (fillId !== undefined && fills && fillId < fills.length) result.fill = fills[fillId];
    const borderId = xf.borderId;
    if (borderId !== undefined && borders && borderId < borders.length)
      result.border = borders[borderId];
    const numFmtId = xf.numFmtId;
    if (numFmtId !== undefined) {
      // Custom <numFmts> entries win; built-in ids (0-49) resolve through the
      // builtin table so date/percent cells keep their format on round-trip.
      const code = customNumFmtById?.get(numFmtId) ?? builtinNumFmtCode(numFmtId);
      if (code !== undefined) result.numFmt = code;
    }
    if (xf.alignment) result.alignment = xf.alignment;
    if (xf.protection) result.protection = xf.protection;
    if (xf.quotePrefix) result.quotePrefix = xf.quotePrefix;
    if (xf.pivotButton) result.pivotButton = xf.pivotButton;

    return result as StyleOptions;
  }
}
