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
import type { DxfOptions, StyleOptions, StylesParseResult } from "@parts/styles";
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
  /** Parsed styles (fonts, fills, borders, cellXfs). Set by parseWorkbook(). */
  public parsedStyles?: StylesParseResult;

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

  /**
   * Resolve a relationship rId from a worksheet-level rels file.
   * Worksheet rels paths: `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`
   */
  public resolveWorksheetRel(wsPath: string, rId: string): string | undefined {
    const relsPath = wsPathToRelsPath(wsPath);
    const rels = this.xlsx.doc.get(relsPath);
    if (!rels?.elements) return undefined;
    for (const child of rels.elements) {
      if (child.name !== "Relationship") continue;
      if (child.attributes?.["Id"] === rId) {
        const target = child.attributes["Target"] as string | undefined;
        if (!target) return undefined;
        return resolveWsTarget(wsPath, target);
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
    const relsPath = wsPathToRelsPath(wsPath);
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
        result.push({ rId, target: resolveWsTarget(wsPath, target) });
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
    const { cellXfs, fonts, fills, borders, customNumFmts } = ps;
    if (!cellXfs || styleIndex >= cellXfs.length) return undefined;
    const xf = cellXfs[styleIndex];
    const result: StyleOptions = {};

    const fontId = xf.fontId;
    if (fontId !== undefined && fonts && fontId < fonts.length) result.font = fonts[fontId];
    const fillId = xf.fillId;
    if (fillId !== undefined && fills && fillId < fills.length) result.fill = fills[fillId];
    const borderId = xf.borderId;
    if (borderId !== undefined && borders && borderId < borders.length)
      result.border = borders[borderId];
    const numFmtId = xf.numFmtId;
    if (numFmtId !== undefined && customNumFmts) {
      for (const [code, id] of Object.entries(customNumFmts)) {
        if (id === numFmtId) {
          result.numFmt = code;
          break;
        }
      }
    }
    if (xf.alignment) result.alignment = xf.alignment;
    if (xf.protection) result.protection = xf.protection;
    if (xf.quotePrefix) result.quotePrefix = xf.quotePrefix;
    if (xf.pivotButton) result.pivotButton = xf.pivotButton;

    return result as StyleOptions;
  }
}

// ── Worksheet rels helpers ──

/** Derive rels path from worksheet path. */
function wsPathToRelsPath(wsPath: string): string {
  // "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
  const idx = wsPath.lastIndexOf("/");
  const dir = wsPath.substring(0, idx);
  const file = wsPath.substring(idx + 1);
  return `${dir}/_rels/${file}.rels`;
}

/** Resolve a relative target from a worksheet rels file to an absolute archive path. */
function resolveWsTarget(wsPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  // Target is relative to the worksheet's directory, e.g. "../comments1.xml" from "xl/worksheets/"
  const wsDir = wsPath.substring(0, wsPath.lastIndexOf("/"));
  const parts = target.split("/");
  const dirParts = wsDir.split("/");
  for (const part of parts) {
    if (part === "..") {
      dirParts.pop();
    } else {
      dirParts.push(part);
    }
  }
  return dirParts.join("/");
}
