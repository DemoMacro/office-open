import { escapeXml } from "@office-open/xml";

import type { XmlifyedFile } from "./packer";

export type RelationshipType =
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  | "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
  | "http://schemas.microsoft.com/office/2007/relationships/diagramLayout"
  | "http://schemas.microsoft.com/office/2007/relationships/diagramStyle"
  | "http://schemas.microsoft.com/office/2007/relationships/diagramColors"
  | "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
  // WordprocessingML specific
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  | "http://schemas.microsoft.com/office/2011/relationships/people"
  | "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
  // Native-format embedding (embedded xlsx/docx behind OLE objects)
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
  | "http://schemas.microsoft.com/office/2006/relationships/txbx"
  // PresentationML specific
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncProperties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeManager"
  | "http://schemas.microsoft.com/office/2007/relationships/media"
  // SpreadsheetML specific
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/users"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";

export const TargetModeType = {
  EXTERNAL: "External",
} as const;

interface RelationshipEntry {
  id: string;
  type: RelationshipType;
  target: string;
  targetMode?: string;
}

/**
 * Manages OOXML relationship entries and serializes to XML.
 *
 * Standalone class — no XmlComponent inheritance.
 * Pure string concatenation for zero-allocation XML output.
 */
export class Relationships {
  private entries: RelationshipEntry[] = [];

  public addRelationship(
    id: number | string,
    type: RelationshipType,
    target: string,
    targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType],
  ): void {
    this.entries.push({ id: `rId${id}`, type, target, targetMode });
  }

  /**
   * Register a relationship with an auto-allocated sequential id and return
   * the numeric id. Prefer this over the `relationshipCount + 1` +
   * `addRelationship` pair at every call site that wants the next id; reach
   * for `addRelationship` directly only when the id is externally determined
   * (a contiguous batch pre-computed from an offset, a fixed rId1, …).
   */
  public add(
    type: RelationshipType,
    target: string,
    targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType],
  ): number {
    const id = this.entries.length + 1;
    this.entries.push({ id: `rId${id}`, type, target, targetMode });
    return id;
  }

  public get relationshipCount(): number {
    return this.entries.length;
  }

  /**
   * True when a relationship of the same kind and target is already
   * registered. The kind compares only the relationship type's last path
   * segment, so the transitional and strict URI forms of the same
   * relationship compare equal — passthrough re-emission uses this to skip
   * relationships the compiler already wired (a duplicate entry corrupts the
   * package for Office applications).
   */
  public hasRelationship(type: string, target: string): boolean {
    const kind = type.split("/").pop();
    return this.entries.some((e) => e.type.split("/").pop() === kind && e.target === target);
  }

  /**
   * True when any entry carries the relationship kind (last type segment).
   * Round-trip re-emission uses this as the ownership test: a rebuilt part
   * that already registered a rel of this kind absorbed the concept, so the
   * captured source rel (possibly a renamed target or the ISO-strict type
   * dual) must not be re-emitted alongside it.
   */
  public hasRelationshipKind(kind: string): boolean {
    return this.entries.some((e) => e.type.split("/").pop() === kind);
  }

  /** True when an entry already occupies this exact relationship id. */
  public hasId(id: string): boolean {
    return this.entries.some((e) => e.id === id);
  }

  /**
   * The id of the entry matching kind+target (same last-path-segment equality
   * as {@link hasRelationship}), or undefined when absent. Round-trip legs use
   * this to point a reference element at an already-registered relationship.
   */
  public idOf(type: string, target: string): string | undefined {
    const kind = type.split("/").pop();
    return this.entries.find((e) => e.type.split("/").pop() === kind && e.target === target)?.id;
  }

  /** Directly builds XML string — zero intermediate tree allocation. */
  public serialize(): string {
    const p: string[] = [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ];
    for (const e of this.entries) {
      const tm = e.targetMode ? ` TargetMode="${escapeXml(e.targetMode)}"` : "";
      p.push(
        `<Relationship Id="${escapeXml(e.id)}" Type="${escapeXml(e.type)}" Target="${escapeXml(e.target)}"${tm}/>`,
      );
    }
    p.push("</Relationships>");
    return p.join("");
  }
}

/**
 * Serialize a Relationships part only when it carries at least one
 * relationship. Optional parts (fontTable, headers, footers, charts, drawings,
 * worksheets, …) emit no .rels part when empty — Office strips empty rels
 * shells when re-saving, so skipping them keeps generated packages free of
 * redundant empty parts and matches Office's normalized output.
 *
 * Always-on parts (the package `_rels/.rels` and the main
 * document/presentation/workbook parts) carry relationships by construction
 * and must NOT use this gate.
 */
export function optionalRelsPart(
  rel: Relationships,
  xmlDeclaration: string,
  path: string,
): XmlifyedFile | undefined {
  return rel.relationshipCount > 0 ? { data: xmlDeclaration + rel.serialize(), path } : undefined;
}

/** Build the package root relationships shared by docx, pptx, and xlsx. */
export function buildRootRelationships(
  mainPartTarget: string,
  includeCustomProperties: boolean,
  passthroughRelationships?: readonly {
    source: string;
    relationshipType: string;
    target: string;
    targetMode?: "External";
  }[],
): Relationships {
  const rels = new Relationships();
  rels.addRelationship(
    1,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    mainPartTarget,
  );
  rels.addRelationship(
    2,
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "docProps/core.xml",
  );
  rels.addRelationship(
    3,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "docProps/app.xml",
  );
  if (includeCustomProperties) {
    rels.addRelationship(
      4,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
      "docProps/custom.xml",
    );
  }
  // Root-level passthrough relationships (round-trip): the source _rels/.rels
  // referenced parts the model carries verbatim — the package thumbnail above
  // all. Re-emitted as written; targets never move.
  for (const rel of passthroughRelationships ?? []) {
    if (rel.source !== "") continue;
    if (rels.hasRelationship(rel.relationshipType, rel.target)) continue;
    rels.add(rel.relationshipType as RelationshipType, rel.target, rel.targetMode);
  }
  return rels;
}

/**
 * Derive the .rels part path for a package part:
 * "ppt/slides/slide1.xml" → "ppt/slides/_rels/slide1.xml.rels".
 */
export function partPathToRelsPath(partPath: string): string {
  const idx = partPath.lastIndexOf("/");
  const dir = partPath.substring(0, idx);
  const file = partPath.substring(idx + 1);
  return `${dir}/_rels/${file}.rels`;
}

/**
 * Resolve a relationship target against the referencing part's directory,
 * segment by segment: each ".." pops one directory level. A single
 * String.replace("../", …) only strips the first occurrence and mis-resolves
 * deeper targets like "../../media/image.png".
 */
export function resolveRelationshipTarget(partPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const dir = partPath.substring(0, partPath.lastIndexOf("/"));
  const dirParts = dir ? dir.split("/") : [];
  for (const part of target.split("/")) {
    if (part === "..") {
      dirParts.pop();
    } else if (part !== "." && part !== "") {
      dirParts.push(part);
    }
  }
  return dirParts.join("/");
}
