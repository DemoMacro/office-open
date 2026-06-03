import { escapeXml } from "@office-open/xml";

import { BaseXmlComponent } from "../xml-components/base";
import type { Context } from "../xml-components/base";

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
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument"
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
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
  | "http://schemas.microsoft.com/office/2007/relationships/media"
  // SpreadsheetML specific
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

export const TargetModeType = {
  EXTERNAL: "External",
} as const;

interface RelationshipEntry {
  readonly id: string;
  readonly type: RelationshipType;
  readonly target: string;
  readonly targetMode?: string;
}

export class Relationships extends BaseXmlComponent {
  private entries: RelationshipEntry[] = [];

  public constructor() {
    super("Relationships");
  }

  public addRelationship(
    id: number | string,
    type: RelationshipType,
    target: string,
    targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType],
  ): void {
    this.entries.push({ id: `rId${id}`, type, target, targetMode });
  }

  public get relationshipCount(): number {
    return this.entries.length;
  }

  /**
   * Zero-allocation fast path: directly concatenate XML string.
   * Bypasses the IXmlableObject intermediate tree entirely.
   */
  public override toXml(_context: Context): string {
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
