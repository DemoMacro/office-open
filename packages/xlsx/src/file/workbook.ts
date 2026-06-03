/**
 * Workbook component — generates xl/workbook.xml.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

export interface SheetDefinition {
  readonly name: string;
  readonly sheetId: number;
  readonly rId: string;
  readonly state?: "visible" | "hidden" | "veryHidden";
}

export interface PivotCacheReference {
  readonly cacheId: number;
  readonly rId: string;
}

export interface TablePartReference {
  readonly rId: string;
}

export interface WorkbookProtectionOptions {
  /** Lock workbook structure (add/delete/rename/move sheets) */
  readonly lockStructure?: boolean;
  /** Lock workbook windows */
  readonly lockWindows?: boolean;
  /** Lock revisions */
  readonly lockRevision?: boolean;
  /** Plain-text password — legacy Excel hash computed automatically */
  readonly workbookPassword?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  readonly workbookAlgorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  readonly workbookHashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  readonly workbookSaltValue?: string;
  /** Modern encryption: spin count */
  readonly workbookSpinCount?: number;
  /** Revisions password (legacy) */
  readonly revisionsPassword?: string;
  /** Revisions modern encryption: algorithm name */
  readonly revisionsAlgorithmName?: string;
  /** Revisions modern encryption: base64-encoded hash value */
  readonly revisionsHashValue?: string;
  /** Revisions modern encryption: base64-encoded salt value */
  readonly revisionsSaltValue?: string;
  /** Revisions modern encryption: spin count */
  readonly revisionsSpinCount?: number;
}

export class WorkbookXml extends BaseXmlComponent {
  private readonly sheets: readonly SheetDefinition[];
  private readonly pivotCaches: readonly PivotCacheReference[];
  private readonly protection?: WorkbookProtectionOptions;

  public constructor(
    sheets: readonly SheetDefinition[],
    pivotCaches?: readonly PivotCacheReference[],
    protection?: WorkbookProtectionOptions,
  ) {
    super("workbook");
    this.sheets = sheets;
    this.pivotCaches = pivotCaches ?? [];
    this.protection = protection;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x15 xr xr6 xr10 xr2"' +
        ' xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' +
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
        ' xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"' +
        ' xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"' +
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">',
      '<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="29929"/>',
      "<workbookPr/>",
    ];

    // Workbook protection (after workbookPr, before bookViews per XSD sequence)
    if (this.protection) {
      const prot = this.protection;
      const protAttrs: string[] = [];
      if (prot.lockStructure) protAttrs.push('lockStructure="1"');
      if (prot.lockWindows) protAttrs.push('lockWindows="1"');
      if (prot.lockRevision) protAttrs.push('lockRevision="1"');
      if (prot.workbookPassword)
        protAttrs.push(`workbookPassword="${this.hashPassword(prot.workbookPassword)}"`);
      if (prot.workbookAlgorithmName)
        protAttrs.push(`workbookAlgorithmName="${escapeXml(prot.workbookAlgorithmName)}"`);
      if (prot.workbookHashValue)
        protAttrs.push(`workbookHashValue="${escapeXml(prot.workbookHashValue)}"`);
      if (prot.workbookSaltValue)
        protAttrs.push(`workbookSaltValue="${escapeXml(prot.workbookSaltValue)}"`);
      if (prot.workbookSpinCount !== undefined)
        protAttrs.push(`workbookSpinCount="${prot.workbookSpinCount}"`);
      if (prot.revisionsPassword)
        protAttrs.push(`revisionsPassword="${this.hashPassword(prot.revisionsPassword)}"`);
      if (prot.revisionsAlgorithmName)
        protAttrs.push(`revisionsAlgorithmName="${escapeXml(prot.revisionsAlgorithmName)}"`);
      if (prot.revisionsHashValue)
        protAttrs.push(`revisionsHashValue="${escapeXml(prot.revisionsHashValue)}"`);
      if (prot.revisionsSaltValue)
        protAttrs.push(`revisionsSaltValue="${escapeXml(prot.revisionsSaltValue)}"`);
      if (prot.revisionsSpinCount !== undefined)
        protAttrs.push(`revisionsSpinCount="${prot.revisionsSpinCount}"`);
      if (protAttrs.length > 0) {
        p.push(`<workbookProtection ${protAttrs.join(" ")}/>`);
      }
    }

    p.push(
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>',
      "<sheets>",
    );
    for (const s of this.sheets) {
      const stateAttr = s.state && s.state !== "visible" ? ` state="${s.state}"` : "";
      p.push(
        `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" r:id="${s.rId}"${stateAttr}/>`,
      );
    }
    p.push("</sheets>");

    // externalReferences placeholder — compiler injects the XML here if needed
    p.push("<!--EXTERNAL_REFS-->");

    p.push('<calcPr calcId="162913"/>');

    if (this.pivotCaches.length > 0) {
      p.push("<pivotCaches>");
      for (const pc of this.pivotCaches) {
        p.push(`<pivotCache cacheId="${pc.cacheId}" r:id="${pc.rId}"/>`);
      }
      p.push("</pivotCaches>");
    }

    p.push("</workbook>");
    return p.join("");
  }

  /**
   * Generate tableParts XML fragment for embedding in a worksheet.
   * This is called by the compiler to insert table references into the worksheet XML.
   */
  public static buildTablePartsXml(tableParts: readonly TablePartReference[]): string {
    if (tableParts.length === 0) return "";
    const parts: string[] = [`<tableParts count="${tableParts.length}">`];
    for (const tp of tableParts) {
      parts.push(`<tablePart r:id="${tp.rId}"/>`);
    }
    parts.push("</tableParts>");
    return parts.join("");
  }

  /**
   * Generate externalReferences XML fragment for embedding in the workbook.
   * This is called by the compiler to insert external reference entries.
   */
  public static buildExternalReferencesXml(refs: readonly { rId: string }[]): string {
    if (refs.length === 0) return "";
    const parts: string[] = ["<externalReferences>"];
    for (const ref of refs) {
      parts.push(`<externalReference r:id="${ref.rId}"/>`);
    }
    parts.push("</externalReferences>");
    return parts.join("");
  }

  /** Legacy Excel password hash (XOR-based) */
  private hashPassword(password: string): string {
    let hash = 0;
    for (let i = 0; i < password.length; i++) {
      const c = password.charCodeAt(i);
      hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
      hash ^= c;
      hash = hash & 0x4000 ? hash ^ 0x1 : hash;
    }
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash ^= password.length;
    return hash.toString(16).toUpperCase().padStart(4, "0");
  }
}
