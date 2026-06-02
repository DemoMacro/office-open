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

export class WorkbookXml extends BaseXmlComponent {
  private readonly sheets: readonly SheetDefinition[];
  private readonly pivotCaches: readonly PivotCacheReference[];

  public constructor(
    sheets: readonly SheetDefinition[],
    pivotCaches?: readonly PivotCacheReference[],
  ) {
    super("workbook");
    this.sheets = sheets;
    this.pivotCaches = pivotCaches ?? [];
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
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>',
      "<sheets>",
    ];
    for (const s of this.sheets) {
      const stateAttr = s.state && s.state !== "visible" ? ` state="${s.state}"` : "";
      p.push(
        `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" r:id="${s.rId}"${stateAttr}/>`,
      );
    }
    p.push("</sheets>");

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
}
