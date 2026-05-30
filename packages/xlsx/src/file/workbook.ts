/**
 * Workbook component — generates xl/workbook.xml.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

export interface SheetDefinition {
  readonly name: string;
  readonly sheetId: number;
  readonly rId: string;
  readonly state?: "visible" | "hidden" | "veryHidden";
}

export class WorkbookXml extends BaseXmlComponent {
  private readonly sheets: readonly SheetDefinition[];

  public constructor(sheets: readonly SheetDefinition[]) {
    super("workbook");
    this.sheets = sheets;
  }

  public override prepForXml(_context: Context): IXmlableObject {
    const sheetElements: IXmlableObject[] = [];
    for (const s of this.sheets) {
      const attrs: Record<string, string> = {
        name: s.name,
        sheetId: String(s.sheetId),
        "r:id": s.rId,
      };
      if (s.state && s.state !== "visible") {
        attrs.state = s.state;
      }
      sheetElements.push({ sheet: { _attr: attrs } });
    }

    return {
      workbook: [
        {
          _attr: {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          },
        },
        {
          bookViews: [
            {
              workbookView: {
                _attr: { xWindow: 0, yWindow: 0, windowWidth: 28800, windowHeight: 12300 },
              },
            },
          ],
        },
        { sheets: sheetElements },
        { calcPr: { _attr: { calcId: 191029 } } },
      ],
    };
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>',
      "<sheets>",
    ];
    for (const s of this.sheets) {
      const stateAttr = s.state && s.state !== "visible" ? ` state="${s.state}"` : "";
      p.push(
        `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" r:id="${s.rId}"${stateAttr}/>`,
      );
    }
    p.push('</sheets><calcPr calcId="191029"/></workbook>');
    return p.join("");
  }
}
