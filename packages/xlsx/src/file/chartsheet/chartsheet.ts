/**
 * Chartsheet XML generator — produces xl/chartsheets/sheetN.xml.
 *
 * A chartsheet is a worksheet that contains only a chart (no cells).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Chartsheet
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

// ── Options ──

export interface ChartsheetOptions {
  /** Sheet name */
  readonly name?: string;
  /** Tab color (hex ARGB, e.g. "FF4472C4") */
  readonly tabColor?: string;
  /** Chart definition (type, title, series, etc.) */
  readonly chart: {
    readonly type: string;
    readonly title?: string;
    readonly categories?: readonly string[];
    readonly series: readonly {
      readonly name: string;
      readonly values: readonly number[];
    }[];
  };
}

// ── Component ──

export class Chartsheet extends BaseXmlComponent {
  private readonly opts: ChartsheetOptions;
  private drawingRId: string = "rId1";

  public constructor(options: ChartsheetOptions) {
    super("chartsheet");
    this.opts = options;
  }

  public setDrawingRId(rId: string): void {
    this.drawingRId = rId;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // sheetPr (optional)
    if (this.opts.tabColor) {
      p.push(`<sheetPr><tabColor${attrs({ rgb: this.opts.tabColor })}/></sheetPr>`);
    }

    // sheetViews (required)
    p.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');

    // drawing (required)
    p.push(`<drawing r:id="${escapeXml(this.drawingRId)}"/>`);

    p.push("</chartsheet>");
    return p.join("");
  }
}
