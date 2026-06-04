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

export interface ChartsheetPageMargins {
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  readonly header?: number;
  readonly footer?: number;
}

export interface ChartsheetPageSetup {
  /** Paper size (1=Letter, 9=A4, etc.) */
  readonly paperSize?: number;
  /** Orientation ("default" | "portrait" | "landscape") */
  readonly orientation?: string;
  /** Horizontal DPI */
  readonly horizontalDpi?: number;
  /** Vertical DPI */
  readonly verticalDpi?: number;
  /** Copies to print */
  readonly copies?: number;
}

export interface ChartsheetProtectionOptions {
  /** Content is protected */
  readonly content?: boolean;
  /** Objects are protected */
  readonly objects?: boolean;
}

export interface ChartsheetHeaderFooterOptions {
  /** Different first page header/footer */
  readonly differentFirst?: boolean;
  /** Different odd/even page headers/footers */
  readonly differentOddEven?: boolean;
  /** Odd page header */
  readonly oddHeader?: string;
  /** Odd page footer */
  readonly oddFooter?: string;
}

export interface ChartsheetOptions {
  /** Sheet name */
  readonly name?: string;
  /** Tab color (hex ARGB, e.g. "FF4472C4") */
  readonly tabColor?: string;
  /** Page margins */
  readonly pageMargins?: ChartsheetPageMargins;
  /** Page setup */
  readonly pageSetup?: ChartsheetPageSetup;
  /** Header/footer */
  readonly headerFooter?: ChartsheetHeaderFooterOptions;
  /** Sheet protection */
  readonly sheetProtection?: ChartsheetProtectionOptions;
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

    // sheetProtection (optional, XSD: after sheetViews, before pageMargins)
    if (this.opts.sheetProtection) {
      const sp = this.opts.sheetProtection;
      const spAttrs: string[] = [];
      if (sp.content) spAttrs.push(` content="1"`);
      if (sp.objects) spAttrs.push(` objects="1"`);
      if (spAttrs.length > 0) {
        p.push(`<sheetProtection${spAttrs.join("")}/>`);
      }
    }

    // pageMargins (optional)
    if (this.opts.pageMargins) {
      const pm = this.opts.pageMargins;
      p.push(
        `<pageMargins${attrs({
          left: pm.left ?? 0.7,
          right: pm.right ?? 0.7,
          top: pm.top ?? 0.75,
          bottom: pm.bottom ?? 0.75,
          header: pm.header ?? 0.3,
          footer: pm.footer ?? 0.3,
        })}/>`,
      );
    }

    // pageSetup (optional)
    if (this.opts.pageSetup) {
      const ps = this.opts.pageSetup;
      p.push(
        `<pageSetup${attrs({
          paperSize: ps.paperSize,
          orientation: ps.orientation,
          horizontalDpi: ps.horizontalDpi,
          verticalDpi: ps.verticalDpi,
          copies: ps.copies,
        })}/>`,
      );
    }

    // headerFooter (optional)
    if (this.opts.headerFooter) {
      const hf = this.opts.headerFooter;
      const hfParts: string[] = [];
      if (hf.differentFirst) hfParts.push(` differentFirst="1"`);
      if (hf.differentOddEven) hfParts.push(` differentOddEven="1"`);
      const hfContent: string[] = [];
      if (hf.oddHeader) hfContent.push(`<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>`);
      if (hf.oddFooter) hfContent.push(`<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>`);
      p.push(`<headerFooter${hfParts.join("")}>${hfContent.join("")}</headerFooter>`);
    }

    // drawing (required)
    p.push(`<drawing r:id="${escapeXml(this.drawingRId)}"/>`);

    p.push("</chartsheet>");
    return p.join("");
  }
}
