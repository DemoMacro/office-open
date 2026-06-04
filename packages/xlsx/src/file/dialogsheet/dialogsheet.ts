/**
 * Dialogsheet XML generator — produces xl/dialogsheets/sheetN.xml.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs } from "@office-open/xml";

// ── Options ──

export interface DialogsheetPageMargins {
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  readonly header?: number;
  readonly footer?: number;
}

export interface DialogsheetPageSetup {
  readonly paperSize?: number;
  readonly orientation?: string;
  readonly horizontalDpi?: number;
  readonly verticalDpi?: number;
  readonly copies?: number;
}

export interface DialogsheetProtectionOptions {
  readonly content?: boolean;
  readonly objects?: boolean;
  readonly scenarios?: boolean;
}

export interface DialogsheetOptions {
  /** Sheet name */
  readonly name?: string;
  /** Tab color (hex ARGB) */
  readonly tabColor?: string;
  /** Page margins */
  readonly pageMargins?: DialogsheetPageMargins;
  /** Page setup */
  readonly pageSetup?: DialogsheetPageSetup;
  /** Sheet protection */
  readonly sheetProtection?: DialogsheetProtectionOptions;
}

// ── Component ──

export class Dialogsheet extends BaseXmlComponent {
  private readonly opts: DialogsheetOptions;

  public constructor(options: DialogsheetOptions) {
    super("dialogsheet");
    this.opts = options;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // sheetPr (optional)
    if (this.opts.tabColor) {
      p.push(`<sheetPr><tabColor${attrs({ rgb: this.opts.tabColor })}/></sheetPr>`);
    }

    // sheetViews (required for most implementations)
    p.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');

    // sheetProtection (optional)
    if (this.opts.sheetProtection) {
      const sp = this.opts.sheetProtection;
      const spAttrs: Record<string, string | number | undefined> = {};
      if (sp.content) spAttrs.sheet = "1";
      if (sp.objects) spAttrs.objects = "1";
      if (sp.scenarios) spAttrs.scenarios = "1";
      if (Object.keys(spAttrs).length > 0) {
        p.push(`<sheetProtection${attrs(spAttrs)}/>`);
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

    p.push("</dialogsheet>");
    return p.join("");
  }
}
