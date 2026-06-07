/**
 * Chartsheet descriptor — produces xl/chartsheets/sheetN.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Chartsheet
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, findChild } from "@office-open/xml";

import type { ChartsheetOptions } from "../../file/chartsheet";

// ── Types ──

export interface ChartsheetDescriptorOptions extends ChartsheetOptions {
  /** Relationship ID for the drawing (set by compiler) */
  drawingRId: string;
}

// ── Descriptor ──

export const chartsheetDesc: CustomDescriptor<ChartsheetDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // sheetPr (optional)
    if (opts.tabColor || opts.published) {
      const prAttrs: string[] = [];
      if (opts.tabColor) prAttrs.push(`<tabColor${attrs({ rgb: opts.tabColor })}/>`);
      const spAttr = opts.published ? ' published="1"' : "";
      p.push(`<sheetPr${spAttr}>${prAttrs.join("")}</sheetPr>`);
    }

    // sheetViews (required)
    const svAttrs: string[] = ['workbookViewId="0"'];
    if (opts.zoomToFit) svAttrs.push('zoomToFit="1"');
    p.push(`<sheetViews><sheetView ${svAttrs.join(" ")}/></sheetViews>`);

    // sheetProtection (optional)
    if (opts.sheetProtection) {
      const sp = opts.sheetProtection;
      const spAttrs: string[] = [];
      if (sp.content) spAttrs.push(` content="1"`);
      if (sp.objects) spAttrs.push(` objects="1"`);
      if (spAttrs.length > 0) {
        p.push(`<sheetProtection${spAttrs.join("")}/>`);
      }
    }

    // pageMargins (optional)
    if (opts.pageMargins) {
      const pm = opts.pageMargins;
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
    if (opts.pageSetup) {
      const ps = opts.pageSetup;
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
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      const hfParts: string[] = [];
      if (hf.differentFirst) hfParts.push(` differentFirst="1"`);
      if (hf.differentOddEven) hfParts.push(` differentOddEven="1"`);
      const hfContent: string[] = [];
      if (hf.oddHeader) hfContent.push(`<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>`);
      if (hf.oddFooter) hfContent.push(`<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>`);
      p.push(`<headerFooter${hfParts.join("")}>${hfContent.join("")}</headerFooter>`);
    }

    // drawing (required)
    p.push(`<drawing r:id="${escapeXml(opts.drawingRId)}"/>`);

    p.push("</chartsheet>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // sheetPr
    const sheetPr = findChild(el, "sheetPr");
    if (sheetPr) {
      if (sheetPr.attributes?.["published"] === "1") result.published = true;
      const tabColor = findChild(sheetPr, "tabColor");
      if (tabColor?.attributes?.["rgb"]) result.tabColor = tabColor.attributes["rgb"];
    }

    // sheetView
    const sheetViews = findChild(el, "sheetViews");
    if (sheetViews) {
      const sv = findChild(sheetViews, "sheetView");
      if (sv?.attributes?.["zoomToFit"] === "1") result.zoomToFit = true;
    }

    return result as Record<string, unknown>;
  },
};
