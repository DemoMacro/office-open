import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tableStylesDesc } from "./table-styles";
import type { TableStylesDescriptorOptions } from "./table-styles";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TableStylesDescriptorOptions) {
  const xml = tableStylesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return tableStylesDesc.parse(el, readCtx);
}

/** A solid color element, used verbatim by the table-style builders. */
function color(hex: string): string {
  return `<a:srgbClr val="${hex}"/>`;
}

describe("tableStylesDesc round-trip", () => {
  it("round-trips default (empty) table styles", () => {
    const opts: TableStylesDescriptorOptions = {};
    const result = roundTrip(opts);
    expect(result.opts).toBeDefined();
    expect(result.opts!.defaultStyleId).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
  });

  it("round-trips preserves def style id in XML", () => {
    const opts: TableStylesDescriptorOptions = {};
    const xml = tableStylesDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("def=");
    expect(xml).toContain("tblStyleLst");
  });

  it("round-trips a full table style: regions, text, cell, borders, fillReference", () => {
    const opts: TableStylesDescriptorOptions = {
      opts: {
        defaultStyleId: "{TEST-ID}",
        styles: [
          {
            styleId: "{TEST-ID}",
            styleName: "Test Style",
            regions: {
              wholeTbl: {
                text: {
                  bold: "on",
                  italic: "off",
                  fontReference: { collection: "minor", color: { value: "000000" } },
                },
                cell: {
                  borders: {
                    left: { width: 9525, compound: "single", color: color("333333") },
                    top: { lineReference: { index: 2 }, color: color("666666") },
                  },
                  fillReference: { index: 1, color: { value: "4472C4" } },
                },
              },
              firstRow: {
                text: { color: color("FFFFFF") },
              },
            },
          },
        ],
      },
    };
    const result = roundTrip(opts);
    expect(result.opts!.defaultStyleId).toBe("{TEST-ID}");
    expect(result.opts!.styles).toHaveLength(1);

    const s = result.opts!.styles![0];
    if (!s) throw new Error("missing style");
    expect(s.styleId).toBe("{TEST-ID}");
    expect(s.styleName).toBe("Test Style");

    // wholeTbl — text style
    const wt = s.regions!.wholeTbl!.text!;
    expect(wt.bold).toBe("on");
    expect(wt.italic).toBe("off");
    expect(wt.fontReference?.collection).toBe("minor");
    expect(wt.fontReference?.color).toEqual({ value: "000000" });

    // wholeTbl — cell style (borders + fillReference)
    const wc = s.regions!.wholeTbl!.cell!;
    expect(wc.borders?.left?.width).toBe(9525);
    expect(wc.borders?.left?.compound).toBe("single");
    expect(wc.borders?.left?.color).toContain('val="333333"');
    expect(wc.borders?.top?.lineReference?.index).toBe(2);
    expect(wc.borders?.top?.color).toContain('val="666666"');
    expect(wc.fillReference?.index).toBe(1);
    expect(wc.fillReference?.color).toEqual({ value: "4472C4" });

    // firstRow — standalone color text style
    expect(s.regions!.firstRow?.text?.color).toContain('val="FFFFFF"');
  });

  it("does not inject optional fields when absent (style without regions)", () => {
    const opts: TableStylesDescriptorOptions = {
      opts: {
        defaultStyleId: "{MIN-ID}",
        styles: [{ styleId: "{MIN-ID}", styleName: "Min" }],
      },
    };
    const result = roundTrip(opts);
    expect(result.opts!.styles).toHaveLength(1);
    const s = result.opts!.styles![0];
    if (!s) throw new Error("missing style");
    expect(s.styleId).toBe("{MIN-ID}");
    expect(s.regions).toBeUndefined();
  });
});
