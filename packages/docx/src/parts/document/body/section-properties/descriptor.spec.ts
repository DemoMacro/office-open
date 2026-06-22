import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../../../context";
import { sectionPropertiesDesc } from "./descriptor";
import type { SectionPropertiesOptions } from "./section-properties";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) => String(child),
  fileData: {} as never,
} as unknown as BodyContext;

const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: SectionPropertiesOptions) {
  const xml = sectionPropertiesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return sectionPropertiesDesc.parse(el, readCtx);
}

describe("sectionPropertiesDesc round-trip", () => {
  it("round-trips default section (empty opts)", () => {
    const result = roundTrip({});
    // Grid defaults are always present
    expect(result.grid).toBeDefined();
    expect(result.grid!.linePitch).toBe(312);
  });

  it("round-trips page size", () => {
    const result = roundTrip({
      page: { size: { width: 12240, height: 15840 } },
    });
    expect(result.page).toBeDefined();
    expect(result.page!.size!.width).toBe(12240);
    expect(result.page!.size!.height).toBe(15840);
  });

  it("round-trips landscape orientation (swaps w/h and swaps back)", () => {
    const result = roundTrip({
      page: { size: { width: 12240, height: 15840, orientation: "landscape" } },
    });
    // Logical width/height (portrait perspective) must survive the stringify
    // swap (w:w=height, w:h=width) and the parse swap-back.
    expect(result.page!.size!.orientation).toBe("landscape");
    expect(result.page!.size!.width).toBe(12240);
    expect(result.page!.size!.height).toBe(15840);
  });

  it("parses a Word-emitted landscape page size (physical w > h)", () => {
    // Word stores landscape with the long edge in w:w (physical), e.g. Letter
    // landscape: w:w=15840 w:h=12240 orient=landscape. Parse must swap back to
    // logical width=12240 (short) height=15840 (long).
    const xml =
      '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>';
    const el = parseXml(xml).elements![0];
    const result = sectionPropertiesDesc.parse(el, readCtx);
    expect(result.page!.size!.orientation).toBe("landscape");
    expect(result.page!.size!.width).toBe(12240);
    expect(result.page!.size!.height).toBe(15840);
  });

  it("parses portrait page size without swapping (w = logical width)", () => {
    const xml =
      '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
    const el = parseXml(xml).elements![0];
    const result = sectionPropertiesDesc.parse(el, readCtx);
    expect(result.page!.size!.width).toBe(12240);
    expect(result.page!.size!.height).toBe(15840);
    expect(result.page!.size!.orientation).toBeUndefined();
  });

  it("round-trips portrait page size with UniversalMeasure (mm → twips)", () => {
    // UniversalMeasure on width/height is normalized to twips on stringify (so
    // the attrNum-based parse reads it back). 210/297mm floor to 11905/16837
    // twips (convertMillimetersToTwip uses Math.floor).
    const result = roundTrip({
      page: { size: { width: "210mm", height: "297mm" } },
    });
    expect(result.page!.size!.width).toBe(11905);
    expect(result.page!.size!.height).toBe(16837);
    // orientation defaults to portrait when omitted.
    expect(result.page!.size!.orientation).toBe("portrait");
  });

  it("round-trips page size code (printer paper code)", () => {
    const result = roundTrip({
      page: { size: { width: 12240, height: 15840, code: 1 } },
    });
    expect(result.page!.size!.code).toBe(1);
  });

  it("round-trips page margins", () => {
    const result = roundTrip({
      page: {
        margin: {
          top: 1440,
          right: 1440,
          bottom: 1440,
          left: 1440,
          header: 720,
          footer: 720,
          gutter: 0,
        },
      },
    });
    const margin = result.page!.margin!;
    expect(margin.top).toBe(1440);
    expect(margin.right).toBe(1440);
    expect(margin.bottom).toBe(1440);
    expect(margin.left).toBe(1440);
    expect(margin.header).toBe(720);
    expect(margin.footer).toBe(720);
    expect(margin.gutter).toBe(0);
  });

  it("round-trips section type", () => {
    const result = roundTrip({ type: "nextPage" });
    expect(result.type).toBe("nextPage");
  });

  it("round-trips titlePage", () => {
    const result = roundTrip({ titlePage: true });
    expect(result.titlePage).toBe(true);
  });

  it("round-trips noEndnote", () => {
    const result = roundTrip({ noEndnote: true });
    expect(result.noEndnote).toBe(true);
  });

  it("round-trips formProtection", () => {
    const result = roundTrip({ formProtection: true });
    expect(result.formProtection).toBe(true);
  });

  it("round-trips bidi", () => {
    const result = roundTrip({ bidi: true });
    expect(result.bidi).toBe(true);
  });

  it("round-trips rtlGutter", () => {
    const result = roundTrip({ rtlGutter: true });
    expect(result.rtlGutter).toBe(true);
  });

  it("round-trips verticalAlign", () => {
    const result = roundTrip({ verticalAlign: "center" });
    expect(result.verticalAlign).toBe("center");
  });

  it("round-trips column properties", () => {
    const result = roundTrip({
      column: { count: 3, space: 720 },
    });
    expect(result.column).toBeDefined();
    expect(result.column!.count).toBe(3);
    expect(result.column!.space).toBe(720);
  });

  it("round-trips line numbers", () => {
    const result = roundTrip({
      lineNumbers: { countBy: 5, start: 1, restart: "continuous", distance: 360 },
    });
    expect(result.lineNumbers).toBeDefined();
    expect(result.lineNumbers!.countBy).toBe(5);
    expect(result.lineNumbers!.start).toBe(1);
  });

  it("round-trips docGrid", () => {
    const result = roundTrip({
      grid: { linePitch: 240, charSpace: 100, type: "lines" },
    });
    expect(result.grid).toBeDefined();
    expect(result.grid!.linePitch).toBe(240);
    expect(result.grid!.charSpace).toBe(100);
    expect(result.grid!.type).toBe("lines");
  });

  it("round-trips page numbers", () => {
    const result = roundTrip({
      page: { pageNumbers: { start: 10, formatType: "decimal" } },
    });
    expect(result.page!.pageNumbers).toBeDefined();
    expect(result.page!.pageNumbers!.start).toBe(10);
    expect(result.page!.pageNumbers!.formatType).toBe("decimal");
  });

  it("round-trips paperSrc", () => {
    const result = roundTrip({
      paperSrc: { first: 1, other: 2 },
    });
    expect(result.paperSrc).toBeDefined();
    expect(result.paperSrc!.first).toBe(1);
    expect(result.paperSrc!.other).toBe(2);
  });

  it("round-trips rsid attributes", () => {
    const result = roundTrip({
      rsidR: "00112233",
      rsidRPr: "AABBCCDD",
      rsidSect: "11223344",
    });
    expect(result.rsidR).toBe("00112233");
    expect(result.rsidRPr).toBe("AABBCCDD");
    expect(result.rsidSect).toBe("11223344");
  });

  it("round-trips combined options", () => {
    const result = roundTrip({
      type: "continuous",
      titlePage: true,
      verticalAlign: "both",
      page: {
        size: { width: 11906, height: 16838 },
        margin: {
          top: 1440,
          right: 1800,
          bottom: 1440,
          left: 1800,
          header: 720,
          footer: 720,
          gutter: 0,
        },
      },
    });
    expect(result.type).toBe("continuous");
    expect(result.titlePage).toBe(true);
    expect(result.verticalAlign).toBe("both");
  });
});
