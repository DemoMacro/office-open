import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { findChild, parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseColorMap, parseHeaderFooter } from "../handout-master";
import { buildSlideMasterXml } from "../slide-master";
import type { SlideMasterDescriptorOptions } from "./slide-master";
import { slideMasterDesc } from "./slide-master";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: SlideMasterDescriptorOptions) {
  const xml = slideMasterDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return slideMasterDesc.parse(el, readCtx);
}

describe("slideMasterDesc round-trip", () => {
  it("round-trips a minimal slide master XML", () => {
    const masterXml =
      '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      "</p:spTree></p:cSld></p:sldMaster>";

    const opts: SlideMasterDescriptorOptions = { master: masterXml };
    const result = roundTrip(opts);

    // The master property should contain a valid XML string
    expect(result.master).toBeDefined();
    expect(typeof result.master).toBe("string");
    expect(result.master.length).toBeGreaterThan(0);

    // Re-parse the round-tripped master to verify it's valid XML.
    // Note: xmlStringify(el) serializes child elements only, so the
    // root p:sldMaster tag is stripped — the result starts with p:cSld.
    const doc2 = parseXml(result.master);
    expect(doc2.elements).toBeDefined();
    expect(doc2.elements?.length).toBe(1);
    expect(doc2.elements?.[0]?.name).toBe("p:cSld");
  });

  it("round-trips a slide master with nested content", () => {
    const masterXml =
      '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:bgPr></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      "</p:spTree></p:cSld></p:sldMaster>";

    const opts: SlideMasterDescriptorOptions = { master: masterXml };
    const result = roundTrip(opts);

    expect(result.master).toBeDefined();
    // Verify the color survived round-trip
    expect(result.master).toContain("4472C4");
    expect(result.master).toContain("srgbClr");
  });
});

describe("slide-master colorMap/headerFooter round-trip", () => {
  it("emits custom colorMap and headerFooter, parseable back", () => {
    const xml = buildSlideMasterXml(0, undefined, {
      colorMap: { bg1: "dk1", tx1: "lt1" },
      headerFooter: { date: true, footer: true },
    });
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMap(findChild(el, "p:clrMap"));
    expect(clrMap?.bg1).toBe("dk1");
    expect(clrMap?.tx1).toBe("lt1");
    // Untouched keys keep the standard defaults.
    expect(clrMap?.bg2).toBe("lt2");

    const hf = parseHeaderFooter(findChild(el, "p:hf"));
    expect(hf?.date).toBe(true);
    expect(hf?.footer).toBe(true);
    expect(hf?.header).toBe(false);
    expect(hf?.slideNumber).toBe(false);
  });

  it("emits standard defaults when colorMap/headerFooter are undefined", () => {
    const xml = buildSlideMasterXml(0, undefined, undefined);
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMap(findChild(el, "p:clrMap"));
    expect(clrMap?.bg1).toBe("lt1");
    expect(clrMap?.tx1).toBe("dk1");

    const hf = parseHeaderFooter(findChild(el, "p:hf"));
    expect(hf?.date).toBe(false);
    expect(hf?.slideNumber).toBe(false);
  });
});
