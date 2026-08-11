import { findChild, parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { PptxWriteContext } from "../../context";
import { parseColorMap, parseHeaderFooter } from "../handout-master";
import type { SlideMasterDescriptorOptions } from "./slide-master";
import { slideMasterDesc } from "./slide-master";
import type { TextListStyleOptions } from "./text-list-style";
import { parseTextListStyle } from "./text-list-style";

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as const;

function makeWriteCtx() {
  const ctx = new PptxWriteContext();
  ctx.slideWidth = 12192000;
  return ctx;
}

function freshXml(opts: SlideMasterDescriptorOptions = {}, ctx = makeWriteCtx()): string {
  const xml = slideMasterDesc.stringify(opts, ctx);
  if (!xml) throw new Error("stringify returned undefined");
  return xml;
}

function roundTrip(opts: SlideMasterDescriptorOptions): SlideMasterDescriptorOptions {
  const xml = freshXml(opts);
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return slideMasterDesc.parse(el, readCtx);
}

describe("slideMasterDesc fresh emit", () => {
  it("emits the standard structure (bgRef/spTree/clrMap/sldLayoutIdLst/hf/txStyles)", () => {
    const xml = freshXml();
    expect(xml).toContain("<p:sldMaster");
    expect(xml).toContain('<p:bgRef idx="1001">'); // MS Office default background
    expect(xml).toContain("<p:spTree>");
    expect(xml).toContain("<p:clrMap ");
    expect(xml).toContain("<p:sldLayoutIdLst>");
    expect(xml).toContain("<p:hf ");
    expect(xml).toContain("<p:txStyles>");
  });

  it("emits the @preserve attribute only when requested", () => {
    expect(freshXml()).not.toContain("preserve=");
    expect(freshXml({ preserve: true })).toContain('preserve="1"');
  });
});

describe("slideMasterDesc round-trip", () => {
  it("round-trips placeholders (positions derived from spTree)", () => {
    const result = roundTrip({ placeholders: { title: true, body: true } });
    expect(result.placeholders).toBeDefined();
    const title = result.placeholders?.title;
    const body = result.placeholders?.body;
    expect(typeof title).toBe("object");
    expect(typeof body).toBe("object");
    // title position (EMU) at the reference slide width survives
    expect((title as { x: number }).x).toBe(838200);
    expect((body as { x: number }).x).toBe(838200);
  });

  it('round-trips a hidden placeholder (sz="0")', () => {
    // A round-tripped master with a hidden date placeholder parses to false.
    const xml =
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Date"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="dt" sz="0" idx="2"/></p:nvPr></p:nvSpPr>' +
      "<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>" +
      '</p:spTree></p:cSld><p:clrMap bg1="lt1"/></p:sldMaster>';
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = slideMasterDesc.parse(el, readCtx);
    expect(result.placeholders?.date).toBe(false);
  });

  it("round-trips preserve attribute", () => {
    expect(roundTrip({ preserve: true }).preserve).toBe(true);
    expect(roundTrip({}).preserve).toBeUndefined();
  });

  it("round-trips slideLayoutIds", () => {
    const ids = [
      { id: 2147483649, relationshipId: "rId1" },
      { id: 2147483650, relationshipId: "rId2" },
    ];
    expect(roundTrip({ slideLayoutIds: ids }).slideLayoutIds).toEqual(ids);
  });

  it("round-trips custom children shapes", () => {
    const result = roundTrip({
      children: [{ shape: { x: 100, y: 100, width: 200, height: 200, fill: "FF0000" } }],
    });
    expect(result.children?.length).toBe(1);
  });
});

describe("slide-master colorMap/headerFooter round-trip", () => {
  it("emits custom colorMap and headerFooter, parseable back", () => {
    const el = parseXml(
      freshXml({
        colorMap: { bg1: "dk1", tx1: "lt1" },
        headerFooter: { date: true, footer: true },
      }),
    ).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMap(findChild(el, "p:clrMap"));
    expect(clrMap?.bg1).toBe("dk1");
    expect(clrMap?.tx1).toBe("lt1");
    expect(clrMap?.bg2).toBe("lt2"); // untouched keys keep standard defaults

    const hf = parseHeaderFooter(findChild(el, "p:hf"));
    expect(hf?.date).toBe(true);
    expect(hf?.footer).toBe(true);
    expect(hf?.header).toBe(false);
    expect(hf?.slideNumber).toBe(false);
  });

  it("emits standard defaults when colorMap/headerFooter are undefined", () => {
    const el = parseXml(freshXml()).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMap(findChild(el, "p:clrMap"));
    expect(clrMap?.bg1).toBe("lt1");
    expect(clrMap?.tx1).toBe("dk1");

    const hf = parseHeaderFooter(findChild(el, "p:hf"));
    expect(hf?.date).toBe(false);
    expect(hf?.slideNumber).toBe(false);
  });
});

describe("slide-master textStyles round-trip", () => {
  it("fresh master emits the default txStyles block", () => {
    const xml = freshXml();
    expect(xml).toContain("<p:txStyles>");
    expect(xml).toContain("<p:titleStyle>");
    expect(xml).toContain("<p:bodyStyle>");
    expect(xml).toContain("<p:otherStyle>");
    expect(xml).toContain('sz="4400"'); // title lvl1 default size
  });

  it("structured custom textStyles is emitted, replacing the default", () => {
    const custom: TextListStyleOptions = {
      title: { levels: [{ defaultRun: { size: 9000 } }] },
      body: { levels: [{}] },
      other: { levels: [{}] },
    };
    const xml = freshXml({ textStyles: custom });
    expect(xml).toContain('sz="9000"');
    expect(xml).not.toContain('sz="4400"');
  });

  it("round-trips a parsed master's txStyles structured", () => {
    const source: TextListStyleOptions = {
      title: { levels: [{ defaultRun: { size: 9999 } }] },
      body: { levels: [{}] },
      other: { levels: [{}] },
    };
    const el = parseXml(freshXml({ textStyles: source })).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const extracted = parseTextListStyle(findChild(el, "p:txStyles")!);

    const reEmitted = freshXml({ textStyles: extracted });
    expect(reEmitted).toContain('sz="9999"');
    expect(reEmitted).not.toContain('sz="4400"');
  });
});
