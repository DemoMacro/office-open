import { parseColorMapping } from "@office-open/core";
import { findChild, parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { PptxWriteContext } from "../../context";
import { parseHeaderFooter } from "../handout-master";
import type { SlideMasterDescriptorOptions } from "./slide-master";
import { slideMasterDesc } from "./slide-master";
import type { TextStylesOptions } from "./text-list-style";
import { parseTextStyles } from "./text-list-style";

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
    // MS Office masters carry no p:hf by default — only when the source had one.
    expect(xml).not.toContain("<p:hf");
    expect(freshXml({ headerFooter: { slideNumber: false } })).toContain("<p:hf ");
    expect(xml).toContain("<p:txStyles>");
  });

  it("emits the @preserve attribute only when requested", () => {
    expect(freshXml()).not.toContain("preserve=");
    expect(freshXml({ preserve: true })).toContain('preserve="1"');
  });
});

describe("slideMasterDesc round-trip", () => {
  it("round-trips placeholders (positions derived from spTree)", () => {
    const result = roundTrip({ placeholders: {} });
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

  it("keeps non-standard placeholder shapes in children and records the map entry", () => {
    // sldImg has no fresh-emit branch: the whole sp stays in children (re-emitted
    // from there) and the definition lands on the map for inheritance only.
    const spTreeOpen =
      '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Image Placeholder"/><p:cNvSpPr/>' +
      '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="300000" cy="400000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>";
    const xml =
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      spTreeOpen +
      sp +
      '</p:spTree></p:cSld><p:clrMap bg1="lt1"/></p:sldMaster>';
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = slideMasterDesc.parse(el, readCtx);

    expect(result.children?.length).toBe(1);
    const slideImage = result.placeholders?.slideImage;
    expect(typeof slideImage).toBe("object");
    expect((slideImage as { x: number }).x).toBe(100000);
    // Non-standard keys do not participate in the standard-slot false fill.
    expect(result.placeholders?.header).toBeUndefined();

    const reEmitted = freshXml(result as SlideMasterDescriptorOptions);
    expect(reEmitted).toContain('type="sldImg"');
  });

  it("maps the pic placeholder token to the picture key", () => {
    const xml =
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Picture Placeholder"/><p:cNvSpPr/>' +
      '<p:nvPr><p:ph type="pic"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="300000" cy="400000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>" +
      '</p:spTree></p:cSld><p:clrMap bg1="lt1"/></p:sldMaster>';
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = slideMasterDesc.parse(el, readCtx);
    expect(typeof result.placeholders?.picture).toBe("object");
    expect(result.children?.length).toBe(1);
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
      children: [
        { shape: { x: 100, y: 100, width: 200, height: 200, properties: { fill: "FF0000" } } },
      ],
    });
    expect(result.children?.length).toBe(1);
  });
});

describe("slide-master placeholder facets round-trip", () => {
  it("round-trips a non-rect geometry facet", () => {
    const result = roundTrip({
      placeholders: {
        title: { x: 100, y: 100, width: 200, height: 200, geometry: "ellipse" },
      },
    });
    const title = result.placeholders?.title as { geometry?: { preset?: string } };
    expect(title.geometry?.preset).toBe("ellipse");
  });

  it("round-trips a shape style facet", () => {
    const result = roundTrip({
      placeholders: {
        title: {
          x: 100,
          y: 100,
          width: 200,
          height: 200,
          style: { fillReference: { index: 2 } },
        },
      },
    });
    const title = result.placeholders?.title as { style?: { fillReference?: { index: number } } };
    expect(title.style?.fillReference?.index).toBe(2);
  });

  it("round-trips a fill facet on a placeholder", () => {
    const result = roundTrip({
      placeholders: {
        body: { x: 100, y: 100, width: 200, height: 200, fill: { type: "solid", color: "FF0000" } },
      },
    });
    const body = result.placeholders?.body as {
      fill?: { type: string; color?: { value: string } };
    };
    expect(body.fill?.type).toBe("solid");
    expect(body.fill?.color?.value).toBe("FF0000");
  });

  it("does not carry the default rect geometry as a facet", () => {
    // rect is the placeholder default — extraction omits it so the fresh emit
    // path stays byte-equivalent with MS Office's master output.
    const result = roundTrip({ placeholders: {} });
    expect((result.placeholders?.title as { geometry?: string })?.geometry).toBeUndefined();
  });
});

describe("slide-master colorMapping/headerFooter round-trip", () => {
  it("emits custom colorMapping and headerFooter, parseable back", () => {
    const el = parseXml(
      freshXml({
        colorMapping: { background1: "dark1", text1: "light1" },
        headerFooter: { date: true, footer: true },
      }),
    ).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMapping(findChild(el, "p:clrMap"));
    expect(clrMap?.background1).toBe("dark1");
    expect(clrMap?.text1).toBe("light1");
    expect(clrMap?.background2).toBe("light2"); // untouched keys keep standard defaults

    const hf = parseHeaderFooter(findChild(el, "p:hf"));
    expect(hf?.date).toBe(true);
    expect(hf?.footer).toBe(true);
    expect(hf?.header).toBe(false);
    expect(hf?.slideNumber).toBe(false);
  });

  it("emits standard defaults when colorMapping/headerFooter are undefined", () => {
    const el = parseXml(freshXml()).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    const clrMap = parseColorMapping(findChild(el, "p:clrMap"));
    expect(clrMap?.background1).toBe("light1");
    expect(clrMap?.text1).toBe("dark1");

    // No headerFooter → no p:hf at all (MS Office masters carry none).
    expect(findChild(el, "p:hf")).toBeUndefined();
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
    const custom: TextStylesOptions = {
      title: { levels: [{ defaultRunProperties: { size: 90 } }] },
      body: { levels: [{}] },
      other: { levels: [{}] },
    };
    const xml = freshXml({ textStyles: custom });
    expect(xml).toContain('sz="9000"');
    expect(xml).not.toContain('sz="4400"');
  });

  it("round-trips a parsed master's txStyles structured", () => {
    // 11.1pt is a non-integer-hundredth size (11.1 * 100 drifts to
    // 1110.0000000000002 in float); verifies Math.round keeps the round-trip
    // byte-equal instead of emitting an XSD-invalid fractional sz.
    const source: TextStylesOptions = {
      title: { levels: [{ defaultRunProperties: { size: 11.1 } }] },
      body: { levels: [{}] },
      other: { levels: [{}] },
    };
    const el = parseXml(freshXml({ textStyles: source })).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const extracted = parseTextStyles(findChild(el, "p:txStyles")!, readCtx);

    const reEmitted = freshXml({ textStyles: extracted });
    expect(reEmitted).toContain('sz="1110"');
    expect(reEmitted).not.toContain('sz="4400"');
  });
});
