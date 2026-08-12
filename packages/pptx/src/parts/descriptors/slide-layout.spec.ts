import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { LayoutDefinition } from "@shared/file";
import { describe, expect, it } from "vite-plus/test";

import { PptxWriteContext } from "../../context";
import { slideLayoutDesc } from "./slide-layout";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as PptxWriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: LayoutDefinition): LayoutDefinition {
  const xml = slideLayoutDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return slideLayoutDesc.parse(el, readCtx);
}

function parseXmlDef(xml: string): LayoutDefinition {
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return slideLayoutDesc.parse(el, readCtx);
}

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;

// A realistic layout: title placeholder (positioned) + body placeholder (no xfrm).
const LAYOUT_WITH_PLACEHOLDERS = `<p:sldLayout ${NS} type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

describe("slideLayoutDesc stringify/parse", () => {
  it("round-trips CT_SlideLayout attributes", () => {
    const result = roundTrip({
      type: "title",
      name: "Title Slide",
      matchingName: "Title Slide",
      preserve: true,
      userDrawn: true,
      showMasterSp: false,
    });
    expect(result.type).toBe("title");
    expect(result.name).toBe("Title Slide");
    expect(result.matchingName).toBe("Title Slide");
    expect(result.preserve).toBe(true);
    expect(result.userDrawn).toBe(true);
    expect(result.showMasterSp).toBe(false);
  });

  it("defaults colorMapOverride to masterClrMapping and round-trips it", () => {
    const result = roundTrip({ type: "blank" });
    expect(result.colorMapOverride).toEqual({ kind: "master" });
  });

  it("round-trips an explicit color-mapping override", () => {
    const result = roundTrip({
      type: "blank",
      colorMapOverride: { kind: "override", mapping: { bg1: "dk1", tx1: "lt1" } },
    });
    expect(result.colorMapOverride).toEqual({
      kind: "override",
      mapping: { bg1: "dk1", tx1: "lt1" },
    });
  });

  it("round-trips a solidFill background", () => {
    const result = roundTrip({
      type: "blank",
      background: { fill: { type: "solid", color: "FF0000" } },
    });
    // Color round-trips as the EG_ColorChoice form ({ value }) on parse.
    expect(result.background?.fill).toEqual({ type: "solid", color: { value: "FF0000" } });
  });

  it("round-trips a transition", () => {
    const result = roundTrip({
      type: "blank",
      transition: { type: "fade", speed: "slow" },
    });
    expect(result.transition).toEqual({ type: "fade", speed: "slow" });
  });

  it("parses spTree children and derives placeholder positions", () => {
    const result = parseXmlDef(LAYOUT_WITH_PLACEHOLDERS);
    expect(result.type).toBe("title");
    expect(result.name).toBe("Title Slide");
    expect(result.children).toHaveLength(2);
    // Title placeholder position read from a:xfrm.
    expect(result.placeholders?.title).toEqual({ x: 100, y: 200, width: 300, height: 400 });
  });

  it("derives type from cSld name when @type is absent", () => {
    const result = parseXmlDef(`<p:sldLayout ${NS}><p:cSld name="Blank"/></p:sldLayout>`);
    expect(result.type).toBe("blank");
  });

  it("edits to children take effect on stringify (verbatim bypass broken)", () => {
    // The core A1 goal: a parsed layout's structured children drive stringify,
    // so removing a shape is reflected in the emitted XML.
    const parsed = parseXmlDef(LAYOUT_WITH_PLACEHOLDERS);
    expect(parsed.children).toHaveLength(2);
    // Drop the first child (title) — simulating an editor deletion.
    const edited: LayoutDefinition = { ...parsed, children: parsed.children!.slice(1) };
    const xml = slideLayoutDesc.stringify(edited, writeCtx)!;
    expect(xml).not.toContain('name="Title 1"');
    // Re-parse confirms only one shape remains.
    const reParsed = parseXmlDef(xml);
    expect(reParsed.children).toHaveLength(1);
  });
});
