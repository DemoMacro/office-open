import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseChild, stringifyChild } from "./bridge";

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;
const writeCtx = {} as never;

describe("slide child bridge", () => {
  it("parses content parts without rawXml absorption", () => {
    const doc = parseXml(
      '<p:contentPart xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId12"/>',
    );
    const element = doc.elements?.[0];
    if (!element) throw new Error("expected a content part element");
    expect(parseChild(element, readCtx)).toEqual({ contentPart: { rId: "rId12" } });
  });

  it("round-trips PowerPoint 2010 content part extensions", () => {
    const doc = parseXml(
      '<p:contentPart xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:bwMode="auto" r:id="rId12">' +
        '<p14:nvContentPartPr><p14:cNvPr id="50" name="Ink"/>' +
        "<p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr>" +
        '<p14:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p14:xfrm>' +
        "</p:contentPart>",
    );
    const element = doc.elements?.[0];
    if (!element) throw new Error("expected a content part element");
    const parsed = parseChild(element, readCtx);
    expect(parsed).toEqual({
      contentPart: {
        rId: "rId12",
        blackWhiteMode: "auto",
        nonVisualProperties: {
          drawingProperties: { id: "50", name: "Ink" },
          contentPartProperties: undefined,
          properties: {},
        },
        transform: { x: 10, y: 20, width: 30, height: 40 },
      },
    });
    const generated = stringifyChild(parsed as Parameters<typeof stringifyChild>[0], writeCtx);
    expect(generated).toContain('p14:bwMode="auto"');
    expect(generated).toContain("<p14:nvContentPartPr>");
    expect(generated).toContain(
      '<p14:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p14:xfrm>',
    );
    const generatedDocument = parseXml(generated ?? "");
    const generatedElement = generatedDocument.elements?.[0];
    if (!generatedElement) throw new Error("expected generated content part XML");
    const reparsed = parseChild(generatedElement, readCtx);
    expect(reparsed).toEqual(parsed);
  });

  it("preserves compatibility branches without choosing one", () => {
    const doc = parseXml(
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        '<mc:Choice xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" Requires="control">' +
        '<p:contentPart r:id="rId12"/></mc:Choice>' +
        '<mc:Fallback xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Fallback"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
        "<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>" +
        "</mc:Fallback></mc:AlternateContent>",
    );
    const element = doc.elements?.[0];
    if (!element) throw new Error("expected a compatibility element");
    const parsed = parseChild(element, readCtx);
    expect(parsed).toEqual({
      alternateContent: {
        choices: [{ requires: "control", child: { contentPart: { rId: "rId12" } } }],
        fallback: { picture: expect.objectContaining({ name: "Fallback" }) },
      },
    });
  });

  it("preserves an unrecognized compatibility branch as a typed branch", () => {
    const xml =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice Requires="future"><future:node xmlns:future="urn:future"/></mc:Choice></mc:AlternateContent>';
    const doc = parseXml(xml);
    const element = doc.elements?.[0];
    if (!element) throw new Error("expected a compatibility element");
    const parsed = parseChild(element, readCtx);
    expect(parsed).toEqual({
      alternateContent: {
        choices: [
          {
            requires: "future",
            child: { rawXml: '<future:node xmlns:future="urn:future"/>' },
          },
        ],
      },
    });
  });
});
