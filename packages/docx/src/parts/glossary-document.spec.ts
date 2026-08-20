import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import { parseSectionChild } from "../parse/body";
import { glossaryDesc } from "./glossary-document";
import type { GlossaryDocumentOptions } from "./glossary-document";
import { setTableParseChild } from "./table/descriptor";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) => (typeof child === "string" ? child : "<w:p/>"),
  fileData: {} as never,
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

setTableParseChild(parseSectionChild);

function roundTrip(opts: GlossaryDocumentOptions) {
  const xml = glossaryDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return glossaryDesc.parse(el, readCtx);
}

describe("glossaryDesc round-trip", () => {
  it("round-trips a simple building block", () => {
    const result = roundTrip({
      parts: [
        {
          name: "TestBlock",
          gallery: "default",
          children: [],
        },
      ],
    });
    expect(result.parts).toHaveLength(1);
    expect(result.parts[0]?.name).toBe("TestBlock");
    expect(result.parts[0]?.gallery).toBe("default");
  });

  it("round-trips category and gallery", () => {
    const result = roundTrip({
      parts: [
        {
          name: "CoverPage",
          gallery: "coverPg",
          category: "Built-In",
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.gallery).toBe("coverPg");
    expect(result.parts[0]?.category).toBe("Built-In");
  });

  it("round-trips types", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Typed",
          gallery: "default",
          types: ["normal", "autoExp"],
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.types).toEqual(["normal", "autoExp"]);
  });

  it("round-trips behaviors", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Behaved",
          gallery: "default",
          behaviors: ["content", "p"],
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.behaviors).toEqual(["content", "p"]);
  });

  it("round-trips description", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Described",
          gallery: "default",
          description: "A test building block",
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.description).toBe("A test building block");
  });

  it("round-trips guid", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Guided",
          gallery: "default",
          guid: "12345678-ABCD-EF01-2345-6789ABCDEF01",
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.guid).toBe("12345678-ABCD-EF01-2345-6789ABCDEF01");
  });

  it("round-trips the building block style", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Styled",
          gallery: "default",
          style: "Header/Footer",
          children: [],
        },
      ],
    });
    expect(result.parts[0]?.style).toBe("Header/Footer");
  });

  it("round-trips multiple parts", () => {
    const result = roundTrip({
      parts: [
        { name: "Part1", gallery: "default", children: [] },
        { name: "Part2", gallery: "hdrs", children: [] },
      ],
    });
    expect(result.parts).toHaveLength(2);
    expect(result.parts[0]?.name).toBe("Part1");
    expect(result.parts[1]?.name).toBe("Part2");
  });

  it("round-trips empty parts", () => {
    const result = roundTrip({ parts: [] });
    expect(result.parts).toHaveLength(0);
  });

  it("parses every section child in a building block body", () => {
    const doc = parseXml(
      '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="TableBlock"/><w:category><w:gallery w:val="tbls"/></w:category></w:docPartPr><w:docPartBody><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:unknown/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>',
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");

    const result = glossaryDesc.parse(root, readCtx);
    const children = result.parts[0]?.children ?? [];
    expect(children).toHaveLength(2);
    expect(children[0]).toHaveProperty("table");
    expect(JSON.stringify(children[0])).toContain("Cell text");
    expect(children[1]).toEqual({ rawXml: "<w:unknown/>" });
  });
});
