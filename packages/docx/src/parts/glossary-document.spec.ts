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
          sections: [],
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
          sections: [],
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
          sections: [],
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
          sections: [],
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
          sections: [],
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
          sections: [],
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
          sections: [],
        },
      ],
    });
    expect(result.parts[0]?.style).toBe("Header/Footer");
  });

  it("round-trips multiple parts", () => {
    const result = roundTrip({
      parts: [
        { name: "Part1", gallery: "default", sections: [] },
        { name: "Part2", gallery: "hdrs", sections: [] },
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

  it("preserves paragraph-hosted and terminal section properties", () => {
    const doc = parseXml(
      '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Sectioned"/><w:category><w:gallery w:val="default"/></w:category></w:docPartPr><w:docPartBody>' +
        '<w:p><w:pPr><w:sectPr w:rsidR="001C720C"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:right="360"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>' +
        "<w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
        '<w:sectPr w:rsidR="009C39E9"><w:pgSz w:w="15840" w:h="12240"/><w:cols w:num="2"/></w:sectPr>' +
        "</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>",
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");

    const parsed = glossaryDesc.parse(root, readCtx);
    const sections = parsed.parts[0]?.sections ?? [];
    expect(sections).toHaveLength(2);
    expect(sections[0]?.properties).toMatchObject({
      additionRsid: "001C720C",
      pageSize: { width: 12240, height: 15840 },
      pageMargin: { right: 360 },
    });
    expect(sections[1]?.properties).toMatchObject({
      additionRsid: "009C39E9",
      pageSize: { width: 15840, height: 12240 },
      columns: { count: 2 },
    });

    const xml = glossaryDesc.stringify(parsed, writeCtx)!;
    expect(xml).toMatch(/<w:pPr><w:sectPr[^>]*w:rsidR="001C720C"/);
    expect(xml).toMatch(/<w:sectPr[^>]*w:rsidR="009C39E9"[^>]*>.*<\/w:sectPr><\/w:docPartBody>/);
  });

  it("creates a section-break paragraph when a non-final section cannot host sectPr", () => {
    const xml = glossaryDesc.stringify(
      {
        parts: [
          {
            name: "TableSections",
            gallery: "default",
            sections: [
              {
                children: [{ rawXml: "<w:tbl/>" }],
                properties: { pageSize: { width: 12240, height: 15840 } },
              },
              { children: [] },
            ],
          },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain("<w:tbl/><w:p><w:pPr><w:sectPr");
  });

  it("parses every section child in a building block body", () => {
    const doc = parseXml(
      '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="TableBlock"/><w:category><w:gallery w:val="tbls"/></w:category></w:docPartPr><w:docPartBody><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:unknown/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>',
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");

    const result = glossaryDesc.parse(root, readCtx);
    const children = result.parts[0]?.sections[0]?.children ?? [];
    expect(children).toHaveLength(2);
    expect(children[0]).toHaveProperty("table");
    expect(JSON.stringify(children[0])).toContain("Cell text");
    expect(children[1]).toEqual({ rawXml: "<w:unknown/>" });
  });
});
