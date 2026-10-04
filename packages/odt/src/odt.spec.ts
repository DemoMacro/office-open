import { generateOcf, ODF_NAMESPACES } from "@office-open/ocf";
import { describe, expect, it } from "vite-plus/test";

import { generateOdt, OdtParseError, parseOdt, type OdtDocumentOptions } from "./index";

describe("ODT codec", () => {
  it("round-trips a document through canonical DocumentOptions", () => {
    const parsed = parseOdt(generateOdt({ sections: [{ children: [{ paragraph: "ODT" }] }] }));
    expect(parsed.sections[0]?.children).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("wraps invalid packages in OdtParseError", () => {
    expect(() => parseOdt(new Uint8Array([1, 2, 3]))).toThrow(OdtParseError);
  });

  it("round-trips declarations, tracked changes, forms, and bookmarks", () => {
    const options = {
      sections: [
        { children: [{ bookmarkStart: { id: 7, name: "target" } }, { paragraph: "ODT" }] },
      ],
      odfSemantics: {
        sequenceDeclarations: [{ name: "Figure", displayOutlineLevel: 1 }],
        variableDeclarations: [{ name: "Total", valueType: "float" }],
        trackedChanges: {
          trackChanges: true,
          changes: [
            { id: "ct1", kind: "insertion", author: "Author", date: "2026-01-01T00:00:00" },
          ],
        },
        forms: { forms: [{ name: "Editor", controls: [{ kind: "form:text", id: "name" }] }] },
      },
    } satisfies OdtDocumentOptions;
    const parsed = parseOdt(generateOdt(options));
    expect(parsed.odfSemantics).toMatchObject(options.odfSemantics ?? {});
    expect(parsed.sections[0]?.children?.[0]).toEqual({ bookmarkStart: { id: 7, name: "target" } });
  });

  it("round-trips text section identity, style, protection, and order", () => {
    const parsed = parseOdt(
      generateOdt({
        sections: [
          {
            children: [
              { paragraph: "before" },
              {
                sdt: {
                  properties: {
                    alias: "Region",
                    tag: "odf:text-section;style=Section\\;Style",
                    lock: "sdtLocked",
                  },
                  children: [{ paragraph: "first" }, { paragraph: "second" }],
                },
              },
              { paragraph: "after" },
            ],
          },
        ],
      }),
    );
    expect(parsed.sections[0]?.children).toEqual([
      { paragraph: { text: "before" } },
      {
        sdt: {
          properties: {
            alias: "Region",
            tag: "odf:text-section;style=Section\\;Style",
            lock: "sdtLocked",
          },
          children: [{ paragraph: { text: "first" } }, { paragraph: { text: "second" } }],
        },
      },
      { paragraph: { text: "after" } },
    ]);
  });

  it("round-trips body annotations, indexes, references, fields, and shapes", () => {
    const options = {
      sections: [
        {
          children: [
            {
              bibliography: {
                source: { title: "Sources" },
                paragraphs: ["Entry"],
              },
            },
            {
              index: {
                source: { scope: "document", useCaption: true, title: "Figures" },
                paragraphs: ["Illustration"],
              },
            },
            { shape: { geometry: "ellipse", name: "Mark", x: "1cm", width: "2cm" } },
            { shape: { geometry: "rect", name: "Panel", width: "3cm" } },
            {
              shapeHyperlink: {
                href: "https://example.invalid/",
                shapes: [{ geometry: "rect", name: "Linked" }],
              },
            },
            {
              paragraph: {
                children: [
                  {
                    annotation: {
                      author: "Reviewer",
                      date: "2026-01-01T00:00:00",
                      paragraphs: ["Body note"],
                    },
                  },
                  { bookmarkEnd: { name: "range" } },
                  { bookmarkReference: { name: "range", referenceFormat: "number" } },
                  { variableSet: { name: "Total", valueType: "float", value: "2", display: "2" } },
                  {
                    sequence: {
                      name: "Figure",
                      referenceName: "ref",
                      display: "1",
                    },
                  },
                  { chapter: { display: "1", outlineLevel: 1 } },
                  {
                    expression: {
                      formula: "1+1",
                      valueType: "float",
                      value: "2",
                      display: "2",
                    },
                  },
                  { hiddenText: { condition: "0", content: "secret" } },
                  { referenceMarkStart: { name: "mark" } },
                  { referenceMarkEnd: { name: "mark" } },
                  { referenceReference: { name: "mark", referenceFormat: "page" } },
                  {
                    bibliographyMark: {
                      type: "article",
                      identifier: "source-1",
                      fields: [{ name: "title", value: "Paper" }],
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    } as unknown as OdtDocumentOptions;
    const parsed = parseOdt(generateOdt(options));
    expect(parsed.sections[0]?.children).toMatchObject(options.sections[0]?.children ?? []);
  });

  it("reports unknown body elements with structured errors", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:text><text:unknown/></office:text></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseOdt(
        generateOcf("application/vnd.oasis.opendocument.text", {
          "content.xml": content,
          "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdtParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: "/office:document-content/office:body/office:text/text:unknown",
      name: "text:unknown",
      reason: "no canonical DocumentOptions mapping",
    });
  });
});
