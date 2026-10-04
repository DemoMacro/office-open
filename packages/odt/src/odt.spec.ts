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
    const options: OdtDocumentOptions = {
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
    };
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
