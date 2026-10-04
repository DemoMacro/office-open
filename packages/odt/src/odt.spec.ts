import { generateOcf, ODF_NAMESPACES } from "@office-open/ocf";
import { describe, expect, it } from "vite-plus/test";

import { generateOdt, OdtParseError, parseOdt, type OdtDocumentOptions } from "./index";

const FORM_NAMESPACE = 'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"';

function projection(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value));
}

function odtPackage(body: string): Uint8Array {
  return generateOcf("application/vnd.oasis.opendocument.text", {
    "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} ${FORM_NAMESPACE}><office:automatic-styles/><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
    "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
  });
}

function parseBodyError(body: string, name: string, reason?: string): OdtParseError {
  try {
    parseOdt(odtPackage(body));
  } catch (cause) {
    expect(cause).toBeInstanceOf(OdtParseError);
    const error = cause as OdtParseError;
    expect(error.part).toBe("content.xml");
    expect(error.name).toBe(name);
    if (reason) expect(error.reason).toContain(reason);
    return error;
  }
  throw new Error(`Expected ${name} to reject without a canonical mapping`);
}

describe("ODT canonical projection", () => {
  it("projects a minimal document without hidden root fields", () => {
    const parsed = parseOdt(generateOdt({ sections: [{ children: [{ paragraph: "ODT" }] }] }));
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("rejects invalid packages with a structured error", () => {
    expect(() => parseOdt(new Uint8Array([1, 2, 3]))).toThrow(OdtParseError);
  });

  it("round-trips revisions through canonical insertion and deletion children", () => {
    const options: OdtDocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    insertion: {
                      id: 1,
                      author: "Author",
                      date: "2026-01-01T00:00:00Z",
                      children: ["added"],
                    },
                  },
                  {
                    deletion: {
                      id: 2,
                      author: "Author",
                      date: "2026-01-01T00:00:00Z",
                      children: ["removed"],
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections", "settings"]);
    expect(parsed.settings).toEqual({ trackRevisions: true });
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips a simple form control through a canonical form field", () => {
    const field = {
      formField: { name: "Name", textInput: { value: "Value" } },
    };
    const options: OdtDocumentOptions = {
      sections: [{ children: [{ paragraph: { children: [field] } }] }],
    };
    const parsed = parseOdt(generateOdt(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual([
      { paragraph: { children: [field] } },
    ]);
  });

  it("round-trips bookmarks, references, variables, sequences, and hidden text", () => {
    const options: OdtDocumentOptions = {
      sections: [
        {
          children: [
            { bookmarkStart: { id: 7, name: "target" } },
            {
              paragraph: {
                children: [
                  { bookmarkEnd: { id: 7 } },
                  { simpleField: { instruction: "REF target \\r" } },
                  { simpleField: { instruction: "SEQ Figure", cachedValue: "1" } },
                  { simpleField: { instruction: "VARIABLE Total", cachedValue: "2" } },
                  { simpleField: { instruction: "= 1+1", cachedValue: "2" } },
                  { simpleField: { instruction: 'IF condition "secret" ""' } },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips annotations, indexes, bibliography, and shapes", () => {
    const options: OdtDocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    comment: {
                      author: "Reviewer",
                      date: "2026-01-01T00:00:00Z",
                      children: [{ text: "Body note" }],
                    },
                  },
                ],
              },
            },
            {
              toc: {
                alias: "Figures",
                captionLabelIncludingNumbers: "Figure",
                entries: [{ paragraph: { text: "Illustration" } }],
              },
            },
            {
              sdt: {
                properties: { alias: "Sources", bibliography: true, richText: true },
                children: [{ paragraph: { text: "Entry" } }],
              },
            },
            {
              paragraph: {
                children: [
                  {
                    wpsShape: {
                      children: [],
                      geometry: "rect",
                      transformation: { width: 914400, height: 914400 },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("rejects mapped constructs that have no canonical equivalent", () => {
    const tracked =
      '<text:tracked-changes><text:changed-region text:id="r1"><text:format-change><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes>';
    expect(
      parseBodyError(
        "<office:forms><form:form><form:button form:id='field1'/></form:form></office:forms>",
        "form:button",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:bookmark-start text:name='same'/><text:bookmark-start text:name='same'/>",
        "text:bookmark-start",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError("<text:bookmark-end text:name='missing'/>", "text:bookmark-end").reason,
    ).toContain("no canonical");
    expect(parseBodyError(tracked, "text:format-change").reason).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:sequence-decls><text:sequence-decl text:name='Figure' text:display-outline-level='1'/></text:sequence-decls>",
        "text:sequence-decl",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:p><office:annotation office:display='true'><text:p>shown</text:p></office:annotation></text:p>",
        "office:annotation",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:p><text:hidden-text text:condition='0' text:is-fixed='true'>secret</text:hidden-text></text:p>",
        "text:hidden-text",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:illustration-index><text:illustration-index-source text:index-scope='chapter'/><text:index-body/></text:illustration-index>",
        "text:illustration-index",
      ).reason,
    ).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:p><draw:circle svg:width='1cm' svg:height='1cm'/></text:p>",
        "draw:circle",
      ).reason,
    ).toContain("no canonical");
  });

  it("reports unknown body elements with structured errors", () => {
    const error = parseBodyError("<text:unknown/>", "text:unknown");
    expect(error.path).toBe("/office:document-content/office:body/office:text/text:unknown");
    expect(error.reason).toBe("no canonical DocumentOptions mapping");
  });
});
