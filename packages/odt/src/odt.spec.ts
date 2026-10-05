import type { ParagraphOptions } from "@office-open/docx";
import { generateOcf, ODF_NAMESPACES } from "@office-open/odf";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import {
  generateDocument,
  OdtParseError,
  parseDocument,
  type DocumentOptions,
  type OdtDocumentOptions,
} from "./index";

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

function odtStyledPackage(styles: string, body: string): Uint8Array {
  return generateOcf("application/vnd.oasis.opendocument.text", {
    "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} ${FORM_NAMESPACE}><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
    "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
  });
}

function parseBodyError(body: string, name: string, reason?: string): OdtParseError {
  try {
    parseDocument(odtPackage(body));
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
    const parsed = parseDocument(
      generateDocument({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
    );
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("rejects invalid packages with a structured error", () => {
    expect(() => parseDocument(new Uint8Array([1, 2, 3]))).toThrow(OdtParseError);
  });

  it("round-trips mixed canonical and overlay paragraph styles", () => {
    const source = odtStyledPackage(
      '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center" style:contextual-spacing="true"/></style:style>',
      '<text:p text:style-name="P1">Styled</text:p>',
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: { text: "Styled", style: "P1", alignment: "center" },
    });
    expect(parsed.styleOverlays).toEqual([
      {
        name: "P1",
        family: "paragraph",
        properties: [
          {
            name: "style:paragraph-properties",
            attributes: { "fo:text-align": "center", "style:contextual-spacing": "true" },
          },
        ],
      },
    ]);

    expect(
      (parseDocument(generateDocument(parsed)) as OdtDocumentOptions).sections[0]?.children?.[0],
    ).toMatchObject({
      paragraph: { text: "Styled", style: "P1", alignment: "center" },
    });
  });

  it("round-trips LibreOffice graphic-properties overlays", () => {
    const source = odtStyledPackage(
      '<style:style style:name="P1" style:family="paragraph"><loext:graphic-properties draw:fill="none"/><style:paragraph-properties fo:text-align="center"/></style:style>',
      '<text:p text:style-name="P1">Styled</text:p>',
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.styleOverlays?.[0]?.properties[0]?.name).toBe("loext:graphic-properties");
    const reread = parseDocument(generateDocument(parsed)) as OdtDocumentOptions;
    expect(reread.styleOverlays).toEqual(parsed.styleOverlays);
  });

  it("round-trips mixed canonical and overlay run styles", () => {
    const source = odtStyledPackage(
      '<style:style style:name="C1" style:family="text"><style:text-properties fo:font-weight="bold" fo:background-color="#00ff00"/></style:style>',
      '<text:p><text:span text:style-name="C1">Styled</text:span></text:p>',
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: { children: [{ text: "Styled", bold: true, style: "C1" }] },
    });

    expect(
      (parseDocument(generateDocument(parsed)) as OdtDocumentOptions).sections[0]?.children?.[0],
    ).toMatchObject({
      paragraph: { children: [{ text: "Styled", bold: true, style: "C1" }] },
    });
  });

  it("round-trips a linked picture with a parent-relative URL", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    picture: {
                      type: "png",
                      sourceUrl: "../images/logo.png",
                      transformation: { width: 360000, height: 360000 },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseDocument(generateDocument(options));
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: {
        children: [
          {
            picture: {
              type: "png",
              sourceUrl: "../images/logo.png",
              transformation: { width: 360000, height: 360000 },
            },
          },
        ],
      },
    });
  });

  it("preserves a standalone embedded object through a typed overlay", () => {
    const source = generateOcf(
      "application/vnd.oasis.opendocument.text",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles/><office:body><office:text><text:p><draw:frame svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame><draw:frame svg:width="2cm" svg:height="2cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p></office:text></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        "Object 1/content.xml":
          '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      },
      {},
      {
        version: "1.3",
        entries: [
          { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
          { fullPath: "content.xml", mediaType: "text/xml" },
          { fullPath: "styles.xml", mediaType: "text/xml" },
          { fullPath: "Object 1/", mediaType: "application/vnd.oasis.opendocument.presentation" },
        ],
      },
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.embeddedObjects).toEqual([
      {
        path: "Object 1",
        mediaType: "application/vnd.oasis.opendocument.presentation",
        members: [
          {
            path: "Object 1/content.xml",
            mediaType: undefined,
            data: expect.stringContaining("<office:document-content"),
          },
        ],
        frames: [{ shapeId: 1 }, { shapeId: 2 }],
      },
    ]);

    const reread = parseDocument(generateDocument(parsed)) as OdtDocumentOptions;
    expect(reread.embeddedObjects).toEqual(parsed.embeddedObjects);
    const manifest = new TextDecoder().decode(
      unzipSync(generateDocument(parsed))["META-INF/manifest.xml"],
    );
    const content = new TextDecoder().decode(unzipSync(generateDocument(parsed))["content.xml"]);
    expect(content.match(/xlink:href="\.\/Object 1"/g)).toHaveLength(2);
    expect(manifest).toContain('manifest:full-path="Object 1/"');
    expect(manifest).toContain("application/vnd.oasis.opendocument.presentation");
  });

  it("round-trips a body-anchored non-chart object frame", () => {
    const source = generateOcf(
      "application/vnd.oasis.opendocument.text",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none"/></style:style></office:automatic-styles><office:body><office:text><text:p><draw:frame draw:style-name="gr1" draw:name="Formula object" text:anchor-type="char" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p></office:text></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        "Object 1/content.xml":
          '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      },
      {},
      {
        version: "1.3",
        entries: [
          { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
          { fullPath: "content.xml", mediaType: "text/xml" },
          { fullPath: "styles.xml", mediaType: "text/xml" },
          { fullPath: "Object 1/", mediaType: "application/vnd.oasis.opendocument.formula" },
        ],
      },
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.embeddedObjects).toEqual([
      expect.objectContaining({
        path: "Object 1",
        mediaType: "application/vnd.oasis.opendocument.formula",
        frames: [{ shapeId: 1, anchorType: "char", styleName: "gr1" }],
      }),
    ]);
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: {
        children: [
          {
            wpsShape: {
              geometry: "rect",
              transformation: {
                offset: { left: 360000, top: 720000 },
                width: 1440000,
                height: 1080000,
              },
              nonVisualProperties: { id: 1, name: "Formula object" },
              fill: { type: "none" },
              outline: { type: "noFill" },
            },
          },
        ],
      },
    });

    const generated = generateDocument(parsed);
    const content = new TextDecoder().decode(unzipSync(generated)["content.xml"]);
    expect(content).toContain(
      '<draw:frame text:anchor-type="char" draw:style-name="gr1" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm" draw:name="Formula object"><draw:object xlink:href="./Object 1" xlink:type="simple"/></draw:frame>',
    );

    const reread = parseDocument(generated) as OdtDocumentOptions;
    expect(reread.embeddedObjects).toEqual(parsed.embeddedObjects);
    expect(projection(reread.sections[0]?.children?.[0])).toEqual(
      projection(parsed.sections[0]?.children?.[0]),
    );
  });

  it("preserves source-only package members and omits stale overlays", () => {
    const source = generateOcf(
      "application/vnd.oasis.opendocument.text",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles/><office:body><office:text><text:p>Document</text:p></office:text></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        "manifest.rdf": '<rdf:RDF xmlns:rdf="https://example.test/rdf#"/>',
        "audit.bin": new Uint8Array([7, 9]),
      },
      {},
      {
        version: "1.3",
        entries: [
          { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
          { fullPath: "content.xml", mediaType: "text/xml" },
          { fullPath: "styles.xml", mediaType: "text/xml" },
          { fullPath: "manifest.rdf", mediaType: "application/rdf+xml" },
          { fullPath: "audit.bin", mediaType: "application/octet-stream" },
        ],
      },
    );
    const parsed = parseDocument(source) as OdtDocumentOptions;
    expect(parsed.packageMembers).toEqual([
      { path: "audit.bin", mediaType: "application/octet-stream", data: new Uint8Array([7, 9]) },
      {
        path: "manifest.rdf",
        mediaType: "application/rdf+xml",
        data: new TextEncoder().encode('<rdf:RDF xmlns:rdf="https://example.test/rdf#"/>'),
      },
    ]);

    const generated = generateDocument(parsed);
    const entries = unzipSync(generated);
    expect(new TextDecoder().decode(entries["manifest.rdf"])).toContain("<rdf:RDF");
    expect(entries["audit.bin"]).toEqual(new Uint8Array([7, 9]));
    expect(parseDocument(generated) as OdtDocumentOptions).toMatchObject({
      packageMembers: parsed.packageMembers,
    });

    const stale = {
      ...parsed,
      packageMembers: parsed.packageMembers?.filter((member) => member.path !== "manifest.rdf"),
      packageManifest: {
        version: "1.3",
        entries: [
          { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
          { fullPath: "manifest.rdf", mediaType: "application/rdf+xml" },
        ],
      },
    };
    const regenerated = generateDocument(stale);
    const regeneratedEntries = unzipSync(regenerated);
    expect(Object.keys(regeneratedEntries)).not.toContain("manifest.rdf");
    expect(new TextDecoder().decode(regeneratedEntries["META-INF/manifest.xml"])).not.toContain(
      'manifest:full-path="manifest.rdf"',
    );
  });

  it("rejects a standalone object without content.xml", () => {
    const source = generateOcf(
      "application/vnd.oasis.opendocument.text",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles/><office:body><office:text><text:p><draw:frame svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p></office:text></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        "Object 1/media.bin": new Uint8Array([1]),
      },
      {},
      {
        version: "1.3",
        entries: [
          { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
          { fullPath: "content.xml", mediaType: "text/xml" },
          { fullPath: "styles.xml", mediaType: "text/xml" },
          { fullPath: "Object 1/", mediaType: "application/vnd.oasis.opendocument.presentation" },
          { fullPath: "Object 1/media.bin", mediaType: "application/octet-stream" },
        ],
      },
    );
    let error: unknown;
    try {
      parseDocument(source);
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdtParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      name: "draw:object",
      reason: "embedded object subdocument is missing",
    });
  });

  it("round-trips revisions through canonical insertion and deletion children", () => {
    const options: DocumentOptions = {
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
    const parsed = parseDocument(generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections", "settings"]);
    expect(parsed.settings).toEqual({ trackRevisions: true });
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips a simple form control through a canonical form field", () => {
    const field = {
      formField: { name: "Name", textInput: { value: "Value" } },
    };
    const options: DocumentOptions = {
      sections: [{ children: [{ paragraph: { children: [field] } }] }],
    };
    const parsed = parseDocument(generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual(["forms", "sections"]);
    expect(parsed.forms).toEqual([
      { controls: [{ control: "text", name: "Name", value: "Value" }] },
    ]);
  });

  it("round-trips bookmarks, references, variables, sequences, and hidden text", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            { bookmarkStart: { id: 7, name: "target" } },
            {
              paragraph: {
                children: [
                  { bookmarkEnd: { id: 7 } },
                  { simpleField: { instruction: "REF target \\r" } },
                  { bookmarkStart: { id: 8, name: "spaced target" } },
                  { bookmarkEnd: { id: 8 } },
                  { simpleField: { instruction: "REF spaced\\ target \\r" } },
                  { simpleField: { instruction: "SEQ Figure = 1 + 1", cachedValue: "1" } },
                  {
                    simpleField: {
                      instruction: "VARIABLE Total = 1 + 1",
                      cachedValue: "2",
                    },
                  },
                  { simpleField: { instruction: "STYLEREF 2 \\n", cachedValue: "Section" } },
                  { simpleField: { instruction: "= 1+1", cachedValue: "2" } },
                  { simpleField: { instruction: 'IF condition "secret" ""' } },
                  {
                    simpleField: {
                      instruction: 'IF condition "quoted \\"value\\"" ""',
                    },
                  },
                  { simpleField: { instruction: "CITATION article", cachedValue: "Citation" } },
                ],
              },
            },
            { bookmarkEnd: { id: 8 } },
          ],
        },
      ],
    };
    const parsed = parseDocument(generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual([
      "sections",
      "sequenceDeclarations",
      "variableDeclarations",
    ]);
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips paired inline bookmark and reference markers", () => {
    const parsed = parseDocument(
      odtPackage(
        "<text:p>" +
          "<text:bookmark-start xml:id='bookmark-7' text:name='inline target'/>" +
          "<text:reference-mark-start text:name='reference target'/>" +
          "marked" +
          "<text:reference-mark-end text:name='reference target'/>" +
          "<text:bookmark-end text:name='inline target'/>" +
          "</text:p>",
      ),
    );
    const children = parsed.sections[0]?.children[0];
    expect(children).toEqual({
      paragraph: {
        children: [
          { bookmarkStart: { id: 1, name: "inline target" } },
          { bookmarkStart: { id: 2, name: "reference target" } },
          "marked",
          { bookmarkEnd: { id: 2 } },
          { bookmarkEnd: { id: 1 } },
        ],
      },
    });
    expect(projection(parseDocument(generateDocument(parsed)).sections[0]?.children)).toEqual(
      projection(parsed.sections[0]?.children),
    );
  });

  it("round-trips annotations, indexes, bibliography, and shapes", () => {
    const options: DocumentOptions = {
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
    const parsed = parseDocument(generateDocument(options));
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
    expect(
      parseBodyError("<text:p><text:page-number>1</text:page-number></text:p>", "text:page-number")
        .reason,
    ).toContain("no canonical");
    expect(parseBodyError(tracked, "text:format-change").reason).toContain("no canonical");
    expect(
      parseBodyError(
        "<text:sequence-decls><text:sequence-decl text:name='Figure' text:unknown='true'/></text:sequence-decls>",
        "text:sequence-decl",
      ).reason,
    ).toContain("attribute has no canonical");
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

  it("rejects canonical field instructions without ODF representations", () => {
    const section = (children: NonNullable<ParagraphOptions["children"]>): DocumentOptions => ({
      sections: [{ children: [{ paragraph: { children } }] }],
    });
    try {
      generateDocument(section([{ simpleField: { instruction: "DOCPROPERTY Unknown" } }]));
    } catch (cause) {
      expect(cause).toBeInstanceOf(OdtParseError);
      const error = cause as OdtParseError;
      expect(error.part).toBe("content.xml");
      expect(error.name).toBe("text:field");
      expect(error.reason).toBe("unsupported field instruction");
    }
    try {
      generateDocument(section([{ bookmarkEnd: { id: 7 } }]));
    } catch (cause) {
      expect(cause).toBeInstanceOf(OdtParseError);
      const error = cause as OdtParseError;
      expect(error.name).toBe("text:bookmark-end");
      expect(error.reason).toBe("no canonical bookmark start");
      return;
    }
    throw new Error("Expected canonical input without ODF representation to reject generation");
  });

  it("reports unknown body elements with structured errors", () => {
    const error = parseBodyError("<text:unknown/>", "text:unknown");
    expect(error.path).toBe("/office:document-content/office:body/office:text/text:unknown");
    expect(error.reason).toBe("no canonical DocumentOptions mapping");
  });
});
