import { strFromU8, unzipSync } from "@office-open/core";
import type { FormControlOptions, OfficeFormsOptions } from "@office-open/core";
import type { DocumentOptions, TableOptions } from "@office-open/docx";
import type { ParagraphOptions } from "@office-open/docx";
import { CHART_MIME, generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import {
  generateDocument,
  OdtGenerateError,
  OdtParseError,
  parseDocument,
  type OdtDocumentOptions,
} from "./index";
import { trackRevisionBlocks } from "./paragraph";

const FORM_NAMESPACE = 'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"';

function projection(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value));
}

type SectionBodyChild = NonNullable<
  NonNullable<DocumentOptions["sections"]>[number]["children"]
>[number];

type ParagraphRun = NonNullable<ParagraphOptions["children"]>[number];
type RunLikeChild = Extract<ParagraphRun, { children?: unknown }>;
type TableRow = TableOptions["rows"][number];

function sectionParagraph(child: SectionBodyChild | undefined): ParagraphOptions | undefined {
  return child && "paragraph" in child && typeof child.paragraph !== "string"
    ? child.paragraph
    : undefined;
}

function sectionRun(child: ParagraphRun | undefined): RunLikeChild | undefined {
  return typeof child === "object" && child !== null && "children" in child ? child : undefined;
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

async function parseBodyError(body: string, name: string, reason?: string): Promise<OdtParseError> {
  try {
    await parseDocument(odtPackage(body));
  } catch (cause) {
    expect(cause).toBeInstanceOf(OdtParseError);
    const error = cause as OdtParseError;
    expect(error.part).toBe("content.xml");
    expect(error.elementName).toBe(name);
    if (reason) expect(error.reason).toContain(reason);
    return error;
  }
  throw new Error(`Expected ${name} to reject without a canonical mapping`);
}

describe("ODT canonical projection", () => {
  it("writes mandatory label alignment for generated list levels", async () => {
    const options = {
      numbering: {
        abstractNumberings: [
          {
            reference: "List",
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1.",
                paragraph: { indent: { left: 720, hanging: 360 } },
              },
            ],
          },
        ],
      },
      sections: [
        {
          children: [
            { paragraph: { numbering: { reference: "List", level: 0 }, children: ["Item"] } },
          ],
        },
      ],
    };
    const archive = await readOcf(
      await generateDocument(options),
      "application/vnd.oasis.opendocument.text",
    );
    const styles = `${archive.files["content.xml"] ?? ""}${archive.files["styles.xml"] ?? ""}`;
    expect(styles).toContain('text:label-followed-by="nothing"');
    await expect(parseDocument(await generateDocument(options))).resolves.toBeTruthy();
  });

  it("projects a minimal document without hidden root fields", async () => {
    const parsed = await parseDocument(
      await generateDocument({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
    );
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("round-trips common style identity, relationships, indents, and spacing", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [{ paragraph: { text: "Item", numbering: { reference: "List", level: 0 } } }],
        },
      ],
      styles: {
        paragraphStyles: [
          {
            id: "P1",
            name: "Title",
            basedOn: "Standard",
            next: "Normal",
            paragraph: {
              indent: { left: 720, right: 360 },
              spacing: { before: 240, after: 120 },
              alignment: "center",
            },
          },
        ],
      },
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(parsed.styles?.paragraphStyles?.[0]).toEqual(options.styles?.paragraphStyles?.[0]);
  });

  it("rejects invalid packages with a structured error", async () => {
    await expect(parseDocument(new Uint8Array([1, 2, 3]))).rejects.toThrow(OdtParseError);
  });

  it("preserves arbitrary ODF numbering formats as canonical custom formats", async () => {
    const styles =
      '<text:list-style style:name="Custom"><text:list-level-style-number text:level="1" style:num-format="か"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:list-level-style-number></text:list-style>';
    const source = odtStyledPackage(
      styles,
      '<text:list text:style-name="Custom"><text:list-item><text:p>Item</text:p></text:list-item></text:list>',
    );
    const parsed = await parseDocument(source);
    expect(parsed.numbering?.abstractNumberings?.[0]?.levels[0]).toMatchObject({
      format: "custom",
      formatOverride: "か",
      suffix: "tab",
    });
    await expect(parseDocument(await generateDocument(parsed))).resolves.toMatchObject({
      numbering: parsed.numbering,
    });
  });

  it("round-trips list label indents without unit drift", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [{ paragraph: { text: "Item", numbering: { reference: "List", level: 0 } } }],
        },
      ],
      numbering: {
        abstractNumberings: [
          {
            reference: "List",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                suffix: "tab",
                paragraph: { indent: { left: 720, hanging: 180 } },
              },
            ],
          },
        ],
      },
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(parsed.numbering).toEqual(options.numbering);
  });

  it("rejects unknown numbering label suffixes structurally", async () => {
    const styles =
      '<text:list-style style:name="Custom"><text:list-level-style-number text:level="1" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="unknown"/></style:list-level-properties></text:list-level-style-number></text:list-style>';
    await expect(
      parseDocument(odtStyledPackage(styles, "<text:p>Item</text:p>")),
    ).rejects.toMatchObject({ reason: "unknown-label-followed-by" });
  });

  it("treats inherited object names as unknown ODF numbering tokens", async () => {
    const styles =
      '<text:list-style style:name="Custom"><text:list-level-style-number text:level="1" style:num-format="constructor"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="constructor"/></style:list-level-properties></text:list-level-style-number></text:list-style>';
    await expect(
      parseDocument(odtStyledPackage(styles, "<text:p>Item</text:p>")),
    ).rejects.toMatchObject({ reason: "unknown-label-followed-by" });
  });

  it("treats inherited object names as unknown drawing presets", async () => {
    const parsed = await parseDocument(
      odtPackage(
        '<draw:custom-shape svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:enhanced-geometry draw:type="constructor"/></draw:custom-shape>',
      ),
    );
    const child = sectionParagraph(parsed.sections[0]?.children[0])?.children?.[0];
    expect(child).toEqual({
      wpsShape: {
        children: [],
        odfGeometryType: "constructor",
        transformation: {
          height: 1440000,
          offset: { x: 360000, y: 720000 },
          width: 1080000,
        },
      },
    });
  });

  it("escapes header field cached values", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [],
          headers: {
            default: [
              {
                paragraph: {
                  children: [{ simpleField: { instruction: "PAGE", cachedValue: "A&B<C>" } }],
                },
              },
            ],
          },
        },
      ],
    };
    const styles = new TextDecoder().decode(
      unzipSync(await generateDocument(options))["styles.xml"]!,
    );
    expect(styles).toContain("A&amp;B&lt;C&gt;");
  });

  it("round-trips ODF form-control shapes with typed references", async () => {
    const parsed = await parseDocument(
      odtPackage(
        `<draw:control draw:control="control1" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>`,
      ),
    );
    expect(parsed.sections[0]?.children[0]).toMatchObject({
      paragraph: { children: [{ wpsShape: { controlId: "control1", geometry: "rect" } }] },
    });
    const regenerated = await generateDocument(parsed);
    expect(Buffer.from(unzipSync(regenerated)["content.xml"]!).toString()).toContain(
      'draw:control="control1"',
    );
    expect((await parseDocument(regenerated)).sections[0]?.children).toEqual(
      parsed.sections[0]?.children,
    );
  });

  it("parses compact LibreOffice time control values", async () => {
    const parsed = await parseDocument(
      odtPackage(
        '<office:forms><form:form><form:time xml:id="time" form:current-value="12120000"/></form:form></office:forms>',
      ),
    );
    expect(parsed.forms).toMatchObject({
      containers: [{ controls: [{ control: "time", xmlId: "time", currentValue: "12120000" }] }],
    });
  });

  it("round-trips mixed canonical and overlay paragraph styles", async () => {
    const source = odtStyledPackage(
      '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center" style:contextual-spacing="true"/></style:style>',
      '<text:p text:style-name="P1">Styled</text:p>',
    );
    const parsed = await parseDocument(source);
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: { text: "Styled", style: "P1", alignment: "center" },
    });
    expect(parsed.styleResources).toEqual([
      {
        name: "P1",
        family: "paragraph",
        properties: [
          {
            name: "style:paragraph-properties",
            attributes: { "fo:text-align": "center", "style:contextual-spacing": true },
          },
        ],
      },
    ]);

    expect(
      (await parseDocument(await generateDocument(parsed))).sections[0]?.children?.[0],
    ).toMatchObject({
      paragraph: { text: "Styled", style: "P1", alignment: "center" },
    });
  });

  it("round-trips typed paragraph style properties", async () => {
    const property = {
      name: "style:paragraph-properties",
      attributes: {
        "fo:line-height": 150,
        "fo:margin-left": 567,
        "fo:widows": 2,
        "style:contextual-spacing": true,
      },
    } as const;
    const options: OdtDocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "Typed", style: "P1" } }] }],
      styleResources: [{ name: "P1", family: "paragraph", properties: [property] }],
    };
    const generated = await generateDocument(options);
    const content = strFromU8(unzipSync(generated)["content.xml"]!);
    expect(content).toContain('fo:line-height="150%"');
    expect(content).toContain('fo:margin-left="1.000125cm"');
    expect(content).toContain('fo:widows="2"');
    expect(content).toContain('style:contextual-spacing="true"');
    expect((await parseDocument(generated)).styleResources).toEqual(options.styleResources);
  });

  it("rejects unknown paragraph style attributes", async () => {
    await expect(
      parseDocument(
        odtStyledPackage(
          '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties officeooo:spacing="1"/></style:style>',
          '<text:p text:style-name="P1">Unknown</text:p>',
        ),
      ),
    ).rejects.toMatchObject({
      elementName: "officeooo:spacing",
      part: "content.xml",
      reason: "attribute has no ODF RNG mapping",
    });
  });

  it("round-trips LibreOffice graphic-properties overlays", async () => {
    const source = odtStyledPackage(
      '<style:style style:name="P1" style:family="paragraph"><loext:graphic-properties draw:fill="none"/><style:paragraph-properties fo:text-align="center"/></style:style>',
      '<text:p text:style-name="P1">Styled</text:p>',
    );
    const parsed = await parseDocument(source);
    expect(parsed.styleResources?.[0]?.properties[0]?.name).toBe("loext:graphic-properties");
    const reread = await parseDocument(await generateDocument(parsed));
    expect(reread.styleResources).toEqual(parsed.styleResources);
  });

  it("round-trips LibreOffice revision and shading style overlays", async () => {
    const source = odtStyledPackage(
      '<style:style style:name="P1" style:family="paragraph">' +
        '<style:paragraph-properties officeooo:paragraph-rsid="00123258" style:page-number="auto"/>' +
        "</style:style>" +
        '<style:style style:name="T1" style:family="text">' +
        '<style:text-properties loext:char-shading-value="0" officeooo:rsid="0003e86a"/>' +
        "</style:style>",
      '<text:p text:style-name="P1"><text:span text:style-name="T1">Styled</text:span></text:p>',
    );
    const parsed = await parseDocument(source);
    expect(parsed.styleResources).toEqual([
      {
        name: "P1",
        family: "paragraph",
        properties: [
          {
            name: "style:paragraph-properties",
            attributes: { "officeooo:paragraph-rsid": "00123258", "style:page-number": "auto" },
          },
        ],
      },
      {
        name: "T1",
        family: "text",
        properties: [
          {
            name: "style:text-properties",
            attributes: { "loext:char-shading-value": "0", "officeooo:rsid": "0003e86a" },
          },
        ],
      },
    ]);
    const reread = await parseDocument(await generateDocument(parsed));
    expect(reread.styleResources).toEqual(parsed.styleResources);
  });

  it("round-trips mixed canonical and overlay run styles", async () => {
    const source = odtStyledPackage(
      '<style:style style:name="C1" style:family="text"><style:text-properties fo:font-weight="bold" fo:background-color="#00ff00"/></style:style>',
      '<text:p><text:span text:style-name="C1">Styled</text:span></text:p>',
    );
    const parsed = await parseDocument(source);
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: { children: [{ text: "Styled", bold: true, style: "C1" }] },
    });

    expect(
      (await parseDocument(await generateDocument(parsed))).sections[0]?.children?.[0],
    ).toMatchObject({
      paragraph: { children: [{ text: "Styled", bold: true, style: "C1" }] },
    });
  });

  it("round-trips ruby annotations with ruby and content styles", async () => {
    const source = odtStyledPackage(
      '<style:style style:name="R1" style:family="ruby"><style:ruby-properties style:ruby-align="distribute-letter"/></style:style>' +
        '<style:style style:name="C1" style:family="text"><style:text-properties fo:font-size="12pt"/></style:style>' +
        '<style:style style:name="C2" style:family="text"><style:text-properties fo:font-size="8pt" fo:language="ja" fo:country="JP"/></style:style>',
      '<text:p><text:ruby text:style-name="R1">' +
        '<text:ruby-base><text:span text:style-name="C1">漢字</text:span></text:ruby-base>' +
        '<text:ruby-text text:style-name="C2">かんじ</text:ruby-text>' +
        "</text:ruby></text:p>",
    );
    const parsed = await parseDocument(source);
    const child = parsed.sections[0]?.children?.[0];
    const paragraph = sectionParagraph(child);
    const run = paragraph?.children?.[0];
    const ruby = sectionRun(run)?.children?.[0];
    expect(ruby).toMatchObject({
      ruby: {
        properties: {
          alignment: "distributeLetter",
          fontSize: 8,
          baseFontSize: 12,
          languageId: "ja-JP",
        },
        style: "R1",
        base: { children: [{ text: "漢字", style: "C1" }] },
        text: { children: [{ text: "かんじ", style: "C2" }] },
      },
    });
    expect(parsed.styleResources?.some((overlay) => overlay.name === "R1")).toBe(true);

    const reread = await parseDocument(await generateDocument(parsed));
    expect(reread.styleResources?.some((overlay) => overlay.name === "R1")).toBe(true);
    const rereadChild = reread.sections[0]?.children?.[0];
    const rereadParagraph = sectionParagraph(rereadChild);
    const rereadRun = rereadParagraph?.children?.[0];
    const rereadRuby = sectionRun(rereadRun)?.children?.[0];
    expect(rereadRuby).toMatchObject({
      ruby: {
        style: "R1",
        base: { children: [{ text: "漢字", style: "C1" }] },
        text: { children: [{ text: "かんじ", style: "C2" }] },
      },
    });
  });

  it("round-trips a linked picture with a parent-relative URL", async () => {
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
    const parsed = await parseDocument(await generateDocument(options));
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

  it("round-trips a missing package image as a linked-only picture", async () => {
    const parsed = await parseDocument(
      odtPackage(
        '<text:p><draw:frame svg:width="4cm" svg:height="3cm" draw:name="Logo"><draw:image xlink:href="Pictures/missing.png"/></draw:frame></text:p>',
      ),
    );
    expect(parsed.sections[0]?.children?.[0]).toMatchObject({
      paragraph: {
        children: [
          {
            picture: {
              type: "png",
              sourceUrl: "Pictures/missing.png",
              transformation: { width: 1440000, height: 1080000 },
            },
          },
        ],
      },
    });
    expect(
      (await parseDocument(await generateDocument(parsed))).sections[0]?.children?.[0],
    ).toEqual(parsed.sections[0]?.children?.[0]);
  });

  it("round-trips deeply nested tables without exhausting the stack", async () => {
    let child: SectionBodyChild = { paragraph: { children: ["Leaf"] } };
    for (let depth = 0; depth < 2_000; depth += 1)
      child = { table: { rows: [{ cells: [{ children: [child] }] }] } };
    const parsed = await parseDocument(
      await generateDocument({ sections: [{ children: [child] }] }),
    );
    let depth = 0;
    let current = parsed.sections[0]?.children?.[0];
    while (current && "table" in current) {
      depth += 1;
      const row: TableRow | undefined = current.table.rows[0];
      const cells = row && "cells" in row ? row.cells : undefined;
      if (!cells) break;
      const cell = cells[0];
      if (!cell || !("children" in cell)) break;
      current = cell.children?.[0];
    }
    expect(depth).toBe(2_000);
  });

  it("round-trips repeated body bookmarks with the same canonical id", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            { bookmarkStart: { id: 0, name: "_GoBack" } },
            { bookmarkEnd: { id: 0 } },
            { bookmarkStart: { id: 0, name: "_GoBack" } },
            { bookmarkEnd: { id: 0 } },
          ],
        },
      ],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("rejects bookmark ids outside the safe integer range", async () => {
    await expect(
      parseDocument(
        odtPackage(
          "<text:bookmark-start text:id='9007199254740993' text:name='unsafe'/>" +
            "<text:bookmark-end text:name='unsafe'/>",
        ),
      ),
    ).rejects.toThrow(OdtParseError);
    await expect(
      parseDocument(
        odtPackage(
          "<text:bookmark-start text:name='unsafe__office-open-id-9007199254740993__'/>" +
            "<text:bookmark-end text:name='unsafe__office-open-id-9007199254740993__'/>",
        ),
      ),
    ).rejects.toThrow(OdtParseError);
    await expect(
      parseDocument(
        odtPackage(
          "<text:bookmark-start text:id='9007199254740991' text:name='maximum'/>" +
            "<text:bookmark-start text:name='overflow'/>" +
            "<text:bookmark-end text:name='overflow'/>",
        ),
      ),
    ).rejects.toThrow(OdtParseError);
  });

  it("reports bookmark diagnostics with the source path", async () => {
    await expect(
      parseDocument(
        odtPackage(
          "<text:p>prefix</text:p>" +
            "<text:p><text:bookmark-start text:name='unsafe__office-open-id-9007199254740993__'/>" +
            "<text:bookmark-end text:name='unsafe__office-open-id-9007199254740993__'/></text:p>",
        ),
      ),
    ).rejects.toMatchObject({
      elementName: "text:bookmark-start",
      path: "/office:body/office:text/text:p/text:bookmark-start",
      reason: "unsafe-bookmark-id",
    });
  });

  it("reports bookmark paths through nested containers", async () => {
    const bookmark = (name: string): string =>
      `<text:p><text:bookmark-start text:name='${name}'/>` +
      `<text:bookmark-end text:name='${name}'/></text:p>`;
    await expect(
      parseDocument(
        odtPackage(
          `<text:section text:name='section'>${bookmark("unsafe__office-open-id-9007199254740993__")}</text:section>`,
        ),
      ),
    ).rejects.toMatchObject({
      path: "/office:body/office:text/text:section/text:p/text:bookmark-start",
    });
    await expect(
      parseDocument(
        odtPackage(
          `<text:list><text:list-item>${bookmark("unsafe__office-open-id-9007199254740993__")}</text:list-item></text:list>`,
        ),
      ),
    ).rejects.toMatchObject({
      path: "/office:body/office:text/text:list/text:list-item/text:p/text:bookmark-start",
    });
    await expect(
      parseDocument(
        odtPackage(
          `<text:p><text:note text:note-class='footnote'><text:note-body>${bookmark("unsafe__office-open-id-9007199254740993__")}</text:note-body></text:note></text:p>`,
        ),
      ),
    ).rejects.toMatchObject({
      path: "/office:body/office:text/text:p/text:note/text:note-body/text:p/text:bookmark-start",
    });
  });

  it("round-trips sender fields", async () => {
    const source = odtPackage(
      "<text:p><text:sender-initials>AI</text:sender-initials><text:sender-city>City</text:sender-city><text:sender-state-or-province>State</text:sender-state-or-province></text:p>",
    );
    const parsed = await parseDocument(source);
    const expected = {
      paragraph: {
        children: [
          { simpleField: { instruction: "AUTHOR \\initials", cachedValue: "AI" } },
          { simpleField: { instruction: "AUTHOR \\city", cachedValue: "City" } },
          { simpleField: { instruction: "AUTHOR \\state-or-province", cachedValue: "State" } },
        ],
      },
    };
    expect(parsed.sections[0]?.children?.[0]).toEqual(expected);
    expect(
      (await parseDocument(await generateDocument(parsed))).sections[0]?.children?.[0],
    ).toEqual(expected);
  });

  it("preserves a standalone embedded object through a typed overlay", async () => {
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
          { fullPath: "Object 1/", mediaType: "application/vnd.oasis.opendocument.formula" },
          { fullPath: "Object 1/content.xml", mediaType: "text/xml" },
        ],
      },
    );
    const parsed = await parseDocument(source);
    expect(parsed.embeddedObjects).toEqual([
      {
        path: "Object 1",
        mediaType: "application/vnd.oasis.opendocument.formula",
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

    const reread = await parseDocument(await generateDocument(parsed));
    expect(reread.embeddedObjects).toEqual(parsed.embeddedObjects);
    const manifest = new TextDecoder().decode(
      unzipSync(await generateDocument(parsed))["META-INF/manifest.xml"],
    );
    const content = new TextDecoder().decode(
      unzipSync(await generateDocument(parsed))["content.xml"],
    );
    expect(content.match(/xlink:href="\.\/Object 1"/g)).toHaveLength(2);
    expect(manifest).toContain('manifest:full-path="Object 1/"');
    expect(manifest).toContain("application/vnd.oasis.opendocument.formula");
  });

  it("round-trips a body-anchored non-chart object frame", async () => {
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
    const parsed = await parseDocument(source);
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
                offset: { x: 360000, y: 720000 },
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

    const generated = await generateDocument(parsed);
    const content = new TextDecoder().decode(unzipSync(generated)["content.xml"]);
    expect(content).toContain(
      '<draw:frame text:anchor-type="char" draw:style-name="gr1" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm" draw:name="Formula object"><draw:object xlink:href="./Object 1" xlink:type="simple"/></draw:frame>',
    );

    const reread = await parseDocument(generated);
    expect(reread.embeddedObjects).toEqual(parsed.embeddedObjects);
    expect(projection(reread.sections[0]?.children?.[0])).toEqual(
      projection(parsed.sections[0]?.children?.[0]),
    );
  });

  it("preserves source-only package members and omits stale overlays", async () => {
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
    const parsed = await parseDocument(source);
    expect(parsed.packageMembers).toEqual([
      { path: "audit.bin", mediaType: "application/octet-stream", data: new Uint8Array([7, 9]) },
      {
        path: "manifest.rdf",
        mediaType: "application/rdf+xml",
        data: new TextEncoder().encode('<rdf:RDF xmlns:rdf="https://example.test/rdf#"/>'),
      },
    ]);

    const generated = await generateDocument(parsed);
    const entries = unzipSync(generated);
    expect(new TextDecoder().decode(entries["manifest.rdf"])).toContain("<rdf:RDF");
    expect(entries["audit.bin"]).toEqual(new Uint8Array([7, 9]));
    expect(await parseDocument(generated)).toMatchObject({
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
    const regenerated = await generateDocument(stale);
    const regeneratedEntries = unzipSync(regenerated);
    expect(Object.keys(regeneratedEntries)).not.toContain("manifest.rdf");
    expect(new TextDecoder().decode(regeneratedEntries["META-INF/manifest.xml"])).not.toContain(
      'manifest:full-path="manifest.rdf"',
    );
  });

  it("rejects a standalone object without content.xml", async () => {
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
      await parseDocument(source);
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdtParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      elementName: "draw:object",
      reason: "embedded object subdocument is missing",
    });
  });

  it("round-trips revisions through canonical insertion and deletion children", async () => {
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
                  {
                    text: "formatted",
                    revision: {
                      id: 3,
                      author: "Author",
                      date: "2026-01-01T00:00:00Z",
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections", "settings"]);
    expect(parsed.settings).toEqual({ trackRevisions: true });
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips a simple form control through a canonical form field", async () => {
    const field = {
      formField: { name: "", textInput: { value: "Value" } },
    };
    const options: DocumentOptions = {
      sections: [{ children: [{ paragraph: { children: [field] } }] }],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual(["forms", "sections"]);
    expect(parsed.forms).toEqual({
      containers: [
        {
          controls: [
            {
              control: "text",
              name: "",
              currentValue: "Value",
              xmlId: "form-control-1",
            },
          ],
        },
      ],
    });
  });

  it("round-trips every ODF form control through the document form root", async () => {
    const controls: FormControlOptions[] = [
      { control: "text", id: "text", xmlId: "text", value: "", currentValue: "", maxLength: 0 },
      { control: "textArea", id: "textArea", xmlId: "textArea", paragraphs: ["Line"] },
      {
        control: "formattedText",
        id: "formattedText",
        xmlId: "formattedText",
        minValue: "0",
        maxValue: "9",
        validation: false,
      },
      {
        control: "number",
        id: "number",
        xmlId: "number",
        value: 0,
        currentValue: 1,
        minValue: 0,
        maxValue: 9,
      },
      { control: "date", id: "date", xmlId: "date", value: "2026-01-31" },
      { control: "time", id: "time", xmlId: "time", value: "01:02:03" },
      { control: "comboBox", id: "comboBox", xmlId: "comboBox", items: [{ label: "" }] },
      {
        control: "listBox",
        id: "listBox",
        xmlId: "listBox",
        options: [{ label: "", currentSelected: false, selected: false, value: "" }],
      },
      {
        control: "checkbox",
        id: "checkbox",
        xmlId: "checkbox",
        currentState: "unknown",
        state: "unknown",
        triState: true,
      },
      { control: "password", id: "password", xmlId: "password", echoChar: "*" },
      { control: "file", id: "file", xmlId: "file", currentValue: "a.odt" },
      { control: "fixedText", id: "fixedText", xmlId: "fixedText", label: "", multiLine: true },
      {
        control: "button",
        id: "button",
        xmlId: "button",
        buttonType: "push",
        defaultButton: false,
        toggle: false,
        focusOnClick: false,
      },
      { control: "image", id: "image", xmlId: "image", buttonType: "url", href: "logo.png" },
      { control: "radio", id: "radio", xmlId: "radio", currentSelected: false, selected: false },
      { control: "frame", id: "frame", xmlId: "frame", label: "Frame" },
      { control: "imageFrame", id: "imageFrame", xmlId: "imageFrame", imageData: "logo.png" },
      { control: "hidden", id: "hidden", xmlId: "hidden", value: "" },
      {
        control: "grid",
        id: "grid",
        xmlId: "grid",
        columns: [{ name: "title", controls: [{ control: "text", id: "title", xmlId: "title" }] }],
      },
      {
        control: "valueRange",
        id: "valueRange",
        xmlId: "valueRange",
        minValue: 0,
        maxValue: 100,
        stepSize: 1,
        pageStepSize: 10,
      },
      { control: "genericControl", id: "genericControl", xmlId: "genericControl" },
    ];
    const forms: OfficeFormsOptions = { containers: [{ name: "Library", controls }] };
    const options: DocumentOptions = {
      forms,
      sections: [{ children: [{ paragraph: { children: [{ text: "Form" }] } }] }],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(parsed.forms).toEqual(forms);
  });

  it("round-trips bookmarks, references, variables, sequences, and hidden text", async () => {
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
    const parsed = await parseDocument(await generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual([
      "sections",
      "sequenceDeclarations",
      "variableDeclarations",
    ]);
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips paired inline bookmark and reference markers", async () => {
    const parsed = await parseDocument(
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
    expect(
      projection((await parseDocument(await generateDocument(parsed))).sections[0]?.children),
    ).toEqual(projection(parsed.sections[0]?.children));
  });

  it("preserves bookmark starts in table-of-contents leading content", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              toc: {
                leading: [{ bookmarkStart: { id: 4, name: "contents" } }],
                trailing: [{ bookmarkEnd: { id: 4 } }],
              },
            },
          ],
        },
      ],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(parsed.sections[0]?.children?.[0]).toEqual({
      toc: {
        alias: undefined,
        entries: [{ bookmarkStart: { id: 4, name: "contents" } }, { bookmarkEnd: { id: 4 } }],
        headingStyleRange: "1-10",
        hyperlink: false,
      },
    });
  });

  it("projects orphan comment ranges as empty annotations", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ commentRangeStart: { id: 2 } }, { commentRangeEnd: { id: 2 } }],
              },
            },
          ],
        },
      ],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(parsed.sections[0]?.children?.[0]).toEqual({
      paragraph: {
        children: [
          { commentRangeStart: { id: 1 } },
          { commentReference: 1 },
          { commentRangeEnd: { id: 1 } },
        ],
      },
    });
    expect(parsed.comments).toEqual([{ id: 1, children: [] }]);
  });

  it("reports unrepresentable orphan comment ranges consistently", async () => {
    const error = await generateDocument({
      sections: [{ children: [{ paragraph: { children: [{ commentRangeEnd: { id: 2 } }] } }] }],
    }).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(OdtGenerateError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: "/office:document-content/office:body/office:text",
      elementName: "commentRange",
      reason: "unrepresentable-orphan-comment-range",
    });
  });

  it("round-trips duplicate bookmark names with distinct canonical ids", async () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  { bookmarkStart: { id: 4, name: "Repeated" } },
                  { bookmarkEnd: { id: 4 } },
                  { bookmarkStart: { id: 5, name: "Repeated" } },
                  { bookmarkEnd: { id: 5 } },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = await parseDocument(await generateDocument(options));
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("round-trips an unsupported Word field through a typed ODF expression", async () => {
    const field = { instruction: "DOCPROPERTY Unknown", cachedValue: "Value" };
    const options: DocumentOptions = {
      sections: [{ children: [{ paragraph: { children: [{ simpleField: field }] } }] }],
    };
    const parsed = await parseDocument(await generateDocument(options));
    const paragraph = sectionParagraph(parsed.sections[0]?.children[0]);
    const child = paragraph?.children?.[0];
    expect(child).toEqual({ simpleField: field });
  });

  it("round-trips annotations, indexes, bibliography, and shapes", async () => {
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
    const parsed = await parseDocument(await generateDocument(options));
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual(options.sections[0]?.children);
  });

  it("rejects mapped constructs that have no canonical equivalent", async () => {
    expect(
      (
        await parseBodyError(
          "<office:forms><form:form><form:button form:id='field1' form:unknown='x'/></form:form></office:forms>",
          "form:unknown",
        )
      ).reason,
    ).toContain("not allowed");
    expect(
      (
        await parseBodyError(
          "<text:bookmark-start text:name='same'/><text:bookmark-start text:name='same'/>",
          "text:bookmark-start",
        )
      ).reason,
    ).toContain("no canonical");
    expect(
      (await parseBodyError("<text:bookmark-end text:name='missing'/>", "text:bookmark-end"))
        .reason,
    ).toContain("no canonical");
    expect(
      (
        await parseBodyError(
          "<text:sequence-decls><text:sequence-decl text:name='Figure' text:unknown='true'/></text:sequence-decls>",
          "text:sequence-decl",
        )
      ).reason,
    ).toContain("attribute has no canonical");
    expect(
      (
        await parseBodyError(
          "<text:p><office:annotation office:display='true'><text:p>shown</text:p></office:annotation></text:p>",
          "office:annotation",
        )
      ).reason,
    ).toContain("no canonical");
    expect(
      (
        await parseBodyError(
          "<text:p><text:hidden-text text:condition='0' text:is-fixed='true'>secret</text:hidden-text></text:p>",
          "text:hidden-text",
        )
      ).reason,
    ).toContain("no canonical");
    expect(
      (
        await parseBodyError(
          "<text:illustration-index><text:illustration-index-source text:index-scope='chapter'/><text:index-body/></text:illustration-index>",
          "text:illustration-index",
        )
      ).reason,
    ).toContain("no canonical");
  });

  it("represents orphan bookmark ends as zero-length ODF bookmarks", async () => {
    const section = (children: NonNullable<ParagraphOptions["children"]>): DocumentOptions => ({
      sections: [{ children: [{ paragraph: { children } }] }],
    });
    const parsed = await parseDocument(
      await generateDocument(section([{ bookmarkEnd: { id: 7 } }])),
    );
    expect(parsed.sections[0]?.children[0]).toEqual({
      paragraph: {
        children: [{ bookmarkStart: { id: 1, name: "bookmark-7" } }, { bookmarkEnd: { id: 1 } }],
      },
    });
  });

  it("creates unique XML IDs for repeated orphan bookmark ends", async () => {
    const content = new TextDecoder().decode(
      unzipSync(
        await generateDocument({
          sections: [
            {
              children: [
                { paragraph: { children: [{ bookmarkEnd: { id: 7 } }] } },
                { paragraph: { children: [{ bookmarkEnd: { id: 7 } }] } },
              ],
            },
          ],
        }),
      )["content.xml"]!,
    );
    expect(content.match(/xml:id="bookmark-7(?:-\d+)?"/g)).toEqual([
      'xml:id="bookmark-7"',
      'xml:id="bookmark-7-2"',
    ]);
  });

  it("reports unknown body elements with structured errors", async () => {
    const error = await parseBodyError("<text:unknown/>", "text:unknown");
    expect(error.path).toBe("/office:document-content/office:body/office:text/text:unknown");
    expect(error.reason).toBe("no canonical DocumentOptions mapping");
  });

  it("rejects tracked paragraph content without a track-change projection", () => {
    expect(() =>
      trackRevisionBlocks("insertion", { id: 1, author: "Author", date: "2026-01-01T00:00:00Z" }, [
        { paragraph: { children: [{ math: { equations: [] } }] } },
      ]),
    ).toThrow("tracked content has no canonical track-change child mapping");
  });

  it("applies tracked block revisions inside custom XML", () => {
    const revision = { id: 1, author: "Author", date: "2026-01-01T00:00:00Z" };
    const tracked = trackRevisionBlocks("deletion", revision, [
      { customXml: { element: "test", children: [{ paragraph: { text: "Text" } }] } },
    ]);
    expect(tracked).toEqual([
      {
        customXml: {
          element: "test",
          children: [
            { paragraph: { children: [{ deletion: { ...revision, children: ["Text"] } }] } },
          ],
        },
      },
    ]);
  });
});

describe("ODT linked and revision boundaries", () => {
  it("round-trips the alphabetical index auto-mark file", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:alphabetical-index-auto-mark-file xlink:type='simple' xlink:href='../terms.sdi'/><text:p>Body</text:p>",
      ),
    );
    expect(parsed.indexAutoMarkFile).toEqual({ href: "../terms.sdi" });
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.indexAutoMarkFile).toEqual(parsed.indexAutoMarkFile);
  });

  it("round-trips a named section's external source", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:section text:name='Linked'><text:section-source xlink:type='simple' xlink:href='../source.odt' text:filter-name='writer8' text:section-name='Source'/><text:p>Inside</text:p></text:section><text:p>After</text:p>",
      ),
    );
    expect(parsed.sectionSources).toEqual([
      {
        href: "../source.odt",
        filterName: "writer8",
        sectionName: "Linked",
        targetSectionName: "Source",
      },
    ]);
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.sectionSources).toEqual(parsed.sectionSources);
  });

  it("keeps a format change open across paragraph boundaries", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:tracked-changes text:track-changes='false'><text:changed-region text:id='fmt1'><text:format-change><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes><text:p><text:change-start text:change-id='fmt1'/><text:span>One</text:span></text:p><text:p><text:span>Two</text:span><text:change-end text:change-id='fmt1'/></text:p>",
      ),
    );
    const revisionChildren = parsed.sections[0]?.children ?? [];
    const firstRun = sectionParagraph(revisionChildren[0])?.children?.[0] as {
      revision?: unknown;
    };
    const secondRun = sectionParagraph(revisionChildren[1])?.children?.[0] as {
      revision?: unknown;
    };
    expect(firstRun.revision).toMatchObject({ author: "Author", id: 1 });
    expect(secondRun.revision).toMatchObject({ author: "Author", id: 1 });
  });

  it("round-trips shared ODF table calculation declarations", async () => {
    const content =
      "<table:calculation-settings table:case-sensitive='false' table:precision-as-shown='true'" +
      " table:search-criteria-must-apply-to-whole-cell='false' table:automatic-find-labels='true'" +
      " table:use-regular-expressions='true' table:use-wildcards='false' table:null-year='1930'>" +
      "<table:null-date table:value-type='date' table:date-value='1904-01-01'/>" +
      "<table:iteration table:status='enable' table:steps='20' table:maximum-difference='0.001'/>" +
      "</table:calculation-settings><text:p>Body</text:p>";
    const parsed = await parseDocument(odtPackage(content));
    expect(parsed.calculationSettings).toEqual({
      caseSensitive: false,
      precisionAsShown: true,
      searchCriteriaMustApplyToWholeCell: false,
      automaticFindLabels: true,
      regularExpressionSearch: true,
      wildcardSearch: false,
      nullYear: 1930,
      nullDate: { year: 1904, month: 1, day: 1 },
      iteration: { enabled: true, steps: 20, maximumDifference: 0.001 },
    });
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.calculationSettings).toEqual(parsed.calculationSettings);
  });

  it("round-trips ODF value declarations outside DOCX variable types", async () => {
    const body =
      "<text:variable-decls><text:variable-decl text:name='ReportDate' office:value-type='date'/>" +
      "</text:variable-decls><text:p><text:variable-set text:name='ReportDate' office:value-type='date'" +
      " office:date-value='2016-10-05T00:00:00.0Z' text:formula='ooow:Date(2016;10;05)' text:display='none'/></text:p>";
    const parsed = await parseDocument(odtPackage(body));
    expect(parsed.odtVariableDeclarations).toEqual([{ name: "ReportDate", valueType: "date" }]);
    const paragraph = sectionParagraph(parsed.sections[0]?.children[0]);
    expect(paragraph?.children?.[0]).toMatchObject({
      simpleField: { cachedValue: undefined },
    });
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.odtVariableDeclarations).toEqual(parsed.odtVariableDeclarations);
    expect(
      (
        await parseBodyError(
          body.replace("office:value-type='date'/>", "office:value-type='float'/>"),
          "ReportDate",
          "invalid-variable-value-type",
        )
      ).reason,
    ).toBe("invalid-variable-value-type");
  });

  it("round-trips a body-level format change range", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:tracked-changes text:track-changes='false'><text:changed-region text:id='fmt1'><text:format-change><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes><text:change-start text:change-id='fmt1'/><text:p>One</text:p><text:p>Two</text:p><text:change-end text:change-id='fmt1'/><text:p>Three</text:p>",
      ),
    );
    const children = parsed.sections[0]?.children ?? [];
    for (const index of [0, 1]) {
      const paragraph = sectionParagraph(children[index]);
      expect(paragraph?.children?.[0]).toMatchObject({
        text: index === 0 ? "One" : "Two",
        revision: { author: "Author", id: 1 },
      });
    }
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.sections[0]?.children.slice(0, 2)).toEqual(
      parsed.sections[0]?.children.slice(0, 2),
    );
  });

  it("tracks insertion and deletion ranges across paragraphs and tables", async () => {
    const revision = {
      author: "Author",
      date: "2026-01-01T00:00:00Z",
    };
    const parsed = await parseDocument(
      odtPackage(
        `<text:tracked-changes><text:changed-region text:id='ins1'><text:insertion><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info></text:insertion></text:changed-region><text:changed-region text:id='del1'><text:deletion><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info><text:p>Removed</text:p></text:deletion></text:changed-region></text:tracked-changes>` +
          "<text:change-start text:change-id='ins1'/><text:p>Before</text:p><table:table><table:table-column/><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table><text:change-end text:change-id='ins1'/>" +
          "<text:change-start text:change-id='del1'/><text:p>After</text:p><text:change-end text:change-id='del1'/>",
      ),
    );
    const children = parsed.sections[0]?.children ?? [];
    const table = children[1] && "table" in children[1] ? children[1].table : undefined;
    const cell = table?.rows[0] && "cells" in table.rows[0] ? table.rows[0].cells[0] : undefined;
    const cellParagraph =
      cell && "children" in cell ? sectionParagraph(cell.children[0]) : undefined;
    expect(sectionParagraph(children[0])?.children?.[0]).toMatchObject({
      insertion: { ...revision, children: ["Before"] },
    });
    expect(cellParagraph?.children?.[0]).toMatchObject({
      insertion: { ...revision, children: ["Cell"] },
    });
    expect(sectionParagraph(children[2])?.children?.[0]).toMatchObject({
      deletion: { ...revision, children: ["After"] },
    });
  });

  it("preserves empty paragraphs and hyperlinks inside format-change ranges", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:tracked-changes><text:changed-region text:id='fmt1'><text:format-change><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes>" +
          "<text:change-start text:change-id='fmt1'/><text:p/><text:p><text:a xlink:type='simple' xlink:href='https://example.com/'><text:span>Link</text:span></text:a></text:p><text:change-end text:change-id='fmt1'/>",
      ),
    );
    const children = parsed.sections[0]?.children ?? [];
    expect(sectionParagraph(children[0])?.revision).toMatchObject({ author: "Author", id: 1 });
    const link = sectionParagraph(children[1])?.children?.[0] as {
      hyperlink?: { children?: { revision?: unknown }[] };
    };
    expect(link.hyperlink?.children?.[0]?.revision).toMatchObject({ author: "Author", id: 1 });
    const regenerated = await parseDocument(await generateDocument(parsed));
    expect(regenerated.sections[0]?.children.slice(0, 2)).toEqual(
      parsed.sections[0]?.children.slice(0, 2),
    );
  });

  it("tracks a change that starts inside and ends inside later paragraphs", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:tracked-changes><text:changed-region text:id='ins1'><text:insertion><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info></text:insertion></text:changed-region></text:tracked-changes>" +
          "<text:p>Before<text:change-start text:change-id='ins1'/><text:span>Inserted</text:span></text:p><text:p><text:span>Inside</text:span><text:change-end text:change-id='ins1'/>After</text:p>",
      ),
    );
    const children = parsed.sections[0]?.children ?? [];
    expect(sectionParagraph(children[0])?.text).toBe("Before");
    expect(sectionParagraph(children[1])?.children?.[0]).toMatchObject({
      insertion: {
        author: "Author",
        children: [{ text: "Inserted" }],
      },
    });
    expect(sectionParagraph(children[2])?.children?.[0]).toMatchObject({
      insertion: {
        author: "Author",
        children: [{ text: "Inside" }],
      },
    });
    expect(sectionParagraph(children[3])?.text).toBe("After");
  });

  it("retains every deleted paragraph from headings and lists", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:tracked-changes><text:changed-region text:id='del1'><text:deletion><office:change-info><dc:creator>Author</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info><text:h>Heading</text:h><text:list><text:list-item><text:p>Item</text:p></text:list-item></text:list></text:deletion></text:changed-region></text:tracked-changes>" +
          "<text:p>Keep<text:change text:change-id='del1'/></text:p>",
      ),
    );
    const children = sectionParagraph(parsed.sections[0]?.children[0])?.children ?? [];
    expect(children).toHaveLength(3);
    expect(children[1]).toMatchObject({ deletion: { children: ["Heading"] } });
    expect(children[2]).toMatchObject({ deletion: { children: ["Item"] } });
  });

  it("pairs nested fieldmark boundaries before consuming a separator", async () => {
    const parsed = await parseDocument(
      odtPackage(
        "<text:p><field:fieldmark-start field:type='vnd.oasis.opendocument.field.UNHANDLED'><field:param field:name='vnd.oasis.opendocument.field.code' field:value='A'/></field:fieldmark-start><field:fieldmark-start field:type='vnd.oasis.opendocument.field.UNHANDLED'><field:param field:name='vnd.oasis.opendocument.field.code' field:value='B'/></field:fieldmark-start><field:fieldmark-separator/><field:fieldmark-end/><field:fieldmark-separator/><field:fieldmark-end/></text:p>",
      ),
    );
    const children = sectionParagraph(parsed.sections[0]?.children[0])?.children ?? [];
    const complex = children[0] as {
      complexField?: {
        instruction?: string;
        instructionMembers?: { simpleField?: { instruction?: string } }[];
      };
    };
    expect(complex.complexField?.instruction).toBe("A");
    expect(
      complex.complexField?.instructionMembers?.map((member) => member.simpleField?.instruction),
    ).toEqual(["B"]);
  });

  it("keeps an empty chart subdocument as a typed opaque embedded object", async () => {
    const archive = generateOcf(
      "application/vnd.oasis.opendocument.text",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:text><text:p><draw:frame svg:width='1cm' svg:height='1cm'><draw:object xlink:href='Object 1'/></draw:frame></text:p></office:text></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
        "Object 1/content.xml":
          '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:chart/></office:body></office:document-content>',
      },
      { "Object 1/": CHART_MIME },
    );
    const parsed = await parseDocument(archive);
    expect(sectionParagraph(parsed.sections[0]?.children[0])?.children).toHaveLength(1);
    expect(parsed.embeddedObjects?.map((object) => object.path)).toEqual(["Object 1"]);
  });
});
