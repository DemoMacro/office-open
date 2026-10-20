import { unzipSync } from "@office-open/core";
import type { FormControlOptions, OfficeFormsOptions } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import type { ParagraphOptions } from "@office-open/docx";
import { generateOcf, ODF_NAMESPACES } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument, OdtParseError, parseDocument } from "./index";

const FORM_NAMESPACE = 'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"';

function projection(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value));
}

type SectionBodyChild = NonNullable<
  NonNullable<DocumentOptions["sections"]>[number]["children"]
>[number];

type ParagraphRun = NonNullable<ParagraphOptions["children"]>[number];
type RunLikeChild = Extract<ParagraphRun, { children?: unknown }>;

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
    expect(error.name).toBe(name);
    if (reason) expect(error.reason).toContain(reason);
    return error;
  }
  throw new Error(`Expected ${name} to reject without a canonical mapping`);
}

describe("ODT canonical projection", () => {
  it("projects a minimal document without hidden root fields", async () => {
    const parsed = await parseDocument(
      await generateDocument({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
    );
    expect(Object.keys(parsed).sort()).toEqual(["sections"]);
    expect(projection(parsed.sections[0]?.children)).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("round-trips common style identity, relationships, indents, and spacing", async () => {
    const options: DocumentOptions = {
      sections: [],
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
            attributes: { "fo:text-align": "center", "style:contextual-spacing": "true" },
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
      name: "draw:object",
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

  it("rejects canonical field instructions without ODF representations", async () => {
    const section = (children: NonNullable<ParagraphOptions["children"]>): DocumentOptions => ({
      sections: [{ children: [{ paragraph: { children } }] }],
    });
    try {
      await generateDocument(section([{ simpleField: { instruction: "DOCPROPERTY Unknown" } }]));
    } catch (cause) {
      expect(cause).toBeInstanceOf(OdtParseError);
      const error = cause as OdtParseError;
      expect(error.part).toBe("content.xml");
      expect(error.name).toBe("text:field");
      expect(error.reason).toBe("unsupported field instruction");
    }
    try {
      await generateDocument(section([{ bookmarkEnd: { id: 7 } }]));
    } catch (cause) {
      expect(cause).toBeInstanceOf(OdtParseError);
      const error = cause as OdtParseError;
      expect(error.name).toBe("text:bookmark-end");
      expect(error.reason).toBe("no canonical bookmark start");
      return;
    }
    throw new Error("Expected canonical input without ODF representation to reject generation");
  });

  it("reports unknown body elements with structured errors", async () => {
    const error = await parseBodyError("<text:unknown/>", "text:unknown");
    expect(error.path).toBe("/office:document-content/office:body/office:text/text:unknown");
    expect(error.reason).toBe("no canonical DocumentOptions mapping");
  });
});
