import { unzipSync, zipSync } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import { generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";
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

const GLOSSARY_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const packageDocumentXml =
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>';
const packageGlossaryXml =
  '<?xml version="1.0"?><w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  "<w:docParts/></w:glossaryDocument>";
const customGlossaryPackage = () =>
  zipSync({
    "word/document.xml": new TextEncoder().encode(packageDocumentXml),
    "word/_rels/document.xml.rels": new TextEncoder().encode(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${GLOSSARY_RELATIONSHIP_TYPE}" Target="custom/glossary.xml"/>` +
        "</Relationships>",
    ),
    "word/custom/glossary.xml": new TextEncoder().encode(packageGlossaryXml),
  });

describe("glossaryDesc round-trip", () => {
  it("round-trips explicit false doc-part metadata", () => {
    const result = roundTrip({
      parts: [
        {
          name: "ExplicitFalse",
          gallery: "default",
          types: ["normal"],
          allTypes: false,
          decorated: false,
          sections: [],
        },
      ],
    });
    expect(result.parts[0]?.decorated).toBe(false);
    expect(result.parts[0]?.allTypes).toBe(false);

    const xml = glossaryDesc.stringify(result, writeCtx)!;
    expect(xml).toContain('w:decorated="0"');
    expect(xml).toContain('<w:types w:all="0">');
  });

  it("round-trips a custom glossary part target", () => {
    const options = parseDocumentSync(customGlossaryPackage());
    expect(options.glossary?.partName).toBe("custom/glossary.xml");

    const output = unzipSync(generateDocumentSync(options, { type: "uint8array" }));
    expect(output["word/custom/glossary.xml"]).toBeDefined();
    expect(new TextDecoder().decode(output["word/_rels/document.xml.rels"]!)).toContain(
      'Target="custom/glossary.xml"',
    );
    expect(new TextDecoder().decode(output["[Content_Types].xml"]!)).toContain(
      '<Override PartName="/word/custom/glossary.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>',
    );

    const regenerated = parseDocumentSync(generateDocumentSync(options, { type: "uint8array" }));
    expect(regenerated.glossary).toEqual(options.glossary);
  });

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
    expect(result.hasDocParts).toBe(true);
    expect(result.parts).toHaveLength(0);
  });

  it("round-trips an absent doc-parts container", () => {
    const doc = parseXml(
      '<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:glossaryDocument>',
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const result = glossaryDesc.parse(root, readCtx);
    expect(result.hasDocParts).toBe(false);
    expect(glossaryDesc.stringify(result, writeCtx)).not.toContain("<w:docParts");
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

describe("glossary companion parts", () => {
  const GLOSSARY_COMPANION_REL_TYPE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const hyperlinkType = `${GLOSSARY_COMPANION_REL_TYPE}/hyperlink`;
  const numberingType = `${GLOSSARY_COMPANION_REL_TYPE}/numbering`;
  const settingsType = `${GLOSSARY_COMPANION_REL_TYPE}/settings`;
  const stylesType = `${GLOSSARY_COMPANION_REL_TYPE}/styles`;
  const webSettingsType = `${GLOSSARY_COMPANION_REL_TYPE}/webSettings`;
  const fontTableType = `${GLOSSARY_COMPANION_REL_TYPE}/fontTable`;
  const fontType = `${GLOSSARY_COMPANION_REL_TYPE}/font`;
  const odttfBytes = new Uint8Array(40).fill(7);

  const companionPackage = () => {
    const enc = new TextEncoder();
    const docPartXml =
      '<?xml version="1.0"?>' +
      `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      "<w:docParts><w:docPart>" +
      '<w:docPartPr><w:name w:val="Companion"/><w:category><w:gallery w:val="default"/></w:category></w:docPartPr>' +
      '<w:docPartBody><w:p><w:hyperlink r:id="rId9"><w:r><w:t>L</w:t></w:r></w:hyperlink></w:p></w:docPartBody>' +
      "</w:docPart></w:docParts></w:glossaryDocument>";
    const relsXml = (rels: string) =>
      '<?xml version="1.0"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels +
      "</Relationships>";
    return zipSync({
      "word/document.xml": enc.encode(packageDocumentXml),
      "word/_rels/document.xml.rels": enc.encode(
        relsXml(
          `<Relationship Id="rId1" Type="${GLOSSARY_RELATIONSHIP_TYPE}" Target="glossary/document.xml"/>`,
        ),
      ),
      "word/glossary/document.xml": enc.encode(docPartXml),
      "word/glossary/_rels/document.xml.rels": enc.encode(
        relsXml(
          `<Relationship Id="rId9" Type="${hyperlinkType}" Target="https://example.test/" TargetMode="External"/>` +
            `<Relationship Id="rId1" Type="${numberingType}" Target="numbering.xml"/>` +
            `<Relationship Id="rId2" Type="${settingsType}" Target="settings.xml"/>` +
            `<Relationship Id="rId3" Type="${stylesType}" Target="styles.xml"/>` +
            `<Relationship Id="rId4" Type="${webSettingsType}" Target="webSettings.xml"/>` +
            `<Relationship Id="rId5" Type="${fontTableType}" Target="fontTable.xml"/>`,
        ),
      ),
      "word/glossary/numbering.xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>' +
          '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
          '<w:lvlText w:val="%1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>' +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
      ),
      "word/glossary/settings.xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:defaultTabStop w:val="720"/></w:settings>',
      ),
      "word/glossary/styles.xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>" +
          '<w:style w:type="paragraph" w:styleId="GH"><w:name w:val="Glossary Heading"/></w:style></w:styles>',
      ),
      "word/glossary/webSettings.xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      ),
      "word/glossary/fontTable.xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<w:font w:name="Emb"><w:embedRegular r:id="rId1" ' +
          'w:fontKey="{12345678-1234-1234-1234-123456789ABC}"/></w:font></w:fonts>',
      ),
      "word/glossary/_rels/fontTable.xml.rels": enc.encode(
        relsXml(`<Relationship Id="rId1" Type="${fontType}" Target="fonts/emb.odttf"/>`),
      ),
      "word/glossary/fonts/emb.odttf": odttfBytes,
      "[Content_Types].xml": enc.encode(
        '<?xml version="1.0"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>' +
          "</Types>",
      ),
    });
  };

  it("absorbs glossary companions into the model and rebuilds them", () => {
    const options = parseDocumentSync(companionPackage());
    const glossary = options.glossary;
    expect(glossary?.numberingPartName).toBe("glossary/numbering.xml");
    expect(glossary?.settingsPartName).toBe("glossary/settings.xml");
    expect(glossary?.stylesPartName).toBe("glossary/styles.xml");
    expect(glossary?.webSettingsPartName).toBe("glossary/webSettings.xml");
    expect(glossary?.fontTablePartName).toBe("glossary/fontTable.xml");
    expect(glossary?.settings?.defaultTabStop).toBe(720);
    expect(glossary?.styles?.paragraphStyles?.[0]?.id).toBe("GH");
    expect(glossary?.webSettings).toEqual({});
    expect(glossary?.numbering?.abstractNumberings?.length).toBe(1);
    expect(glossary?.fonts?.[0]?.name).toBe("Emb");
    expect(glossary?.fonts?.[0]?.data).toBeDefined();

    const output = unzipSync(generateDocumentSync(options, { type: "uint8array" }));
    const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
    expect(output["word/glossary/settings.xml"]).toBeDefined();
    expect(output["word/glossary/styles.xml"]).toBeDefined();
    expect(output["word/glossary/webSettings.xml"]).toBeDefined();
    expect(output["word/glossary/fontTable.xml"]).toBeDefined();
    expect(output["word/glossary/numbering.xml"]).toBeDefined();
    expect(output["word/glossary/fonts/emb.odttf"]).toEqual(odttfBytes);

    const glossaryRels = decode(output["word/glossary/_rels/document.xml.rels"]!);
    expect(glossaryRels).toContain('Id="rId9"');
    expect(glossaryRels).toContain(`Type="${hyperlinkType}"`);
    expect(glossaryRels).toContain('Target="https://example.test/" TargetMode="External"');
    for (const [type, target] of [
      [numberingType, "numbering.xml"],
      [settingsType, "settings.xml"],
      [stylesType, "styles.xml"],
      [webSettingsType, "webSettings.xml"],
      [fontTableType, "fontTable.xml"],
    ] as const) {
      expect(glossaryRels).toContain(`Type="${type}"`);
      expect(glossaryRels).toContain(`Target="${target}"`);
    }
    expect(decode(output["word/glossary/_rels/fontTable.xml.rels"]!)).toContain(
      'Target="fonts/emb.odttf"',
    );
    const contentTypes = decode(output["[Content_Types].xml"]!);
    expect(contentTypes).toContain(
      '<Override PartName="/word/glossary/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    );
    expect(contentTypes).toContain(
      '<Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    );
    expect(contentTypes).toContain(
      '<Override PartName="/word/glossary/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>',
    );
    expect(contentTypes).toContain(
      '<Override PartName="/word/glossary/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
    );

    const regenerated = parseDocumentSync(generateDocumentSync(options, { type: "uint8array" }));
    expect(regenerated.glossary).toEqual(options.glossary);
  });
});
