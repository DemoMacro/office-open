import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { encryptedPassthroughMatches, isEncryptedPassthrough } from "./library";
import {
  archiveSemanticDiffDetails,
  archiveSemanticDiffs,
  archiveTagDiffs,
  assertEncryptedContainerRoundTrip,
  canonicalXmlNodes,
  classifyPackageFailure,
  explainSemanticPartDiff,
  parseCanonicalXml,
} from "./semantics";

function zip(files: Record<string, string | Uint8Array>): Uint8Array {
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, data]) => [
        name,
        typeof data === "string" ? new TextEncoder().encode(data) : data,
      ]),
    ),
  );
}

describe("corpus semantic comparison", () => {
  it("compares encrypted containers as opaque passthrough payloads", () => {
    const data = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]);
    const options = { encrypted: { data } };

    expect(isEncryptedPassthrough(options)).toBe(true);
    expect(isEncryptedPassthrough({})).toBe(false);
    expect(encryptedPassthroughMatches(data, new Uint8Array(data))).toBe(true);
    expect(encryptedPassthroughMatches(data, data.subarray(0, 4))).toBe(false);
    expect(
      encryptedPassthroughMatches(data, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 4])),
    ).toBe(false);
  });

  it("normalizes document metadata order and namespace dialects", () => {
    const source = parseCanonicalXml(
      '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extended-properties">' +
        "<Application>B</Application><Pages>1</Pages></Properties>",
    );
    const output = parseCanonicalXml(
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/extended-properties">' +
        "<Pages>1</Pages><Application>B</Application></Properties>",
    );
    expect(canonicalXmlNodes(source, "docProps/app.xml")).toEqual(
      canonicalXmlNodes(output, "docProps/app.xml"),
    );
  });

  it("normalizes strict relationship URIs to transitional equivalents", () => {
    const source = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties" Target="docProps/app.xml"/>' +
        "</Relationships>",
    );
    const output = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        "</Relationships>",
    );
    expect(explainSemanticPartDiff("_rels/.rels", source, output)).toEqual([]);
  });

  it("normalizes legacy shared-string vocabulary to transitional SML", () => {
    const source = new TextEncoder().encode(
      '<sst xmlns="http://schemas.microsoft.com/office/excel/2005/8/sst" ' +
        'totalCount="2" uniqueCount="2"><sstItem><t>A</t></sstItem>' +
        "<sstItem><t>B</t></sstItem></sst>",
    );
    const output = new TextEncoder().encode(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'count="2" uniqueCount="2"><si><t>A</t></si><si><t>B</t></si></sst>',
    );
    expect(explainSemanticPartDiff("xl/sharedStrings.xml", source, output)).toEqual([]);
  });

  it("normalizes legacy CT_Color type/value lexemes", () => {
    const source = new TextEncoder().encode(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts><font><color type="theme" val="1"/><color type="rgb" val="FF112233"/>' +
        '<color type="icv" val="9"/></font>' +
        '<fills><fill><pattern patternType="solid"><fgColor type="theme" val="3" tint="0.5"/>' +
        '<bgColor type="icv" val="65"/></pattern></fill></fills>',
    );
    const output = new TextEncoder().encode(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts><font><color theme="1"/><color rgb="FF112233"/><color indexed="9"/></font>' +
        '<fills><fill><pattern patternType="solid"><fgColor theme="3" tint="0.5"/>' +
        '<bgColor indexed="65"/></pattern></fill></fills></styleSheet>',
    );
    expect(explainSemanticPartDiff("xl/styles.xml", source, output)).toEqual([]);
  });

  it("normalizes OnOff toggle lexical variants and omitted values", () => {
    const source = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p><w:pPr><w:keepNext w:val="true"/><w:keepLines w:val="on"/></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:i w:val="t"/></w:rPr></w:r></w:p>' +
        '<w:tbl><w:tr><w:trPr><w:cantSplit w:val="on"/><w:tblHeader/></w:trPr></w:tr></w:tbl>' +
        "</w:body></w:document>",
    );
    const output = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p><w:pPr><w:keepNext/><w:keepLines w:val="1"/></w:pPr>' +
        '<w:r><w:rPr><w:b w:val="1"/><w:i/></w:rPr></w:r></w:p>' +
        '<w:tbl><w:tr><w:trPr><w:cantSplit/><w:tblHeader w:val="on"/></w:trPr></w:tr></w:tbl>' +
        "</w:body></w:document>",
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("fills section margin header/footer writer defaults", () => {
    const source = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="0"/></w:sectPr></w:body></w:document>',
    );
    const output = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="0" w:header="851" w:footer="992"/></w:sectPr></w:body></w:document>',
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes boolean element text lexemes", () => {
    const source = new TextEncoder().encode(
      '<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:o="urn:schemas-microsoft-com:office:office"><o:OLEObject Type="Embed">' +
        "<o:LockedField>false</o:LockedField></o:OLEObject></w:object>",
    );
    const output = new TextEncoder().encode(
      '<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:o="urn:schemas-microsoft-com:office:office"><o:OLEObject Type="Embed">' +
        "<o:LockedField>f</o:LockedField></o:OLEObject></w:object>",
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("folds duplicate singleton toggle elements with last-one-wins", () => {
    const source = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p><w:pPr><w:rPr><w:b w:val="1"/><w:b w:val="1"/>' +
        '<w:bCs w:val="false"/><w:bCs w:val="false"/></w:rPr></w:pPr></w:p></w:body></w:document>',
    );
    const output = new TextEncoder().encode(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        "<w:body><w:p><w:pPr><w:rPr><w:b/>" +
        '<w:bCs w:val="false"/></w:rPr></w:pPr></w:p></w:body></w:document>',
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes legacy Office relationship URIs to transitional equivalents", () => {
    const source = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.microsoft.com/package/2005/06/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/docPropsApp" Target="docProps/app.xml"/>' +
        "</Relationships>",
    );
    const output = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        "</Relationships>",
    );
    expect(explainSemanticPartDiff("_rels/.rels", source, output)).toEqual([]);
  });

  it("normalizes absolute and relative internal relationship targets", () => {
    const source = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="/xl/theme/theme1.xml"/>' +
        "</Relationships>",
    );
    const output = new TextEncoder().encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
        "</Relationships>",
    );
    expect(explainSemanticPartDiff("xl/_rels/workbook.xml.rels", source, output)).toEqual([]);
  });

  it("normalizes versioned strict markup namespaces to transitional equivalents", () => {
    const source = new TextEncoder().encode(
      '<root><graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart"/></root>',
    );
    const output = new TextEncoder().encode(
      '<root><graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></root>',
    );
    expect(explainSemanticPartDiff("ppt/slides/example.xml", source, output)).toEqual([]);
  });

  it("reports source-only elements without comparing raw declarations", () => {
    const source = new TextEncoder().encode("<root><lost/></root>");
    const output = new TextEncoder().encode("<root/>");
    expect(explainSemanticPartDiff("word/example.xml", source, output)).toEqual([
      {
        path: "word/example.xml",
        kind: "xml",
        category: "child",
        xpath: "word/example.xml/lost",
        detail: "source-only lost",
      },
    ]);
  });

  it("normalizes equivalent numeric cell values", () => {
    const source = new TextEncoder().encode(
      "<worksheet><sheetData><row><c><v>1.2300000000000000</v></c></row></sheetData></worksheet>",
    );
    const output = new TextEncoder().encode(
      "<worksheet><sheetData><row><c><v>1.23</v></c></row></sheetData></worksheet>",
    );
    expect(explainSemanticPartDiff("xl/worksheets/example.xml", source, output)).toEqual([]);
  });

  it("normalizes equivalent numeric row heights", () => {
    const source = new TextEncoder().encode('<row ht="17.100000000000001"/>');
    const output = new TextEncoder().encode('<row ht="17.1"/>');
    expect(explainSemanticPartDiff("xl/worksheets/example.xml", source, output)).toEqual([]);
  });

  it("normalizes semantically unordered section and footnote children", () => {
    const source = parseCanonicalXml(
      "<document><body><sectPr><pgSz/><cols/><headerReference/></sectPr></body></document>",
    );
    const output = parseCanonicalXml(
      "<document><body><sectPr><headerReference/><pgSz/><cols/></sectPr></body></document>",
    );
    expect(canonicalXmlNodes(source, "word/document.xml")).toEqual(
      canonicalXmlNodes(output, "word/document.xml"),
    );

    const footnotesSource = parseCanonicalXml(
      "<footnotes><footnote id='2'/><footnote id='1'/></footnotes>",
    );
    const footnotesOutput = parseCanonicalXml(
      "<footnotes><footnote id='1'/><footnote id='2'/></footnotes>",
    );
    expect(canonicalXmlNodes(footnotesSource, "word/footnotes.xml")).toEqual(
      canonicalXmlNodes(footnotesOutput, "word/footnotes.xml"),
    );

    const propertiesSource = parseCanonicalXml("<docPartPr><name/><guid/><category/></docPartPr>");
    const propertiesOutput = parseCanonicalXml("<docPartPr><name/><category/><guid/></docPartPr>");
    expect(canonicalXmlNodes(propertiesSource, "word/glossary/document.xml")).toEqual(
      canonicalXmlNodes(propertiesOutput, "word/glossary/document.xml"),
    );

    const paragraphPropertiesSource = parseCanonicalXml("<pPr><spacing/><keepNext/></pPr>");
    const paragraphPropertiesOutput = parseCanonicalXml("<pPr><keepNext/><spacing/></pPr>");
    expect(canonicalXmlNodes(paragraphPropertiesSource, "word/document.xml")).toEqual(
      canonicalXmlNodes(paragraphPropertiesOutput, "word/document.xml"),
    );
  });

  it("normalizes equivalent OOXML boolean attribute tokens", () => {
    const source = new TextEncoder().encode(
      '<root><flag w:val="off"/><math m:val="false"/></root>',
    );
    const output = new TextEncoder().encode('<root><flag w:val="0"/><math m:val="0"/></root>');
    expect(explainSemanticPartDiff("word/example.xml", source, output)).toEqual([]);
  });

  it("ignores MCE extension attributes when matching unordered children", () => {
    const namespace = 'xmlns:w="urn:w" xmlns:w14="urn:w14" xmlns:mc="urn:mc"';
    const source = new TextEncoder().encode(
      `<root ${namespace} mc:Ignorable="w14">` +
        '<w:pPr><first w14:before="1"/><second/></w:pPr>' +
        "</root>",
    );
    const output = new TextEncoder().encode(
      `<root ${namespace}>` + "<w:pPr><second/><first/></w:pPr>" + "</root>",
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("ignores source-only children from MCE-ignored namespaces", () => {
    const namespace = 'xmlns:w14="urn:w14" xmlns:mc="urn:mc"';
    const source = new TextEncoder().encode(
      `<root ${namespace} mc:Ignorable="w14"><known/><w14:extension/></root>`,
    );
    const output = new TextEncoder().encode(`<root ${namespace}><known/></root>`);
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("ignores MCE processing directives", () => {
    const source = new TextEncoder().encode(
      '<drawing xmlns:mc="urn:mc" mc:MustUnderstand="wps" mc:ProcessContent="wps:wsp"><content/></drawing>',
    );
    const output = new TextEncoder().encode("<drawing><content/></drawing>");
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes omitted inline distances to explicit schema defaults", () => {
    const source = new TextEncoder().encode("<root><wp:inline/></root>");
    const output = new TextEncoder().encode(
      '<root><wp:inline distT="0" distB="0" distL="0" distR="0"/></root>',
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes page measure lexemes to twips", () => {
    const source = new TextEncoder().encode(
      '<root><w:pgSz w:w="612pt" w:h="792pt"/><w:pgMar w:top="72pt" w:right="72pt"/></root>',
    );
    const output = new TextEncoder().encode(
      '<root><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"/></root>',
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes omitted page margins to explicit writer defaults", () => {
    const source = new TextEncoder().encode(
      '<root><w:pgMar w:top="0" w:right="0" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></root>',
    );
    const output = new TextEncoder().encode(
      '<root><w:pgMar w:top="0" w:right="0" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></root>',
    );
    expect(explainSemanticPartDiff("word/document.xml", source, output)).toEqual([]);
  });

  it("normalizes equivalent OOXML bullet percentage lexemes", () => {
    const source = new TextEncoder().encode('<root><a:buSzPct val="75000"/></root>');
    const output = new TextEncoder().encode('<root><a:buSzPct val="75%"/></root>');
    expect(explainSemanticPartDiff("ppt/example.xml", source, output)).toEqual([]);
  });

  it("normalizes relationship references by their part targets", () => {
    const files = {
      "ppt/slides/slide1.xml":
        '<p:root xmlns:p="urn:p" xmlns:r="urn:r"><p:ref r:id="rId2"/></p:root>',
      "ppt/slides/_rels/slide1.xml.rels":
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
    };
    const output = {
      ...files,
      "ppt/slides/slide1.xml": files["ppt/slides/slide1.xml"].replace('r:id="rId2"', 'r:id="rId9"'),
      "ppt/slides/_rels/slide1.xml.rels": files["ppt/slides/_rels/slide1.xml.rels"].replace(
        "rId2",
        "rId9",
      ),
    };
    expect(archiveSemanticDiffDetails(zip(files), zip(output))).toEqual([]);
  });

  it("normalizes relationship targets against their owning part", () => {
    const relationships =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
      "</Relationships>";
    const source = new TextEncoder().encode(relationships);
    const output = new TextEncoder().encode(relationships.replace("../media", "/xl/media"));
    expect(explainSemanticPartDiff("xl/drawings/_rels/drawing1.xml.rels", source, output)).toEqual(
      [],
    );
  });

  it("normalizes redundant identical relationship edges", () => {
    const relationship =
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>';
    const relationships =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    const source = new TextEncoder().encode(
      `${relationships}${relationship}${relationship}</Relationships>`,
    );
    const output = new TextEncoder().encode(`${relationships}${relationship}</Relationships>`);
    expect(explainSemanticPartDiff("xl/drawings/_rels/drawing1.xml.rels", source, output)).toEqual(
      [],
    );
  });

  it("normalizes unprefixed comment anchor corners from compatibility fallbacks", () => {
    const marker = (prefix: string): string =>
      `<${prefix === "xdr:" ? prefix : ""}from>` +
      `<xdr:col>1</xdr:col><xdr:row>1</xdr:row></${prefix === "xdr:" ? `${prefix}from` : "from"}>`;
    const source = new TextEncoder().encode(
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<commentList><comment ref="A1"><commentPr><anchor>${marker("")}</anchor></commentPr></comment></commentList></comments>`,
    );
    const output = new TextEncoder().encode(
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<commentList><comment ref="A1"><commentPr><anchor>${marker("xdr:")}</anchor></commentPr></comment></commentList></comments>`,
    );
    expect(explainSemanticPartDiff("xl/comments1.xml", source, output)).toEqual([]);
  });

  it("reports duplicate children even when their semantic payloads match", () => {
    const source = new TextEncoder().encode("<root><item/></root>");
    const output = new TextEncoder().encode("<root><item/><item/></root>");
    expect(explainSemanticPartDiff("xl/example.rels", source, output)).toEqual([
      {
        path: "xl/example.rels",
        kind: "relationship",
        category: "child",
        xpath: "xl/example.rels/item",
        detail: "output-only item",
      },
    ]);
  });

  it("decodes UTF-16 XML parts by BOM", () => {
    const source = new TextEncoder().encode("<root><item/></root>");
    const utf16 = new TextDecoder().decode(source);
    const bytes = new Uint8Array(utf16.length * 2 + 2);
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    for (let index = 0; index < utf16.length; index++)
      bytes.set([utf16.charCodeAt(index) & 0xff, utf16.charCodeAt(index) >> 8], index * 2 + 2);
    const output = new TextEncoder().encode("<root><item/></root>");
    expect(explainSemanticPartDiff("ppt/slideLayouts/example.xml", bytes, output)).toEqual([]);
  });

  it("classifies malformed packages precisely", () => {
    expect(classifyPackageFailure(new Error("Invalid zip data"))).toBe("invalid-zip");
    expect(classifyPackageFailure(new Error("compound file is truncated"))).toBe("invalid-cfb");
    expect(classifyPackageFailure(new Error("unexpected end of XML"))).toBe("truncated-xml");
    expect(classifyPackageFailure(new Error("primary part not found"))).toBe("invalid-package");
  });

  it("keeps transitional tag counting for the baseline gate", () => {
    const files = { "word/a.xml": "<a><b/><b/></a>", "word/b.xml": "<b/>" };
    expect(archiveTagDiffs(zip(files), zip({ ...files, "word/a.xml": "<a><b/></a>" }))).toEqual([
      "word/a.xml",
    ]);
  });
  it("treats empty relationship parts as semantically optional", () => {
    const files = { "word/document.xml": "<a/>" };
    const source = zip({
      ...files,
      "word/_rels/footnotes.xml.rels":
        '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />',
    });
    expect(archiveTagDiffs(source, zip(files))).toEqual([]);
    expect(archiveSemanticDiffs(source, zip(files))).toEqual([]);
  });
});

describe("encrypted DOCX containers", () => {
  it("requires verbatim passthrough and ignores plain ZIP archives", () => {
    const encrypted = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(() => assertEncryptedContainerRoundTrip(encrypted, encrypted)).not.toThrow();
    expect(() => assertEncryptedContainerRoundTrip(encrypted, encrypted.slice(0, 7))).toThrow(
      "encrypted DOCX container was not re-emitted verbatim",
    );
    expect(() =>
      assertEncryptedContainerRoundTrip(new Uint8Array([0x50, 0x4b]), new Uint8Array([0x50, 0x4c])),
    ).not.toThrow();
  });
});
