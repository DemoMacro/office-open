import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  archiveSemanticDiffDetails,
  archiveTagDiffs,
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

  it("normalizes equivalent OOXML boolean attribute tokens", () => {
    const source = new TextEncoder().encode(
      '<root><flag w:val="off"/><math m:val="false"/></root>',
    );
    const output = new TextEncoder().encode('<root><flag w:val="0"/><math m:val="0"/></root>');
    expect(explainSemanticPartDiff("word/example.xml", source, output)).toEqual([]);
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
});
