import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
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

  it("reports source-only elements without comparing raw declarations", () => {
    const source = new TextEncoder().encode("<root><lost/></root>");
    const output = new TextEncoder().encode("<root/>");
    expect(explainSemanticPartDiff("word/example.xml", source, output)).toEqual([
      {
        path: "word/example.xml",
        kind: "xml",
        category: "child",
        xpath: "word/example.xml/lost",
        detail: "source-only child",
      },
    ]);
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
