import { describe, expect, it } from "vite-plus/test";

import { resolverFromRegistry } from "./content-types-input";
import { opaquePassthroughPolicy, PART_REGISTRIES } from "./part-registry";

describe("opaque part policy", () => {
  it("allows only registry-declared opaque or standalone parts", () => {
    const docx = opaquePassthroughPolicy("docx");
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/vbaProject.bin"))).toBe(true);
    expect(docx.opaquePatterns.some((pattern) => pattern.test("customXml/item1.xml"))).toBe(true);
    expect(docx.opaquePatterns.some((pattern) => pattern.test("docProps/thumbnail.wmf"))).toBe(
      true,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/embeddings/ole.bin"))).toBe(
      true,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/theme/theme1.xml"))).toBe(
      false,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/vbaData.xml"))).toBe(true);
  });

  it("expands repeated opaque part templates", () => {
    const xlsx = opaquePassthroughPolicy("xlsx");
    expect(xlsx.opaquePatterns.some((pattern) => pattern.test("xl/vbaProject.bin"))).toBe(true);
  });

  it("keeps policy prefixes in one registry per package", () => {
    expect(PART_REGISTRIES.docx.opaquePrefixes).toContain("media/");
    expect(PART_REGISTRIES.docx.opaquePrefixes).toContain("word/printerSettings/");
    expect(PART_REGISTRIES.docx.opaquePrefixes).toContain("customXml/");
    expect(PART_REGISTRIES.docx.opaquePrefixes).toContain("docProps/thumbnail.");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/embeddings/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/media/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/printerSettings/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/ink/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("docProps/thumbnail.");
    expect(PART_REGISTRIES.xlsx.opaquePrefixes).toContain("xl/printerSettings/");
  });

  it("allows standalone PowerPoint revision metadata", () => {
    const pptx = opaquePassthroughPolicy("pptx");
    expect(pptx.opaquePatterns.some((pattern) => pattern.test("ppt/revisionInfo.xml"))).toBe(true);
    expect(pptx.opaquePatterns.some((pattern) => pattern.test("ppt/ink/ink1.xml"))).toBe(true);
  });

  it("resolves printer settings parts to their family content types", () => {
    const docx = resolverFromRegistry(PART_REGISTRIES.docx)(
      "word/printerSettings/printerSettings1.bin",
    );
    expect(docx).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings",
    );
    const pptx = resolverFromRegistry(PART_REGISTRIES.pptx)(
      "ppt/printerSettings/printerSettings1.bin",
    );
    expect(pptx).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.printerSettings",
    );
    const xlsx = resolverFromRegistry(PART_REGISTRIES.xlsx)(
      "xl/printerSettings/printerSettings1.bin",
    );
    expect(xlsx).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
    );
  });
});
