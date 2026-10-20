import { describe, expect, it } from "vite-plus/test";

import { partContentTypeResolver } from "./content-types-input";
import { opaquePassthroughPolicy, PACKAGE_PART_MANIFESTS } from "./package-parts";

describe("opaque part policy", () => {
  it("allows only manifest-declared opaque or standalone parts", () => {
    const docx = opaquePassthroughPolicy("docx");
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/vbaProject.bin"))).toBe(true);
    expect(docx.opaquePatterns.some((pattern) => pattern.test("customXml/item1.xml"))).toBe(true);
    expect(docx.opaquePatterns.some((pattern) => pattern.test("docProps/thumbnail.wmf"))).toBe(
      true,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/embeddings/ole.bin"))).toBe(
      true,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/activeX/activeX1.bin"))).toBe(
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

  it("keeps policy prefixes in one manifest per package", () => {
    expect(PACKAGE_PART_MANIFESTS.docx.opaquePrefixes).toContain("media/");
    expect(PACKAGE_PART_MANIFESTS.docx.opaquePrefixes).toContain("word/printerSettings/");
    expect(PACKAGE_PART_MANIFESTS.docx.opaquePrefixes).toContain("customXml/");
    expect(PACKAGE_PART_MANIFESTS.docx.opaquePrefixes).toContain("docProps/thumbnail.");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/embeddings/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/media/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/ddp/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ddp/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/activeX/activeX");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/printerSettings/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/ink/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/fonts/");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/drawings/legacyDiagramText");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("ppt/legacyDocTextInfo.bin");
    expect(PACKAGE_PART_MANIFESTS.pptx.opaquePrefixes).toContain("docProps/thumbnail.");
    expect(PACKAGE_PART_MANIFESTS.xlsx.opaquePrefixes).toContain("xl/printerSettings/");
  });

  it("allows standalone PowerPoint revision metadata", () => {
    const pptx = opaquePassthroughPolicy("pptx");
    expect(pptx.opaquePatterns.some((pattern) => pattern.test("ppt/revisionInfo.xml"))).toBe(true);
    expect(pptx.opaquePatterns.some((pattern) => pattern.test("ppt/ink/ink1.xml"))).toBe(true);
    expect(pptx.opaquePatterns.some((pattern) => pattern.test("ddp/ddpfile.xen"))).toBe(true);
  });

  it("resolves printer settings parts to their family content types", () => {
    const docx = partContentTypeResolver(PACKAGE_PART_MANIFESTS.docx)(
      "word/printerSettings/printerSettings1.bin",
    );
    expect(docx).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings",
    );
    const pptx = partContentTypeResolver(PACKAGE_PART_MANIFESTS.pptx)(
      "ppt/printerSettings/printerSettings1.bin",
    );
    expect(pptx).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.printerSettings",
    );
    const xlsx = partContentTypeResolver(PACKAGE_PART_MANIFESTS.xlsx)(
      "xl/printerSettings/printerSettings1.bin",
    );
    expect(xlsx).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
    );
  });
});
