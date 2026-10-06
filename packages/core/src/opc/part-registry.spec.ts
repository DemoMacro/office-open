import { describe, expect, it } from "vite-plus/test";

import { opaquePassthroughPolicy, PART_REGISTRIES } from "./part-registry";

describe("opaque part policy", () => {
  it("allows only registry-declared opaque binaries", () => {
    const docx = opaquePassthroughPolicy("docx");
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/vbaProject.bin"))).toBe(true);
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/embeddings/ole.bin"))).toBe(
      true,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/theme/theme1.xml"))).toBe(
      false,
    );
    expect(docx.opaquePatterns.some((pattern) => pattern.test("word/vbaData.xml"))).toBe(false);
  });

  it("expands repeated opaque part templates", () => {
    const xlsx = opaquePassthroughPolicy("xlsx");
    expect(xlsx.opaquePatterns.some((pattern) => pattern.test("xl/vbaProject.bin"))).toBe(true);
  });

  it("keeps policy prefixes in one registry per package", () => {
    expect(PART_REGISTRIES.docx.opaquePrefixes).toContain("word/printerSettings/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/embeddings/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("ppt/media/");
    expect(PART_REGISTRIES.pptx.opaquePrefixes).toContain("docProps/thumbnail.");
    expect(PART_REGISTRIES.xlsx.opaquePrefixes).toContain("xl/printerSettings/");
  });
});
