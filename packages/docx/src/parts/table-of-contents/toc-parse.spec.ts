import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseToc, parseTocFieldFromElements, parseTocFieldInstruction } from "./toc-parse";

const readCtx = {} as unknown as DocxReadContext;

describe("parseTocFieldInstruction", () => {
  it("parses heading range, hyperlink, and web-view switches", () => {
    const opts: Record<string, unknown> = {};
    parseTocFieldInstruction('TOC \\o "1-3" \\h \\z', opts);
    expect(opts).toEqual({
      headingStyleRange: "1-3",
      hyperlink: true,
      hideTabAndPageNumbersInWebView: true,
    });
  });

  it("parses outline level and custom style-to-level mappings", () => {
    const opts: Record<string, unknown> = {};
    parseTocFieldInstruction('TOC \\u \\t "ChapterTitle,1,Appendix,2"', opts);
    expect(opts.useAppliedParagraphOutlineLevel).toBe(true);
    expect(opts.stylesWithLevels).toEqual([
      { styleName: "ChapterTitle", level: 1 },
      { styleName: "Appendix", level: 2 },
    ]);
  });

  it("ignores non-TOC field instructions", () => {
    const opts: Record<string, unknown> = {};
    parseTocFieldInstruction('HYPERLINK \\l "_Toc123"', opts);
    expect(opts).toEqual({});
  });
});

describe("parseTocFieldFromElements", () => {
  it("extracts TOC switches from a bare field, ignoring rendered HYPERLINK/PAGEREF entries", () => {
    // A bare cross-paragraph TOC field: begin + instrText + separate, then the
    // rendered entries (HYPERLINK/PAGEREF — Word-generated, must be ignored),
    // then end. This is the shape Word writes when no SDT wraps the TOC.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>
        <w:t>Heading One</w:t></w:p>
      <w:p><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const opts = parseTocFieldFromElements(body.elements ?? []);
    expect(opts).toEqual({
      headingStyleRange: "1-3",
      hyperlink: true,
      hideTabAndPageNumbersInWebView: true,
    });
  });
});

describe("parseToc", () => {
  it("parses an SDT-wrapped TOC by docPartGallery", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj>
        <w:alias w:val="Contents"/></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent>
    </w:sdt>`;
    const result = parseToc(parseXml(xml).elements![0], readCtx);
    expect(result?.alias).toBe("Contents");
    expect(result?.headingStyleRange).toBe("1-3");
    expect(result?.hyperlink).toBe(true);
  });
});
