import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseBody, parseSectionChild } from "../../parse/body";
import { stringifyTableOfContents } from "./descriptor";
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

  it("captures rendered entries inside an SDT-wrapped TOC", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Heading Two</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent>
    </w:sdt>`;
    const result = parseToc(parseXml(xml).elements![0], readCtx, (els, ctx) =>
      els.map((e) => parseSectionChild(e, ctx)),
    );
    expect(result?.entries).toHaveLength(2);
  });
});

describe("parseBody TOC page-break rescue", () => {
  it("preserves a page break that shares the TOC field-end paragraph", () => {
    // Bare TOC field whose end paragraph also carries a page break (the
    // section break before the first heading) — must not be swallowed.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:br w:type="page"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const sections = parseBody(body, readCtx);
    const json = JSON.stringify(sections[0].children ?? []);
    expect(json).toContain('"toc"');
    expect(json).toContain('"pageBreak"');
  });
});

describe("parseBody TOC entry preservation", () => {
  it("captures rendered entries between separate and end as `entries`", () => {
    // A TOC field whose separate→end region carries rendered entry paragraphs.
    // These must be captured structurally so MS Office and WPS both display the
    // existing TOC (field emitted clean — no dirty/regeneration prompt).
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Heading Two</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0].children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
  });

  it("omits entries when the TOC field has no rendered result", () => {
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0].children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toBeUndefined();
  });

  it("captures the first entry when separate shares its paragraph", () => {
    // Word sometimes emits the `separate` fldChar in the same paragraph as the
    // first rendered entry (not in the field-head paragraph). The depth tracker
    // must still capture that entry — not drop it as a non-entry.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r><w:t>Heading Two</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0].children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
  });

  it("keeps a nested HYPERLINK field inside an entry from fooling depth tracking", () => {
    // Real Word TOC entries wrap text+page in a nested HYPERLINK field (and
    // often a PAGEREF). The depth tracker must treat the entry paragraph as one
    // entry — not let the inner begin/end split or drop it.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading One</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements![0];
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0].children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(1);
  });
});

describe("stringifyTableOfContents dirty strategy", () => {
  it("marks a fresh TOC dirty so the app regenerates entries", () => {
    const xml = stringifyTableOfContents("Table of Contents", {
      headingStyleRange: "1-3",
      hyperlink: true,
    });
    expect(xml).toContain('w:dirty="1"');
  });

  it("emits a clean field when rendered entries are carried", () => {
    const xml = stringifyTableOfContents(
      "Table of Contents",
      { headingStyleRange: "1-3" },
      "<w:p/>",
    );
    expect(xml).not.toContain("w:dirty");
    expect(xml).toContain("<w:p/>");
  });
});
