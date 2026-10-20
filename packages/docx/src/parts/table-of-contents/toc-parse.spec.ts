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
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const opts = parseTocFieldFromElements(body.elements ?? []);
    expect(opts).toEqual({
      headingStyleRange: "1-3",
      hyperlink: true,
      hideTabAndPageNumbersInWebView: true,
    });
  });
});

describe("parseBody revision and shared-run TOC fields", () => {
  it("does not treat a generic SDT wrapper as a second TOC", () => {
    const inner =
      '<w:sdt><w:sdtPr><w:id w:val="5983877"/></w:sdtPr>' +
      '<w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>' +
      "<w:p><w:r><w:t>Chapter</w:t></w:r></w:p>" +
      '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>';
    const outer =
      `<w:sdt><w:sdtPr><w:id w:val="12652376"/></w:sdtPr>` +
      `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
    const root = parseXml(outer).elements?.[0];
    if (!root) throw new Error("parsed SDT has no root element");
    expect(
      parseToc(root, readCtx, (elements) =>
        elements.map((element) => ({ paragraph: parseTocFieldFromElements([element]) as never })),
      ),
    ).toBeUndefined();
  });

  it("does not aggregate a TOC field hidden in a moved-from block", () => {
    const xml = `<w:body>
      <w:p><w:moveFrom><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:moveFrom></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const children = parseBody(body, readCtx)[0]?.children ?? [];
    expect(children[0]).toHaveProperty("rawXml");
    expect(children[0]).not.toHaveProperty("toc");
  });

  it("does not duplicate result text shared with the first field-control run", () => {
    const xml = `<w:body>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/>
        <w:instrText xml:space="preserve">TOC \\o "1-3"</w:instrText>
        <w:fldChar w:fldCharType="separate"/><w:t>First</w:t></w:r></w:p>
      <w:p><w:r><w:t>2</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const child = parseBody(body, readCtx)[0]?.children[0];
    if (!child || !("toc" in child)) throw new Error("parsed document has no TOC child");
    const toc = child.toc;
    expect(toc?.headResultInFirstRun).toBe(true);
    expect(JSON.stringify(toc?.entries?.[0])).not.toContain("First");
    expect(toc?.headRunsXml).toContain("First");
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
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx);
    expect(result?.alias).toBe("Contents");
    expect(result?.headingStyleRange).toBe("1-3");
    expect(result?.hyperlink).toBe(true);
  });

  it("keeps gallery-only content as a generic SDT", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:t>Custom gallery content</w:t></w:r></w:p></w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    expect(parseToc(el, readCtx)).toBeUndefined();
  });

  it("recognizes a deleted TOC instruction", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:delInstrText> TOC \\o "1-3" </w:delInstrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p></w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    expect(parseToc(el, readCtx)?.headingStyleRange).toBe("1-3");
  });

  it("preserves the lexical value of w:docPartUnique", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique w:val="0"/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx);

    expect(result?.docPartUnique).toBe(true);
    expect(result?.docPartUniqueRaw).toBe("0");
    expect(stringifyTableOfContents(undefined, result)).toContain('<w:docPartUnique w:val="0"/>');
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
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx, (els, ctx) => els.map((e) => parseSectionChild(e, ctx)));
    // 2 rendered entries + the bare field-closing paragraph (the stringify
    // path injects the end run back into it, keeping the paragraph count).
    expect(result?.entries).toHaveLength(3);
  });

  it("does not duplicate deferred outer TOC controls inside an entry", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent>
        <w:p><w:ins w:id="21"><w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:ins w:id="22"><w:r><w:t>Heading One</w:t></w:r></w:ins></w:ins></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
      </w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx, (els, ctx) =>
      els.map((element) => parseSectionChild(element, ctx)),
    );
    expect(result).toBeDefined();
    expect(JSON.stringify(result?.entries ?? [])).not.toContain('fldCharType="begin"');
    expect(result?.headRunsXml).toBeDefined();
    expect(result?.headRunsXml).toMatch(/^<w:ins\b/);
    expect(result?.headRunsXml).toMatch(/<\/w:ins>$/);
  });

  it("does not duplicate deferred deleted TOC controls inside an entry", () => {
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent>
        <w:p><w:del w:id="21"><w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:delInstrText xml:space="preserve"> TOC \\o "1-3" </w:delInstrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:ins w:id="22"><w:r><w:t>Heading One</w:t></w:r></w:ins></w:del></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
      </w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx, (els, ctx) =>
      els.map((element) => parseSectionChild(element, ctx)),
    );
    expect(JSON.stringify(result?.entries ?? [])).not.toContain('fldCharType="begin"');
  });

  it("does not inject a second field end when the last entry already closes the field", () => {
    const fieldEnd = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r>${fieldEnd}</w:p>
      </w:sdtContent>
    </w:sdt>`;
    const element = parseXml(xml).elements?.[0];
    if (!element) throw new Error("parsed document has no root element");
    const result = parseToc(element, readCtx, (elements, context) =>
      elements.map((child) => parseSectionChild(child, context)),
    );
    expect(result?.endMarkerInEntries).toBe(true);
    expect(
      stringifyTableOfContents(
        undefined,
        result,
        `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r>${fieldEnd}</w:p>`,
      ).match(/w:fldCharType="end"/g),
    ).toHaveLength(1);
  });
});

describe("parseBody self-contained TOC entries", () => {
  it("does not duplicate the aggregate field inside its first entry", () => {
    const xml = `<w:body>
      <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\f \\o "1-9" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:hyperlink w:anchor="heading"><w:r><w:t>Heading One</w:t></w:r></w:hyperlink>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const toc = parseBody(body, readCtx)
      .flatMap((section) => section.children)
      .find((child) => "toc" in child)?.toc;
    const firstEntry = toc?.entries?.[0];
    const entryChildren =
      firstEntry && "paragraph" in firstEntry && typeof firstEntry.paragraph !== "string"
        ? (firstEntry.paragraph.children ?? [])
        : [];
    expect(entryChildren).toHaveLength(2);
    expect(entryChildren[0]).toHaveProperty("hyperlink");
    expect(entryChildren[1]).toMatchObject({ fieldMarker: { type: "end" } });
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
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const json = JSON.stringify(sections[0]?.children ?? []);
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
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(3);
  });

  it("omits entries when the TOC field has no rendered result", () => {
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
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
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(3);
  });

  it("captures the first entry when begin+separate share its paragraph", () => {
    // Word often emits the TOC begin + separate + first rendered entry all in
    // one paragraph (no separate field-head paragraph). The first entry must
    // still be captured.
    const xml = `<w:body>
      <w:p><w:r><w:t>Contents</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r><w:t>Heading Two</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(3);
  });

  it("flags and injects head runs after entry text carried before the chain", () => {
    // Some TOC templates carry the entry text ("Contents" heading) in the same
    // paragraph BEFORE the begin/instr/separate chain — the head runs must
    // re-inject after that text run to keep the source child order.
    const xml = `<w:body>
      <w:p><w:pPr><w:pStyle w:val="TOCHeading1"/></w:pPr><w:r><w:t>Contents</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: Record<string, unknown> }
      | undefined;
    expect(tocChild?.toc.headRunsAfterText).toBe(true);
    const entriesXml =
      `<w:p><w:pPr><w:pStyle w:val="TOCHeading1"/></w:pPr>` +
      `<w:r><w:t>Contents</w:t></w:r></w:p>`;
    const out = stringifyTableOfContents(undefined, tocChild!.toc as never, entriesXml);
    expect(out.indexOf("<w:t>Contents</w:t>")).toBeLessThan(out.indexOf('w:fldCharType="begin"'));
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
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
  });

  it("preserves a complete nested PAGEREF field in the TOC head paragraph", () => {
    // Pandoc and Word can place the first rendered entry in the same paragraph
    // as the outer TOC head. The incomplete outer field must release typed
    // result members without losing the complete nested page-number field.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\u </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading One</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
    const entry = JSON.stringify(tocChild?.toc.entries?.[0]);
    expect(entry).toContain("Heading One");
    expect(entry).toContain("PAGEREF _Toc1");
    expect(entry).toContain('"result":"1"');
  });

  it("keeps the field-closing paragraph's properties as a trailing entry", () => {
    // Word parks the entry style's pPr (pStyle/divId) on the field-closing
    // paragraph. Dropping it loses that formatting on round-trip.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="10"/><w:divId w:val="1748839239"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
    const closing = JSON.stringify(tocChild?.toc.entries?.[1]);
    expect(closing).toContain('"divId":1748839239');
    expect(closing).toContain('"style":"10"');
  });

  it("returns a text-carrying closing paragraph to the body, not the TOC", () => {
    // Word drops the field end into a following heading when it updates the
    // TOC. That heading is body content — it must not move inside the TOC.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t>Heading Two</w:t></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const children = sections[0]?.children ?? [];
    const tocIdx = children.findIndex((c) => "toc" in c);
    const tocChild = children[tocIdx] as { toc: { entries?: unknown[] } };
    expect(tocChild.toc.entries).toHaveLength(1);
    // The heading follows the TOC as a regular body paragraph. Its field end
    // stays as a verbatim child so the cross-paragraph field remains closed.
    const after = children[tocIdx + 1];
    const afterJson = JSON.stringify(after);
    expect(afterJson).toContain("Heading2");
    expect(afterJson).toContain("Heading Two");
    expect(afterJson).toContain("fieldMarker");
  });

  it("carries the field control runs' rPr through the structured TOC", () => {
    // Word parks explicit style overrides (b w:val="0" negating an inherited
    // bold) on the invisible begin/instr/separate runs, and styles the end run
    // differently. The re-emitted field chain must carry both.
    const ctrl = '<w:rPr><w:noProof/><w:b w:val="0"/></w:rPr>';
    const end = "<w:rPr><w:b/></w:rPr>";
    const xml = `<w:body>
      <w:p><w:r>${ctrl}<w:fldChar w:fldCharType="begin"/></w:r>
        <w:r>${ctrl}<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r>${ctrl}<w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r>${end}<w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { rPrXml?: string; endRPrXml?: string } }
      | undefined;
    expect(tocChild?.toc.rPrXml).toBe(ctrl);
    expect(tocChild?.toc.endRPrXml).toBe(end);

    // The emitted chain threads the rPr through begin/instr/separate and the
    // distinct one through end; exactly one end marker for one begin.
    const emitted = stringifyTableOfContents("Contents", tocChild!.toc);
    expect(emitted.match(/<w:noProof\/>/g)).toHaveLength(3);
    expect(emitted).toContain(`${end}<w:fldChar w:fldCharType="end"/>`);
    expect(emitted.match(/w:fldCharType="end"/g)).toHaveLength(1);
  });

  it("preserves field-end run revision ids", () => {
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r w:rsidR="00000001" w:rsidRPr="00000002" w:rsidDel="00000003"><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | {
          toc: {
            endAdditionRsid?: string;
            endRunPropertiesRsid?: string;
            endDeletionRsid?: string;
          };
        }
      | undefined;
    expect(tocChild?.toc.endAdditionRsid).toBe("00000001");
    expect(tocChild?.toc.endRunPropertiesRsid).toBe("00000002");
    expect(tocChild?.toc.endDeletionRsid).toBe("00000003");
  });

  it("does not double-emit a page break kept inside the closing paragraph", () => {
    // When the closing paragraph joins the entries (it carries pPr), its page
    // break round-trips inside it — the standalone rescue must not fire too.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:br w:type="page"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const children = sections[0]?.children ?? [];
    // The page break lives inside the TOC's closing entry — no rescue copy.
    const pageBreaks = JSON.stringify(children).match(/"pageBreak"/g)?.length ?? 0;
    expect(pageBreaks).toBe(1);
    const standaloneBreaks = children.filter(
      (c) =>
        "paragraph" in c &&
        Array.isArray((c as { paragraph?: { children?: unknown[] } }).paragraph?.children) &&
        (c as { paragraph: { children: unknown[] } }).paragraph.children.length === 1 &&
        String(
          JSON.stringify((c as { paragraph: { children: unknown[] } }).paragraph.children[0]),
        ).includes("pageBreak"),
    );
    expect(standaloneBreaks).toHaveLength(0);
  });
});

describe("TOC field boundary round-trip", () => {
  it("captures a single-paragraph field's entry and slices around it", () => {
    // A one-entry TOC renders the entire field — begin, separate, entry text,
    // nested PAGEREF, end — inside one w:p. The splitter must still start the
    // field slice there (depth opens AND closes in the same element) and keep
    // the paragraph as the rendered entry.
    const xml = `<w:sdt>
      <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>Contents</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
          <w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>Heading One</w:t></w:r>
          <w:r><w:tab/></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
        <w:p><w:r><w:t>After</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = parseToc(el, readCtx, (els, ctx) => els.map((e) => parseSectionChild(e, ctx)));
    expect(result?.entries).toHaveLength(1);
    expect(JSON.stringify(result?.entries?.[0])).toContain("Heading One");
    // The heading paragraph and the paragraph after the field stay outside it.
    expect(result?.leading).toHaveLength(1);
    expect(JSON.stringify(result?.leading?.[0])).toContain("Contents");
    expect(result?.trailing).toHaveLength(1);
    expect(JSON.stringify(result?.trailing?.[0])).toContain("After");
  });

  it("keeps the last entry when the field end is parked in its paragraph", () => {
    // Word parks the outer field end in the last rendered entry's paragraph
    // (after the nested PAGEREF end). That paragraph carries an entry's markup
    // and must stay in the TOC; its consumed end is re-injected on emit, so the
    // body must not receive a field-less copy and endInBody must stay off.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Heading Two</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const children = sections[0]?.children ?? [];
    const tocChild = children.find((c) => "toc" in c) as
      | { toc: { entries?: unknown[]; endInBody?: boolean } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
    expect(tocChild?.toc.endInBody).not.toBe(true);
    const closing = JSON.stringify(tocChild?.toc.entries?.[1]);
    expect(closing).toContain("Heading Two");
    expect(closing).toContain("TOC2");
    // The closing paragraph lives inside the TOC — it is not duplicated into
    // the body as a standalone paragraph.
    expect(children).toHaveLength(1);
  });

  it("keeps bookmark order when the consumed field end precedes it", () => {
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:bookmarkStart w:id="7" w:name="_Ref463345912"/></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = sections[0]?.children.find((child) => "toc" in child) as
      | { toc: { fieldEndBeforeChildren?: boolean } }
      | undefined;
    expect(tocChild?.toc.fieldEndBeforeChildren).toBe(true);

    const entries = `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:bookmarkStart w:id="7" w:name="_Ref463345912"/></w:p>`;
    const output = stringifyTableOfContents(
      undefined,
      { bare: true, fieldEndBeforeChildren: true },
      entries,
    );
    const paragraph = output.slice(output.indexOf('<w:pStyle w:val="TOC1"/>'));
    const endIndex = paragraph.indexOf('<w:fldChar w:fldCharType="end"/>');
    const bookmarkIndex = paragraph.indexOf('<w:bookmarkStart w:id="7"');
    expect(endIndex).toBeGreaterThan(-1);
    expect(bookmarkIndex).toBeGreaterThan(endIndex);
  });

  it("injects the field end into a self-closing standalone paragraph", () => {
    const output = stringifyTableOfContents(
      undefined,
      { bare: true },
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>` +
        `<w:p w:rsidR="00ED3A69"/>`,
    );
    expect(output).toContain(
      `<w:p w:rsidR="00ED3A69"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
  });

  it("keeps a hyperlink-wrapped closing paragraph as the last entry", () => {
    // The closing paragraph may carry a custom entry style instead of TOCn;
    // the hyperlink wrapping the entry text is what identifies it as an entry.
    const xml = `<w:body>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Contents2"/></w:pPr>
        <w:hyperlink w:anchor="_Toc2"><w:r><w:t>Heading Two</w:t></w:r></w:hyperlink>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const children = sections[0]?.children ?? [];
    const tocChild = children.find((c) => "toc" in c) as
      | { toc: { entries?: unknown[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(2);
    expect(children).toHaveLength(1);
  });
});

describe("parseBody TOC property-only head paragraphs", () => {
  it("keeps a property-only head paragraph inside rendered TOC entries", () => {
    const xml = `<w:body>
      <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Heading One</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:body>`;
    const body = parseXml(xml).elements?.[0];
    if (!body) throw new Error("parsed document has no root element");
    const sections = parseBody(body, readCtx);
    const tocChild = (sections[0]?.children ?? []).find((c) => "toc" in c) as
      | { toc: { entries?: { paragraph?: { style?: string } }[] } }
      | undefined;
    expect(tocChild?.toc.entries).toHaveLength(3);
    expect(tocChild?.toc.entries?.[0]?.paragraph?.style).toBe("TOC1");
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
