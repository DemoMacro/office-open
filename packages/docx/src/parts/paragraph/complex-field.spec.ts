import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Complex fields never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseParagraph(el, readCtx) as { children?: unknown[] };
}

function findComplexField(opts: { children?: unknown[] }): Record<string, unknown> | undefined {
  return opts.children?.find(
    (c) => c !== null && typeof c === "object" && "complexField" in (c as Record<string, unknown>),
  ) as Record<string, unknown> | undefined;
}

describe("complex field parse", () => {
  it("parses a plain complex field (PAGE) with instruction and result", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const cf = findComplexField(opts);
    expect(cf).toBeDefined();
    expect(cf!.complexField).toMatchObject({ instruction: " PAGE ", result: "1" });
  });

  it("parses a complex field without a separate/result", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const cf = findComplexField(opts);
    expect(cf).toBeDefined();
    expect(cf!.complexField).toMatchObject({ instruction: " DATE " });
    expect((cf!.complexField as Record<string, unknown>).result).toBeUndefined();
  });

  it("concatenates instrText and result across multiple runs", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> HYPER</w:instrText></w:r>' +
        '<w:r><w:instrText xml:space="preserve">LINK </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>cli</w:t></w:r>" +
        "<w:r><w:t>ck</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const cf = findComplexField(opts);
    expect(cf!.complexField).toMatchObject({ instruction: " HYPERLINK ", result: "click" });
  });

  it("parses a deleted field (w:delInstrText) inside a deletion wrapper", () => {
    // Deleted fields spell the instruction w:delInstrText; the chain collapses
    // to a complexField child of the deletion wrapper (re-emitted with
    // w:delInstrText on stringify).
    const opts = parseParagraphXml(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:delInstrText xml:space="preserve">PAGE</w:delInstrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:del>",
    );
    const del = opts.children?.find(
      (c) => c !== null && typeof c === "object" && "deletion" in (c as Record<string, unknown>),
    ) as Record<string, unknown> | undefined;
    expect(del).toBeDefined();
    expect(del!.deletion).toMatchObject({
      id: 1,
      author: "Alice",
      children: [{ complexField: { instruction: "PAGE" } }],
    });
  });

  it("keeps non-plain instruction runs verbatim (per-run rPr + w:br)", () => {
    // Word splits a PAGE format switch across runs, styling the spacer runs
    // with CommentReference and embedding a line break — a shape the plain
    // instruction template cannot reproduce, so the runs round-trip verbatim.
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:instrText xml:space="preserve"> </w:instrText></w:r>' +
        '<w:r><w:instrText>PAGE \\# "&#x27;Page: &#x27;#&#x27;</w:instrText></w:r>' +
        "<w:r><w:br/><w:instrText>&#x27;&quot;</w:instrText></w:r>" +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:instrText xml:space="preserve"> </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const cf = findComplexField(opts);
    const field = cf!.complexField as Record<string, unknown>;
    // Semantic channel: the instruction text concatenates in order.
    const instruction = field.instruction as string;
    expect(instruction).toContain("PAGE \\#");
    expect(instruction.startsWith(" ") && instruction.endsWith(" ")).toBe(true);
    expect(instruction).toContain("'\"");
    // Fidelity channel: the exact run chain is carried verbatim.
    expect(field.instrRunsXml).toContain('<w:rStyle w:val="CommentReference"/>');
    expect(field.instrRunsXml).toContain("<w:br/>");
  });

  it("does not set the verbatim channel for plain instruction runs", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const cf = findComplexField(opts);
    expect((cf!.complexField as Record<string, unknown>).instrRunsXml).toBeUndefined();
  });
});
