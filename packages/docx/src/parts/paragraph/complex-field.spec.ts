import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";
import type { ParagraphOptions } from "../paragraph";

// Complex fields never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): ParagraphOptions {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseParagraph(el, readCtx);
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

  it("preserves absent and present xml:space on plain instruction text", () => {
    const withoutMarker = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText>PAGE</w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const plainField = findComplexField(withoutMarker)!.complexField as Record<string, unknown>;
    expect(plainField.instructionPreserveSpace).toBe(false);
    expect(stringifyParagraph(withoutMarker, writeCtx)).toContain(
      "<w:instrText>PAGE</w:instrText>",
    );

    const withMarker = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(
      (findComplexField(withMarker)!.complexField as Record<string, unknown>)
        .instructionPreserveSpace,
    ).toBe(true);
    expect(stringifyParagraph(withMarker, writeCtx)).toContain(
      '<w:instrText xml:space="preserve"> PAGE </w:instrText>',
    );
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

  it("preserves simple fields interleaved in a complex field code", () => {
    const inner =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText>tc "</w:instrText></w:r>' +
      '<w:fldSimple w:instr="seq Appendix"><w:r><w:rPr><w:noProof/></w:rPr><w:instrText>A</w:instrText></w:r></w:fldSimple>' +
      "<w:r><w:instrText>-</w:instrText></w:r>" +
      '<w:fldSimple w:instr="seq Figure"><w:r><w:instrText>1</w:instrText></w:r></w:fldSimple>' +
      "<w:r><w:instrText>&quot;</w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const opts = parseParagraphXml(inner);
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    expect(field.instruction).toBe('tc "-"');
    expect(field.instructionMembers).toEqual([
      { run: { children: [{ instructionText: 'tc "' }] } },
      {
        simpleField: {
          cachedInstructionTextPreserveSpace: false,
          cachedValue: "",
          cachedValuePreserveSpace: false,
          instruction: "seq Appendix",
          cachedInstructionText: "A",
          cachedInstructionRPrXml: "<w:rPr><w:noProof/></w:rPr>",
        },
      },
      { run: { children: [{ instructionText: "-" }] } },
      {
        simpleField: {
          cachedInstructionTextPreserveSpace: false,
          cachedValue: "",
          cachedValuePreserveSpace: false,
          instruction: "seq Figure",
          cachedInstructionText: "1",
        },
      },
      { run: { children: [{ instructionText: '"' }] } },
    ]);
    expect(stringifyParagraph(opts, writeCtx)).toContain("<w:instrText>tc &quot;</w:instrText>");
  });

  it("preserves typed instruction runs around bookmarks", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText>A</w:instrText></w:r>" +
        '<w:bookmarkStart w:id="11" w:name="middle"/>' +
        "<w:r><w:instrText>B</w:instrText></w:r>" +
        '<w:bookmarkEnd w:id="11"/>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    expect(field.instruction).toBe("AB");
    expect(field.instructionMembers).toEqual([
      { run: { children: [{ instructionText: "A" }] } },
      { bookmarkStart: { id: 11, name: "middle" } },
      { run: { children: [{ instructionText: "B" }] } },
      { bookmarkEnd: { id: 11 } },
    ]);

    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml.indexOf('w:fldCharType="begin"')).toBeLessThan(
      xml.indexOf("<w:instrText>A</w:instrText>"),
    );
    expect(xml.indexOf("<w:instrText>A</w:instrText>")).toBeLessThan(
      xml.indexOf("<w:bookmarkStart"),
    );
    expect(xml.indexOf("<w:bookmarkStart")).toBeLessThan(
      xml.indexOf("<w:instrText>B</w:instrText>"),
    );
    expect(xml.indexOf("<w:instrText>B</w:instrText>")).toBeLessThan(xml.indexOf("<w:bookmarkEnd"));
    expect(xml.indexOf("<w:bookmarkEnd")).toBeLessThan(xml.indexOf('w:fldCharType="end"'));
  });

  it("preserves typed result runs around bookmarks", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>A</w:t></w:r>" +
        '<w:bookmarkStart w:id="21" w:name="result"/>' +
        "<w:r><w:t>B</w:t></w:r>" +
        '<w:bookmarkEnd w:id="21"/>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    expect(field.result).toBe("AB");
    expect(field.resultMembers).toEqual([
      { run: { text: "A" } },
      { bookmarkStart: { id: 21, name: "result" } },
      { run: { text: "B" } },
      { bookmarkEnd: { id: 21 } },
    ]);

    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(xml.indexOf("<w:t>A</w:t>")).toBeLessThan(xml.indexOf("<w:bookmarkStart"));
    expect(xml.indexOf("<w:bookmarkStart")).toBeLessThan(xml.indexOf("<w:t>B</w:t>"));
    expect(xml.indexOf("<w:t>B</w:t>")).toBeLessThan(xml.indexOf("<w:bookmarkEnd"));
    expect(xml.indexOf("<w:bookmarkEnd")).toBeLessThan(xml.indexOf('w:fldCharType="end"'));
  });
});
