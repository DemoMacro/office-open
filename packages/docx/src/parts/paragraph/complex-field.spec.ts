import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";
import type { ParagraphOptions } from "../paragraph";
import type { SimpleFieldOptions } from "./run/simple-field";

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

  it("does not inherit xml:space from the preceding field", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FIELD1 </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t xml:space="preserve">1 </w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FIELD2 </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>text</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const fields = opts.children?.filter(
      (child) => typeof child === "object" && child !== null && "complexField" in child,
    );
    expect(fields).toHaveLength(2);
    expect(fields?.[1]).toMatchObject({
      complexField: {
        instructionPreserveSpace: true,
        result: "text",
        resultPreserveSpace: false,
      },
    });
  });

  it("preserves an empty cached-result stage after a separate marker", () => {
    const source =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> DOCPROPERTY Unsupported </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const opts = parseParagraphXml(source);
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;

    expect(field).toMatchObject({ instruction: " DOCPROPERTY Unsupported ", hasResult: true });
    expect(field.result).toBeUndefined();
    expect(stringifyParagraph(opts, writeCtx)).toBe(`<w:p>${source}</w:p>`);
  });

  it("preserves an unterminated non-TOC field instead of dropping its control runs", () => {
    const opts = parseParagraphXml(
      '<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGE </w:instrText></w:r>",
    );
    expect(opts.children).toHaveLength(1);
    const raw = opts.children?.[0] as { rawXml?: string };
    expect(raw.rawXml).toContain('w:fldCharType="begin"');
    expect(raw.rawXml).toContain("<w:noProof/>");
    expect(raw.rawXml).toContain("<w:instrText> PAGE </w:instrText>");
    expect(stringifyParagraph(opts, writeCtx)).toContain('w:fldCharType="begin"');
  });

  it("preserves an orphan field separator outside a TOC aggregator", () => {
    const opts = parseParagraphXml(
      '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>',
    );
    const marker = opts.children?.[0] as { fieldMarker?: { type: string; rPrXml?: string } };
    expect(marker.fieldMarker).toEqual({
      type: "separate",
      rPrXml: "<w:rPr><w:b/></w:rPr>",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain('w:fldCharType="separate"');
  });

  it("preserves rsids on a plain simple-field cached run", () => {
    const xml =
      '<w:fldSimple w:instr=" PAGE "><w:r w:rsidR="00112233"><w:t>1</w:t></w:r></w:fldSimple>';
    const opts = parseParagraphXml(xml);
    const field = opts.children?.find(
      (child) => typeof child === "object" && "simpleField" in child,
    ) as { simpleField: SimpleFieldOptions } | undefined;
    expect(field?.simpleField.cachedRuns).toHaveLength(1);
    expect(stringifyParagraph(opts, writeCtx)).toContain('w:rsidR="00112233"');
  });

  it("round-trips field data on the begin marker", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin">' +
        '<w:fldData xml:space="preserve">AQID</w:fldData></w:fldChar></w:r>' +
        "<w:r><w:instrText> ADDIN Example </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(findComplexField(opts)!.complexField).toMatchObject({ fieldData: "AQID" });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:fldData xml:space="preserve">AQID</w:fldData>',
    );
  });

  it("round-trips numbering revision metadata on field markers", () => {
    const xml =
      '<w:r><w:fldChar w:fldCharType="begin">' +
      '<w:numberingChange w:original="1" w:id="7" w:author="Ada" w:date="2026-01-02T03:04:00Z"/>' +
      "</w:fldChar></w:r>" +
      "<w:r><w:instrText> NUMPAGES </w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const opts = parseParagraphXml(xml);
    expect(findComplexField(opts)!.complexField).toMatchObject({
      beginNumberingChange: {
        original: "1",
        id: "7",
        author: "Ada",
        date: "2026-01-02T03:04:00Z",
      },
    });
    expect(stringifyParagraph(opts, writeCtx)).toBe(`<w:p>${xml}</w:p>`);
  });

  it("round-trips numbering revision metadata on an orphan marker", () => {
    const xml =
      '<w:r><w:fldChar w:fldCharType="end">' +
      '<w:numberingChange w:original="2" w:id="8" w:author="Grace"/></w:fldChar></w:r>';
    const opts = parseParagraphXml(xml);
    expect(opts.children?.[0]).toMatchObject({
      fieldMarker: {
        type: "end",
        numberingChange: { original: "2", id: "8", author: "Grace" },
      },
    });
    expect(stringifyParagraph(opts, writeCtx)).toBe(`<w:p>${xml}</w:p>`);
  });

  it("round-trips numbering revision metadata from a closing field marker", () => {
    const xml =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      "<w:r><w:instrText> NUMPAGES </w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end">' +
      '<w:numberingChange w:original="1" w:id="7" w:author="Ada"/></w:fldChar></w:r>';
    const opts = parseParagraphXml(xml);
    expect(findComplexField(opts)!.complexField).toMatchObject({
      endNumberingChange: { original: "1", id: "7", author: "Ada" },
    });
  });

  it("emits a cached footnote reference from a complex field result", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> REF note </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:footnoteReference w:id="7"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(stringifyParagraph(opts, writeCtx)).toContain('<w:footnoteReference w:id="7"/>');
  });

  it("emits a cached footnote reference from a simple field result", () => {
    const opts = parseParagraphXml(
      '<w:fldSimple w:instr=" REF note "><w:r><w:rPr/><w:footnoteReference w:id="8"/></w:r></w:fldSimple>',
    );
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:fldSimple w:instr=" REF note "><w:r><w:rPr/><w:footnoteReference w:id="8"/></w:r></w:fldSimple>',
    );
  });

  it("preserves an empty control rPr only on its source stage", () => {
    const opts = parseParagraphXml(
      '<w:r><w:rPr/><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText>PAGE</w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(findComplexField(opts)!.complexField).toMatchObject({
      rPrXml: "<w:rPr/>",
      separatorRPrXml: "",
      endRPrXml: "",
    });
    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml.match(/<w:rPr\/>/g)).toHaveLength(1);
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

  it("keeps the deletion context inside an inline SDT", () => {
    const opts = parseParagraphXml(
      '<w:del w:id="1" w:author="Alice">' +
        '<w:sdt><w:sdtPr><w:tag w:val="deleted"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>1</w:t></w:r></w:sdtContent></w:sdt>" +
        "</w:del>",
    );
    expect(stringifyParagraph(opts, writeCtx)).toContain("<w:delText>1</w:delText>");
    expect(stringifyParagraph(opts, writeCtx)).not.toContain("<w:t>1</w:t>");
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

  it("preserves nested complex fields in a simple field", () => {
    const opts = parseParagraphXml(
      '<w:fldSimple w:instr=" REF note ">' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:fldSimple>",
    );
    const field = opts.children?.find(
      (child) => typeof child === "object" && "simpleField" in child,
    ) as { simpleField: SimpleFieldOptions } | undefined;
    if (!field) throw new Error("simple field not parsed");
    expect(field.simpleField.cachedRuns).toHaveLength(1);
    expect(field.simpleField.cachedRuns).toMatchObject([
      { complexField: { instruction: " PAGE " } },
    ]);
    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1);
  });

  it("round-trips mixed deleted field stages with deleted text tags", () => {
    const xml =
      '<w:del w:id="1" w:author="Alice">' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:delText xml:space="preserve">1</w:delText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:del>";
    const opts = parseParagraphXml(xml);
    const output = stringifyParagraph(opts, writeCtx);
    expect(output).toContain('<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>');
    expect(output).toContain('<w:delText xml:space="preserve">1</w:delText>');
  });

  it("preserves deleted text in a nested result field", () => {
    const xml =
      '<w:del w:id="1" w:author="Alice">' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      "<w:r><w:instrText> OUTER </w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      "<w:r><w:instrText> INNER </w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:delText xml:space="preserve">1</w:delText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:del>";
    const output = stringifyParagraph(parseParagraphXml(xml), writeCtx);
    expect(output).toContain(
      '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText xml:space="preserve">1</w:delText>',
    );
    expect(output).not.toContain('<w:t xml:space="preserve">1</w:t>');
  });

  it("preserves proof error markers on each complex field stage", () => {
    const xml =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:proofErr w:type="spellStart"/>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      "<w:r><w:t>1</w:t></w:r>" +
      '<w:proofErr w:type="spellEnd"/>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const opts = parseParagraphXml(xml);
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    expect(field.instructionMembers).toEqual([
      { proofErr: "spellStart" },
      { run: { children: [{ instructionText: " PAGE ", preserveSpace: true }] } },
    ]);
    expect(field.resultMembers).toEqual([{ run: { text: "1" } }, { proofErr: "spellEnd" }]);
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:fldChar w:fldCharType="begin"/></w:r><w:proofErr w:type="spellStart"/>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>' +
        '<w:proofErr w:type="spellEnd"/><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
  });

  it("keeps proof markers inside tracked complex-field results", () => {
    const xml =
      '<w:ins w:id="1" w:author="Alice">' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:proofErr w:type="spellStart"/>' +
      "<w:r><w:t>1</w:t></w:r>" +
      '<w:proofErr w:type="spellEnd"/>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:ins>";
    const opts = parseParagraphXml(xml);
    const insertion = opts.children?.find(
      (child) => child !== null && typeof child === "object" && "insertion" in child,
    ) as { insertion: { children: { complexField?: Record<string, unknown> }[] } };
    const field = insertion.insertion.children[0]!.complexField!;
    expect(field.resultMembers).toEqual([
      { proofErr: "spellStart" },
      { run: { text: "1" } },
      { proofErr: "spellEnd" },
    ]);
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="spellStart"/>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:proofErr w:type="spellEnd"/><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
  });

  it("feeds every stage packed into one compound run", () => {
    const opts = parseParagraphXml(
      "<w:r>" +
        "<w:rPr><w:noProof/></w:rPr>" +
        '<w:fldChar w:fldCharType="begin"/>' +
        '<w:instrText xml:space="preserve"> IF </w:instrText>' +
        '<w:fldChar w:fldCharType="begin"/>' +
        "<w:instrText>inner</w:instrText>" +
        '<w:fldChar w:fldCharType="separate"/>' +
        "<w:instrText>1</w:instrText>" +
        '<w:fldChar w:fldCharType="end"/>' +
        "<w:instrText> = 1 </w:instrText>" +
        '<w:fldChar w:fldCharType="separate"/>' +
        "<w:t>one</w:t>" +
        '<w:fldChar w:fldCharType="end"/>' +
        "</w:r>",
    );
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    const nestedField = (field.instructionMembers as unknown[]).find(
      (member) =>
        member !== null &&
        typeof member === "object" &&
        "complexField" in (member as Record<string, unknown>),
    ) as { complexField: Record<string, unknown> } | undefined;
    expect(nestedField?.complexField.instruction).toContain("inner");
    expect(field.instruction).toContain(" = 1 ");
    expect(field.result).toBe("one");
    expect(field.compoundRunXml).toContain(">inner<");
    expect(field.instructionMembers).toBeDefined();
    expect(field.compoundRunXml).toContain("> = 1 <");
    expect(field.result).toBe("one");

    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(2);
    expect(xml).toContain(">one<");
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

  it("keeps hyperlinks inside the complex-field result", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:hyperlink w:anchor="section"><w:r><w:t>Section</w:t></w:r></w:hyperlink>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const field = findComplexField(opts)!.complexField as Record<string, unknown>;
    expect(field.resultMembers).toEqual([
      {
        hyperlink: {
          anchor: "section",
          children: [{ text: "Section" }],
        },
      },
    ]);

    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml.indexOf('w:fldCharType="separate"')).toBeLessThan(xml.indexOf("<w:hyperlink"));
    expect(xml.indexOf("<w:hyperlink")).toBeLessThan(xml.indexOf('w:fldCharType="end"'));
  });
});
