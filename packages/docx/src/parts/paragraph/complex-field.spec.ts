import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Complex fields never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  return parseParagraph(doc.elements![0], readCtx) as { children?: unknown[] };
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
});
