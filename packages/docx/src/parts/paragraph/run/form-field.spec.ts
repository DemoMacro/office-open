import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../../body";
import type { DocxReadContext } from "../../../context";

// Form fields never touch the read context (no relationships/media), so an
// empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  return parseParagraph(doc.elements![0], readCtx) as { children?: unknown[] };
}

function findFormField(opts: {
  children?: unknown[];
}): { formField: Record<string, unknown> } | undefined {
  return opts.children?.find(
    (c) => c !== null && typeof c === "object" && "formField" in (c as Record<string, unknown>),
  ) as { formField: Record<string, unknown> } | undefined;
}

describe("form field parse", () => {
  it("parses an unchecked checkbox form field", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>☐</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    const found = findFormField(opts);
    expect(found).toBeDefined();
    expect(found!.formField.name).toBe("Check1");
    expect(found!.formField.checkBox).toMatchObject({
      checked: false,
      default: false,
      sizeAuto: true,
    });
    expect(found!.formField.enabled).toBe(true);
    expect(found!.formField.calcOnExit).toBe(false);
  });

  it("parses a checked checkbox form field (val-less on/off defaults to true)", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check2"/><w:enabled/><w:checkBox><w:sizeAuto/><w:default/><w:checked/></w:checkBox></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>☒</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    const found = findFormField(opts);
    expect(found).toBeDefined();
    expect(found!.formField.checkBox).toMatchObject({ checked: true, default: true });
  });

  it("collapses the fldChar sequence to a single formField child (skips instrText + result)", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check3"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>☐</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    // Only the formField child — no stray ☐ result text, no FORMCHECKBOX instrText.
    expect(opts.children).toHaveLength(1);
    expect("formField" in (opts.children![0] as Record<string, unknown>)).toBe(true);
  });

  it("does not treat a plain complex field (no ffData, e.g. PAGE) as a form field", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    // A plain complex field has no w:ffData, so it must NOT be collected as a
    // form field; the field markers/instrText are skipped and the result text
    // ("1") survives as a normal run instead of being swallowed.
    expect(findFormField(opts)).toBeUndefined();
  });
});
