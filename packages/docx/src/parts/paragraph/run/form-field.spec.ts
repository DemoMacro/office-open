import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../../body";
import type { DocxReadContext } from "../../../context";

// Form fields never touch the read context (no relationships/media), so an
// empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseParagraph(el, readCtx) as { children?: unknown[] };
}

function findFormField(opts: {
  children?: unknown[];
}): { formField: Record<string, unknown>; lastRenderedPageBreak?: boolean } | undefined {
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

  it("round-trips entry and exit macros with help and status text", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData>' +
        '<w:name w:val="Text1"/><w:enabled/><w:entryMacro w:val="OnEntry"/>' +
        '<w:exitMacro w:val="OnExit"/><w:helpText w:type="text" w:val="Press F1"/>' +
        '<w:statusText w:type="autoText" w:val="StatusEntry"/>' +
        "<w:textInput/></w:ffData></w:fldChar></w:r>" +
        "<w:r><w:instrText> FORMTEXT </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const formField = findFormField(opts)!.formField;
    expect(formField).toMatchObject({
      entryMacro: "OnEntry",
      exitMacro: "OnExit",
      helpText: { type: "text", value: "Press F1" },
      statusText: { type: "autoText", value: "StatusEntry" },
    });
    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain('<w:entryMacro w:val="OnEntry"/>');
    expect(xml).toContain('<w:exitMacro w:val="OnExit"/>');
    expect(xml).toContain('<w:helpText w:type="text" w:val="Press F1"/>');
    expect(xml).toContain('<w:statusText w:type="autoText" w:val="StatusEntry"/>');
  });

  it("keeps a pagination hint parked on the form-field begin run", () => {
    const opts = parseParagraphXml(
      '<w:r><w:lastRenderedPageBreak/><w:fldChar w:fldCharType="begin">' +
        '<w:ffData><w:name w:val="List1"/><w:ddList><w:listEntry w:val="one"/></w:ddList></w:ffData>' +
        "</w:fldChar></w:r><w:r><w:instrText> FORMDROPDOWN </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const field = findFormField(opts)!;
    expect(field.lastRenderedPageBreak).toBe(true);
    expect(stringifyParagraph(opts as never, writeCtx)).toContain(
      '<w:lastRenderedPageBreak/><w:fldChar w:fldCharType="begin"',
    );
  });

  it("preserves absent help and status attributes without defaults", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/>' +
        "<w:helpText/><w:statusText/><w:textInput/></w:ffData></w:fldChar></w:r>" +
        "<w:r><w:instrText> FORMTEXT </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(findFormField(opts)!.formField.helpText).toEqual({});
    expect(findFormField(opts)!.formField.statusText).toEqual({});
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

  it("does not synthesize an absent checkbox default from checked", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check"/>' +
        "<w:checkBox><w:sizeAuto/><w:checked/></w:checkBox></w:ffData></w:fldChar></w:r>" +
        '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(findFormField(opts)!.formField.checkBox).toEqual({ sizeAuto: true, checked: true });
  });

  it("preserves absent enabled and calcOnExit controls", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check"/>' +
        "<w:checkBox><w:sizeAuto/></w:checkBox></w:ffData></w:fldChar></w:r>" +
        '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const field = findFormField(opts)!.formField;
    expect(field.enabled).toBeUndefined();
    expect(field.calcOnExit).toBeUndefined();
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
    // A plain complex field has no w:ffData, so it is NOT a form field — it is
    // collected as a complexField instead (see complex-field.spec.ts).
    expect(findFormField(opts)).toBeUndefined();
  });

  it("parses a text input form field with its current (filled-in) value", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/><w:enabled/><w:textInput><w:type w:val="regular"/><w:default w:val="Placeholder"/></w:textInput></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t xml:space="preserve">Hello World</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    const found = findFormField(opts);
    expect(found).toBeDefined();
    expect(found!.formField.textInput).toMatchObject({
      type: "regular",
      default: "Placeholder",
      value: "Hello World",
    });
  });

  it("concatenates result text split across multiple runs into value", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text2"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>foo</w:t></w:r>" +
        "<w:r><w:t>bar</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    const found = findFormField(opts);
    const ti = found!.formField.textInput as { value?: string };
    expect(ti.value).toBe("foobar");
  });

  it("preserves bookmark markers interleaved with form-field result runs", async () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput/></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>foo</w:t></w:r>" +
        '<w:bookmarkStart w:id="7" w:name="middle"/>' +
        '<w:bookmarkEnd w:id="7"/>' +
        "<w:r><w:t>bar</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const found = findFormField(opts);
    expect((found as { resultMembers?: unknown[] })?.resultMembers).toEqual([
      { run: { text: "foo" } },
      { bookmarkStart: { id: 7, name: "middle" } },
      { bookmarkEnd: { id: 7 } },
      { run: { text: "bar" } },
    ]);

    const { stringifyChildDispatch } = await import("../../inline");
    const xml = stringifyChildDispatch(found as never, {} as never) as string;
    expect(xml.indexOf(">foo<")).toBeLessThan(xml.indexOf("<w:bookmarkStart"));
    expect(xml.indexOf("<w:bookmarkEnd")).toBeLessThan(xml.indexOf(">bar<"));
  });

  it("leaves value unset when the textInput result run is empty", () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text3"/><w:textInput><w:default w:val="Placeholder"/></w:textInput></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );

    const found = findFormField(opts);
    const ti = found!.formField.textInput as { value?: string; default?: string };
    expect(ti.value).toBeUndefined();
    expect(ti.default).toBe("Placeholder");
  });

  it("preserves a result-less field and instruction run styling", async () => {
    const opts = parseParagraphXml(
      '<w:r w:rsidRPr="007B7AC4"><w:rPr><w:b/></w:rPr>' +
        '<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/>' +
        '<w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/>' +
        '<w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
        '<w:r w:rsidRPr="007B7AC4"><w:rPr><w:b/></w:rPr>' +
        '<w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
        '<w:r w:rsidRPr="007B7AC4"><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const found = findFormField(opts);
    expect(found).toMatchObject({
      instructionRunPropertiesRsid: "007B7AC4",
      hasResult: false,
      rPrXml: "<w:rPr><w:b/></w:rPr>",
    });

    const { stringifyChildDispatch } = await import("../../inline");
    const xml = stringifyChildDispatch(found as never, {} as never) as string;
    expect(xml).toContain('w:rsidRPr="007B7AC4"');
    expect(xml).toContain("<w:rPr><w:b/></w:rPr><w:instrText");
    expect(xml).not.toContain('w:fldCharType="separate"');
    expect(xml).not.toContain("☐");
  });

  it("preserves a locked begin field character", async () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin" w:fldLock="1">' +
        '<w:ffData><w:name w:val="Text1"/><w:textInput/></w:ffData>' +
        "</w:fldChar></w:r>" +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>value</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const found = findFormField(opts);
    expect(found).toMatchObject({ fieldLock: true });

    const { stringifyChildDispatch } = await import("../../inline");
    const xml = stringifyChildDispatch(found as never, {} as never) as string;
    expect(xml).toContain('<w:fldChar w:fldLock="1" w:fldCharType="begin">');
  });

  it("preserves a result-less dropdown field separator", async () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val=""/>' +
        '<w:enabled/><w:calcOnExit w:val="0"/><w:ddList>' +
        '<w:listEntry w:val="Location 1"/><w:listEntry w:val="Location 2"/>' +
        "</w:ddList></w:ffData></w:fldChar></w:r>" +
        '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
        "<w:r/>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const found = findFormField(opts);
    expect(found).toMatchObject({ hasResult: true });
    expect((found as { resultRunsXml?: string }).resultRunsXml).toBe("");

    const { stringifyChildDispatch } = await import("../../inline");
    const xml = stringifyChildDispatch(found as never, {} as never) as string;
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(1);
  });

  it("preserves a form field nested in another form field result", async () => {
    const opts = parseParagraphXml(
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Outer"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Inner"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    const found = findFormField(opts);
    expect(found).toMatchObject({ hasResult: true });
    const nestedField = (found as { resultMembers?: unknown[] }).resultMembers?.find(
      (member) =>
        member !== null &&
        typeof member === "object" &&
        "formField" in (member as Record<string, unknown>),
    ) as { formField: { name?: string } } | undefined;
    expect(nestedField?.formField.name).toBe("Inner");

    const { stringifyChildDispatch } = await import("../../inline");
    const xml = stringifyChildDispatch(found as never, {} as never) as string;
    expect([...xml.matchAll(/w:fldCharType="begin"/g)]).toHaveLength(2);
    expect([...xml.matchAll(/w:fldCharType="end"/g)]).toHaveLength(2);
  });
});
