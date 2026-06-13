import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Inline metadata carriers never touch the read context, so an empty mock
// suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  return parseParagraph(doc.elements![0], readCtx) as { children?: unknown[] };
}

function findChildByKey(
  opts: { children?: unknown[] },
  key: string,
): Record<string, unknown> | undefined {
  return opts.children?.find(
    (c) => c !== null && typeof c === "object" && key in (c as Record<string, unknown>),
  ) as Record<string, unknown> | undefined;
}

describe("inline metadata parse", () => {
  it("parses a simple field (fldSimple) with cached value", () => {
    const opts = parseParagraphXml(
      `<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>`,
    );

    const sf = findChildByKey(opts, "simpleField");
    expect(sf).toBeDefined();
    expect(sf!.simpleField).toMatchObject({ instruction: " PAGE ", cachedValue: "1" });
  });

  it("parses a simple field without a cached value", () => {
    const opts = parseParagraphXml(`<w:fldSimple w:instr=" AUTHOR \\* MERGEFORMAT "/>`);

    const sf = findChildByKey(opts, "simpleField");
    expect(sf).toBeDefined();
    expect(sf!.simpleField).toMatchObject({ instruction: " AUTHOR \\* MERGEFORMAT " });
    expect((sf!.simpleField as Record<string, unknown>).cachedValue).toBeUndefined();
  });

  it("parses a smartTag with uri, properties and children", () => {
    const opts = parseParagraphXml(
      `<w:smartTag w:element="Address" w:uri="http://schemas.example.com/addr">` +
        `<w:smartTagPr><w:attr w:name="type" w:val="home" w:uri="http://x"/></w:smartTagPr>` +
        `<w:r><w:t>123 Main St</w:t></w:r>` +
        `</w:smartTag>`,
    );

    const st = findChildByKey(opts, "smartTag");
    expect(st).toBeDefined();
    expect(st!.smartTag).toMatchObject({
      element: "Address",
      uri: "http://schemas.example.com/addr",
    });
    const stOpts = st!.smartTag as Record<string, unknown>;
    expect(stOpts.properties).toEqual([{ name: "type", val: "home", uri: "http://x" }]);
    expect(stOpts.children).toEqual([{ text: "123 Main St" }]);
  });

  it("parses an inline customXml with element, uri, customXmlPr and children", () => {
    const opts = parseParagraphXml(
      `<w:customXml w:element="Customer" w:uri="http://ns.example.com/cust">` +
        `<w:customXmlPr><w:placeholder w:val="Enter customer"/><w:attr w:name="id" w:val="42"/></w:customXmlPr>` +
        `<w:r><w:t>Acme</w:t></w:r>` +
        `</w:customXml>`,
    );

    const cx = findChildByKey(opts, "customXml");
    expect(cx).toBeDefined();
    expect(cx!.customXml).toMatchObject({
      element: "Customer",
      uri: "http://ns.example.com/cust",
    });
    const cxOpts = cx!.customXml as Record<string, unknown>;
    expect(cxOpts.customXmlPr).toEqual({
      placeholder: "Enter customer",
      attrs: [{ name: "id", val: "42" }],
    });
    expect(cxOpts.children).toEqual([{ text: "Acme" }]);
  });

  it("parses nested containers (customXml wrapping a smartTag)", () => {
    const opts = parseParagraphXml(
      `<w:customXml w:element="Root">` +
        `<w:smartTag w:element="Inner"><w:r><w:t>nested</w:t></w:r></w:smartTag>` +
        `</w:customXml>`,
    );

    const cx = findChildByKey(opts, "customXml");
    const cxOpts = cx!.customXml as Record<string, unknown>;
    const inner = cxOpts.children as Array<Record<string, unknown>>;
    expect(inner).toHaveLength(1);
    expect(inner[0].smartTag).toMatchObject({ element: "Inner" });
    expect((inner[0].smartTag as Record<string, unknown>).children).toEqual([{ text: "nested" }]);
  });

  it("drops a smartTag/customXml missing the required w:element", () => {
    const opts = parseParagraphXml(
      `<w:smartTag w:uri="http://x"><w:r><w:t>x</w:t></w:r></w:smartTag>`,
    );
    expect(findChildByKey(opts, "smartTag")).toBeUndefined();
  });
});

describe("bidirectional containers parse", () => {
  it("parses dir (w:dir with val and children)", () => {
    const opts = parseParagraphXml(`<w:dir w:val="rtl"><w:r><w:t>RTL text</w:t></w:r></w:dir>`);
    const d = findChildByKey(opts, "dir");
    expect(d).toBeDefined();
    expect(d!.dir).toMatchObject({ val: "rtl" });
    expect((d!.dir as Record<string, unknown>).children).toEqual([{ text: "RTL text" }]);
  });

  it("parses bdo (w:bdo with val and children)", () => {
    const opts = parseParagraphXml(`<w:bdo w:val="ltr"><w:r><w:t>text</w:t></w:r></w:bdo>`);
    const b = findChildByKey(opts, "bdo");
    expect(b!.bdo).toMatchObject({ val: "ltr" });
    expect((b!.bdo as Record<string, unknown>).children).toEqual([{ text: "text" }]);
  });
});
