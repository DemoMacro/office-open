import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Inline metadata carriers never touch the read context, so an empty mock
// suffices.
const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W16SE_NS = 'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseParagraph(el, readCtx) as { children?: unknown[] };
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
    expect(cxOpts.properties).toEqual({
      placeholder: "Enter customer",
      attributes: [{ name: "id", val: "42" }],
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
    expect(inner[0]?.smartTag).toMatchObject({ element: "Inner" });
    expect((inner[0]!.smartTag as Record<string, unknown>).children).toEqual([{ text: "nested" }]);
  });

  it("drops a smartTag/customXml missing the required w:element", () => {
    const opts = parseParagraphXml(
      `<w:smartTag w:uri="http://x"><w:r><w:t>x</w:t></w:r></w:smartTag>`,
    );
    expect(findChildByKey(opts, "smartTag")).toBeUndefined();
  });

  it("round-trips an empty run inside customXml as <w:r/>", () => {
    const opts = parseParagraphXml(`<w:customXml w:element="field"><w:r/></w:customXml>`);

    const cx = findChildByKey(opts, "customXml");
    expect(cx).toBeDefined();
    expect((cx!.customXml as Record<string, unknown>).children).toEqual([{}]);

    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain('<w:customXml w:element="field"><w:r/></w:customXml>');
  });

  it("round-trips an empty run as a paragraph child", () => {
    const opts = parseParagraphXml(`<w:r/>`);
    expect(opts.children).toEqual([{}]);

    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain("<w:r/>");
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

describe("Office 2016 symbol round-trip", () => {
  it("re-emits the extension symbol without downgrading it to w:sym", () => {
    const doc = parseXml(
      `<w:p ${W_NS} ${W16SE_NS}><w:r><w16se:sym w:font="Webdings" w:char="F04E"/></w:r></w:p>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseParagraph(el, readCtx);

    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w16se:sym w:char="F04E" w:font="Webdings"/>',
    );
  });
});

describe("ruby annotation parse", () => {
  it("parses ruby as formatted content inside a run", () => {
    const opts = parseParagraphXml(
      `<w:r><w:ruby><w:rubyPr>` +
        `<w:rubyAlign w:val="center"/><w:hps w:val="20"/><w:hpsRaise w:val="20"/>` +
        `<w:hpsBaseText w:val="40"/><w:lid w:val="ja-JP"/><w:dirty w:val="0"/>` +
        `</w:rubyPr>` +
        `<w:rt><w:r><w:rPr><w:b/></w:rPr><w:t>fu</w:t></w:r><w:r><w:t>ri</w:t></w:r></w:rt>` +
        `<w:rubyBase><w:r><w:rPr><w:i/></w:rPr><w:t>base</w:t></w:r></w:rubyBase>` +
        `</w:ruby></w:r>`,
    );
    const run = opts.children?.[0] as Record<string, unknown>;
    const runChildren = run.children as Array<Record<string, unknown>>;
    expect(runChildren).toHaveLength(1);
    expect(runChildren[0]!.ruby).toEqual({
      properties: {
        alignment: "center",
        fontSize: 10,
        raise: 10,
        baseFontSize: 20,
        languageId: "ja-JP",
        dirty: false,
      },
      text: { children: [{ bold: true, text: "fu" }, { text: "ri" }] },
      base: { children: [{ italic: true, text: "base" }] },
    });

    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain("<w:r><w:ruby><w:rubyPr>");
    expect(xml).toContain('<w:dirty w:val="off"/>');
    expect(xml).toContain(
      '<w:rt><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">fu</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">ri</w:t></w:r></w:rt>',
    );
    expect(xml).toContain(
      '<w:rubyBase><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">base</w:t></w:r></w:rubyBase>',
    );
    expect(xml).not.toContain("<w:r><w:ruby><w:r><w:ruby>");
  });
});

describe("hyperlink relationships", () => {
  const hlWriteCtx = {
    viewWrapper: {
      relationships: {
        add: (type: string, target: string, mode: string) => {
          relationships.push({ type, target, mode });
          return relationships.length;
        },
      },
    },
  } as never;
  let relationships: Array<{ type: string; target: string; mode: string }>;

  it("registers one relationship per hyperlink element even for a repeated URL", () => {
    relationships = [];
    const xml = stringifyParagraph(
      {
        children: [
          { hyperlink: { url: "https://example.com/a", children: ["one"] } },
          { hyperlink: { url: "https://example.com/a", children: ["two"] } },
        ],
      },
      hlWriteCtx,
    );
    expect(relationships).toHaveLength(2);
    expect(
      relationships.every(
        (r) =>
          r.type ===
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" &&
          r.target === "https://example.com/a" &&
          r.mode === "External",
      ),
    ).toBe(true);
    expect(xml).toContain('r:id="rId1"');
    expect(xml).toContain('r:id="rId2"');
  });
});
