import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Inline SDTs never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseParagraph(el, readCtx) as { children?: unknown[] };
}

function findInlineSdt(opts: { children?: unknown[] }): Record<string, unknown> | undefined {
  return opts.children?.find(
    (c) => c !== null && typeof c === "object" && "sdt" in (c as Record<string, unknown>),
  ) as Record<string, unknown> | undefined;
}

describe("inline SDT (CT_SdtRun) parse", () => {
  it("parses a plain-text inline SDT with run content", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:tag w:val="name"/><w:text/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>Jane</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    const sdt = child!.sdt as Record<string, unknown>;
    expect(sdt.properties).toMatchObject({ alias: "Name", tag: "name" });
    expect(sdt.children).toBeDefined();
    expect((sdt.children as unknown[]).length).toBeGreaterThan(0);
  });

  it("parses an inline SDT alongside surrounding runs (position preserved)", () => {
    const opts = parseParagraphXml(
      "<w:r><w:t>before </w:t></w:r>" +
        '<w:sdt><w:sdtPr><w:alias w:val="X"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>mid</w:t></w:r></w:sdtContent></w:sdt>" +
        "<w:r><w:t> after</w:t></w:r>",
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({ alias: "X" });
  });

  it("parses inline SDT end properties (w:sdtEndPr)", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="E"/></w:sdtPr>' +
        "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>" +
        "<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect((child!.sdt as Record<string, unknown>).endProperties).toBeDefined();
  });

  it("preserves inline SDT properties presence", () => {
    const absent = parseParagraphXml(
      "<w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    expect((findInlineSdt(absent)!.sdt as Record<string, unknown>).properties).toEqual({
      propertiesPresent: false,
    });
    expect(stringifyParagraph(absent as never, writeCtx)).not.toContain("<w:sdtPr");

    const empty = parseParagraphXml(
      "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    expect((findInlineSdt(empty)!.sdt as Record<string, unknown>).properties).toEqual({
      propertiesPresent: true,
    });
    expect(stringifyParagraph(empty as never, writeCtx)).toContain("<w:sdtPr/>");
  });

  it("parses an inline checkbox SDT (w14:checkbox) — sdtContent is a symbol run, not a w:p", () => {
    const opts = parseParagraphXml(
      `<w:sdt><w:sdtPr ${W14}><w:alias w:val="C"/><w14:checkbox>` +
        '<w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>' +
        '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>',
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({
      alias: "C",
      checkbox: { checked: true },
    });
  });

  it("serializes cached inline checkbox SDT children without generated fonts", () => {
    const opts = parseParagraphXml(
      `<w:sdt><w:sdtPr ${W14}><w:alias w:val="C"/><w14:checkbox>` +
        '<w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>' +
        "<w:sdtContent><w:r><w:rPr></w:rPr></w:r><w:r><w:rPr></w:rPr>" +
        "<w:t>☒</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const sdt = findInlineSdt(opts)!.sdt as { children?: unknown[] };
    expect(sdt.children).toHaveLength(2);
    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain("<w:t>☒</w:t>");
    expect(xml).not.toContain("MS Gothic");
    expect(xml).not.toContain("<w:rFonts");
  });

  it("parses a nested inline SDT (SDT inside SDT content)", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Outer"/></w:sdtPr>' +
        '<w:sdtContent><w:sdt><w:sdtPr><w:alias w:val="Inner"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>deep</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt>",
    );
    const outer = findInlineSdt(opts);
    const outerChildren = (outer!.sdt as Record<string, unknown>).children as unknown[];
    const inner = outerChildren.find(
      (c) => c !== null && typeof c === "object" && "sdt" in (c as Record<string, unknown>),
    ) as Record<string, unknown> | undefined;
    expect(inner).toBeDefined();
    expect((inner!.sdt as Record<string, unknown>).properties).toMatchObject({ alias: "Inner" });
  });
});

describe("inline SDT run identity", () => {
  it("preserves rsids on runs inside sdtContent", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Identity"/></w:sdtPr>' +
        '<w:sdtContent><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD"><w:t>Jane</w:t></w:r>' +
        "</w:sdtContent></w:sdt>",
    );
    const sdt = findInlineSdt(opts)!.sdt as Record<string, unknown>;
    expect((sdt.children as Record<string, unknown>[])[0]).toMatchObject({
      text: "Jane",
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
    });
    expect(stringifyParagraph(opts as never, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });

  it("preserves run properties on empty runs inside sdtContent", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Styled"/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:r>' +
        "<w:r><w:t>Jane</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const children = (findInlineSdt(opts)!.sdt as Record<string, unknown>).children as Record<
      string,
      unknown
    >[];
    expect(children).toHaveLength(2);
    expect(children[0]).toMatchObject({ language: { value: "de-DE" } });
    expect(stringifyParagraph(opts as never, writeCtx)).toContain('<w:lang w:val="de-DE"/>');
  });
});

describe("SDT properties fidelity (CT_SdtPr)", () => {
  it("round-trips the sdtPr run properties (leading w:rPr)", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:rPr><w:color w:val="76923C"/></w:rPr>' +
        '<w:alias w:val="Company"/><w:dataBinding w:xpath="/ns0:Properties[1]/ns0:Company[1]" ' +
        'w:storeItemID="{6668398D-A668-4E3E-A5EB-62B293D839F1}"/><w:text/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>Microsoft</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({
      alias: "Company",
      runProperties: { color: "76923C" },
    });
  });

  it("preserves typed sdtPr child order", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:id w:val="42"/><w:alias w:val="Ordered"/><w:text/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const sdt = findInlineSdt(opts)!.sdt as Record<string, unknown>;
    expect(sdt.properties).toMatchObject({
      id: 42,
      alias: "Ordered",
      childOrder: ["id", "alias", "text"],
    });
    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml.indexOf("<w:id")).toBeLessThan(xml.indexOf("<w:alias"));
    expect(xml.indexOf("<w:alias")).toBeLessThan(xml.indexOf("<w:text"));
  });

  it("round-trips w:docPartUnique inside docPartObj", () => {
    const opts = parseParagraphXml(
      "<w:sdt><w:sdtPr><w:docPartObj>" +
        '<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>' +
        "</w:sdtPr><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({
      docPartObj: { gallery: "Table of Contents", unique: true },
    });
  });

  it("round-trips explicit docPartUnique lexical values", () => {
    const opts = parseParagraphXml(
      "<w:sdt><w:sdtPr><w:docPartObj>" +
        '<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique w:val="true"/>' +
        "</w:docPartObj></w:sdtPr><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({
      docPartObj: { unique: true, uniqueRaw: "true" },
    });
    expect(stringifyParagraph(opts as never, writeCtx)).toContain(
      '<w:docPartUnique w:val="true"/>',
    );
  });

  it("round-trips an empty placeholder and explicit multiline lexical value", () => {
    const opts = parseParagraphXml(
      "<w:sdt><w:sdtPr><w:placeholder/>" +
        '<w:text w:multiLine="0"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    expect((findInlineSdt(opts)!.sdt as Record<string, unknown>).properties).toMatchObject({
      placeholder: {},
      text: { multiLine: false, multiLineRaw: "0" },
    });
    const xml = stringifyParagraph(opts as never, writeCtx);
    expect(xml).toContain("<w:placeholder/>");
    expect(xml).toContain('<w:text w:multiLine="0"/>');
  });

  it("round-trips the date SDT dateFormat attribute", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>2026-01-02</w:t></w:r></w:sdtContent></w:sdt>",
    );
    expect((findInlineSdt(opts)!.sdt as Record<string, any>).properties.date).toMatchObject({
      dateFormat: "yyyy-MM-dd",
    });
    expect(stringifyParagraph(opts as never, writeCtx)).toContain(
      '<w:dateFormat w:val="yyyy-MM-dd"/>',
    );
  });
});
