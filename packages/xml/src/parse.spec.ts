import { describe, expect, it } from "vitest";

import { XmlParseError } from "../src/errors";
import { parse } from "../src/parse";
import { stringify } from "../src/stringify";
import { attr } from "../src/utils";

describe("parse", () => {
  it("throws structured errors for malformed XML", () => {
    expect(() => parse("<root>")).toThrowError(XmlParseError);
    try {
      parse("<root>unfinished");
    } catch (error) {
      expect(error).toBeInstanceOf(XmlParseError);
      expect((error as XmlParseError).name).toBe("XmlParseError");
    }
  });
  it("parses a simple element", () => {
    const result = parse("<w:t>Hello</w:t>");
    expect(result.elements?.[0]).toEqual({
      type: "element",
      name: "w:t",
      elements: [{ type: "text", text: "Hello" }],
    });
  });

  it("parses an element with attributes", () => {
    const result = parse('<w:pStyle w:val="Title"/>');
    expect(result.elements?.[0]).toEqual({
      type: "element",
      name: "w:pStyle",
      attributes: { "w:val": "Title" },
    });
  });

  it("parses nested elements", () => {
    const result = parse("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>");
    expect(result.elements?.[0]?.name).toBe("w:p");
    expect(result.elements?.[0]?.elements?.[0]?.name).toBe("w:r");
    expect(result.elements?.[0]?.elements?.[0]?.elements?.[0]?.name).toBe("w:t");
  });

  it("parses XML declaration", () => {
    const result = parse('<?xml version="1.0" encoding="UTF-8"?><root/>');
    expect(result.declaration).toEqual({
      attributes: { version: "1.0", encoding: "UTF-8" },
    });
  });

  it("parses and ignores processing instructions", () => {
    const source = '<?xml version="1.0"?><?mso-application progid="Word.Document"?><root/>';
    expect(parse(source).elements?.[0]).toEqual({
      type: "instruction",
      instruction: 'mso-application progid="Word.Document"',
    });
    expect(parse(source, { ignoreInstruction: true }).elements?.[0]).toMatchObject({
      type: "element",
      name: "root",
    });
  });

  it("parses standalone declaration", () => {
    const result = parse('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>', {});
    expect(result.declaration?.attributes?.standalone).toBe("yes");
  });

  it("ignores whitespace between elements by default", () => {
    const result = parse("<root>\n  <a>1</a>\n  <b>2</b>\n</root>");
    expect(result.elements?.[0]?.elements?.length).toBe(2);
    // No whitespace text nodes
    expect(result.elements?.[0]?.elements?.every((e: any) => e.type === "element")).toBe(true);
  });

  it("captures spaces between elements when option is set", () => {
    const result = parse("<root>\n  <a>1</a>\n  <b>2</b>\n</root>", {
      captureSpacesBetweenElements: true,
    });
    // Should have text nodes for whitespace
    expect(result.elements?.[0]?.elements?.some((e: any) => e.type === "text")).toBe(true);
  });

  it("parses self-closing tags", () => {
    const result = parse("<w:p/>");
    expect(result.elements?.[0]).toEqual({ type: "element", name: "w:p" });
  });

  it("parses CDATA sections", () => {
    const result = parse("<root><![CDATA[data]]></root>");
    expect(result.elements?.[0]?.elements?.[0]).toEqual({ type: "cdata", cdata: "data" });
  });

  it("parses comments", () => {
    const result = parse("<root><!-- comment --><a>1</a></root>");
    expect(result.elements?.[0]?.elements?.[0]).toEqual({
      type: "comment",
      comment: " comment ",
    });
  });

  it("parses attributes with single quotes", () => {
    const result = parse("<root attr='value'/>");
    expect(result.elements?.[0]?.attributes).toEqual({ attr: "value" });
  });

  it("omits attributes when ignoreAttributes is enabled", () => {
    const result = parse('<root attr="value"/>', { ignoreAttributes: true });
    expect(result.elements?.[0]?.attributes).toBeUndefined();
  });

  it("rejects mismatched and unclosed elements", () => {
    expect(() => parse("<a><b></a></b>")).toThrowError(
      "Malformed XML: expected closing tag </b>, got </a>.",
    );
    expect(() => parse("<a><b>")).toThrowError("Malformed XML: unclosed element <b>.");
  });

  it("rejects duplicate attributes", () => {
    expect(() => parse('<root x="1" x="2"/>')).toThrowError(
      "Malformed XML: duplicate attribute x.",
    );
  });

  it("rejects unclosed markup constructs", () => {
    expect(() => parse("<?xml ")).toThrowError("Malformed XML: unclosed processing instruction.");
    expect(() => parse("<!--")).toThrowError("Malformed XML: unclosed comment.");
    expect(() => parse("<![CDATA[")).toThrowError("Malformed XML: unclosed CDATA section.");
    expect(() => parse("<!DOCTYPE")).toThrowError("Malformed XML: unclosed DOCTYPE declaration.");
    expect(() => parse("<root")).toThrowError("Malformed XML: unclosed element <root>.");
    expect(() => parse("<root></root")).toThrowError("Malformed XML: unclosed closing tag.");
  });

  it("rejects malformed opening-tag attributes", () => {
    expect(() => parse('<root a="1" b></root>')).toThrowError(
      "Malformed XML: attribute b is missing a value.",
    );
    expect(() => parse("<root a=1></root>")).toThrowError(
      "Malformed XML: attribute a is missing a quoted value.",
    );
  });

  it("captures deferred content through self-closing same-name children", () => {
    const deferred = parse("<root><sheetData><sheetData/><row/></sheetData></root>", {
      deferElements: ["sheetData"],
    }).elements?.[0]?.elements?.[0];
    expect(deferred?.raw).toBe("<sheetData/><row/>");
    expect(stringify({ elements: [deferred!] })).toBe("<sheetData><sheetData/><row/></sheetData>");
  });

  it("captures deferred content when CDATA contains the close tag", () => {
    const parsed = parse(
      "<root><sheetData><row><v><![CDATA[</sheetData>]]></v></row></sheetData></root>",
      { deferElements: ["sheetData"] },
    );
    const deferred = parsed.elements?.[0]?.elements?.[0];
    expect(deferred?.raw).toBe("<row><v><![CDATA[</sheetData>]]></v></row>");
    expect(parsed.elements?.[0]?.name).toBe("root");
  });

  it("skips quoted attribute text while capturing deferred content", () => {
    const parsed = parse(
      `<root><sheetData><row a="</sheetData>"><b x='<'/></row></sheetData></root>`,
      { deferElements: ["sheetData"] },
    );
    expect(parsed.elements?.[0]?.elements?.[0]?.raw).toBe(`<row a="</sheetData>"><b x='<'/></row>`);
  });

  it("rejects unclosed deferred content", () => {
    expect(() =>
      parse("<root><sheetData><row/></root>", { deferElements: ["sheetData"] }),
    ).toThrowError("Malformed XML: unclosed element <sheetData>.");
  });

  it("matches namespace-equivalent opening and closing names", () => {
    const result = parse(
      '<w:a xmlns:w="http://example/w" xmlns:q="http://example/w" x="1"></q:a>',
      { normalizeNamespaces: { "http://example/w": "w" } },
    );
    expect(result.elements?.[0]).toMatchObject({ name: "w:a" });
  });

  it("renames native attribute raw values with namespaces", () => {
    const result = parse('<root xmlns:q="http://example/w" q:flag="False"/>', {
      nativeTypeAttributes: true,
      normalizeNamespaces: { "http://example/w": "w" },
    }).elements![0]!;
    expect(result.attributes).toEqual({
      "xmlns:w": "http://example/w",
      "w:flag": false,
    });
    expect(result.attributeRawValues).toEqual({ "w:flag": "False" });
  });

  it("rejects attributes that collide after namespace normalization", () => {
    expect(() =>
      parse('<root xmlns:w="http://example/w" xmlns:q="http://example/w" q:a="1" w:a="2"/>', {
        normalizeNamespaces: { "http://example/w": "w" },
      }),
    ).toThrowError("Malformed XML: duplicate attribute w:a.");
  });

  it("keeps the default namespace binding when another prefix canonicalizes to empty", () => {
    const result = parse('<root xmlns="http://example/default" xmlns:q="http://example/main"/>', {
      normalizeNamespaces: { "http://example/default": "d", "http://example/main": "" },
    }).elements![0]!;
    expect(result.name).toBe("d:root");
    expect(result.attributes).toEqual({
      "xmlns:d": "http://example/default",
      xmlns: "http://example/main",
    });
  });

  it("rejects canonical prefixes claimed by different namespaces", () => {
    expect(() =>
      parse(
        '<root xmlns:a="http://example/known" xmlns:w="http://example/unknown"><a:e/><w:e/></root>',
        { normalizeNamespaces: { "http://example/known": "w" } },
      ),
    ).toThrowError("Malformed XML: normalized prefix w has conflicting namespaces.");
  });

  it("allows a child to redeclare a normalized prefix", () => {
    const result = parse(
      '<root xmlns:w="http://example/known"><child xmlns:w="http://example/other"><w:e/></child><w:e/></root>',
      { normalizeNamespaces: { "http://example/known": "w" } },
    );
    const child = result.elements?.[0]?.elements?.[0];
    expect(result.elements?.[0]?.attributes).toEqual({
      "xmlns:w": "http://example/known",
    });
    expect(child?.attributes).toEqual({ "xmlns:w": "http://example/other" });
    expect(child?.elements?.[0]?.name).toBe("w:e");
    expect(result.elements?.[0]?.elements?.[1]?.name).toBe("w:e");
  });

  it("rejects unsupported and invalid character references", () => {
    expect(() => parse("<root>&nbsp;</root>")).toThrowError(
      "Malformed XML: unsupported or malformed entity.",
    );
    expect(() => parse("<root>&#0;</root>")).toThrowError(
      "Malformed XML: invalid character reference &#0;.",
    );
    expect(parse("<root>&#65;</root>").elements?.[0]?.elements?.[0]).toEqual({
      type: "text",
      text: "A",
    });
  });

  it("rejects raw characters outside XML 1.0", () => {
    expect(() => parse("<root>\u0000</root>")).toThrowError(
      "Malformed XML: invalid character U+0000.",
    );
  });

  it("captures a DOCTYPE internal subset as one declaration", () => {
    const source = `<!DOCTYPE r [ <!ENTITY greeting "hello"> ]><root/>`;
    const result = parse(source);
    expect(result.elements?.[0]).toMatchObject({
      type: "doctype",
      doctype: `r [ <!ENTITY greeting "hello"> ]`,
    });
    expect(stringify(result)).toBe(source);
  });

  it("parses empty element", () => {
    const result = parse("<root></root>");
    expect(result.elements?.[0]?.name).toBe("root");
    // xml-js does not set elements for empty non-self-closing tags
    expect(result.elements?.[0]?.elements).toBeUndefined();
  });

  it("parses mixed content (text + elements)", () => {
    const result = parse("<root>text<a>1</a>more</root>");
    const children = result.elements?.[0]?.elements;
    expect(children?.[0]).toEqual({ type: "text", text: "text" });
    expect(children?.[1]?.name).toBe("a");
    expect(children?.[2]).toEqual({ type: "text", text: "more" });
  });

  it("parses OOXML-style namespaced tags and attributes", () => {
    const result = parse(
      '<w:p w:val="1" xml:space="preserve"><w:r><w:t xml:space="preserve">Hello</w:t></w:r></w:p>',
    );
    expect(result.elements?.[0]?.name).toBe("w:p");
    expect(result.elements?.[0]?.attributes?.["w:val"]).toBe("1");
    expect(result.elements?.[0]?.attributes?.["xml:space"]).toBe("preserve");
  });

  it("ignores document type declarations", () => {
    const document = parse(
      '<?xml version="1.0"?><!DOCTYPE root><root><child>Value</child></root>',
      { ignoreDeclaration: true, ignoreDoctype: true },
    );
    expect(document.elements?.[0]).toMatchObject({ name: "root" });
  });
});

describe("native attribute parsing", () => {
  it("preserves the source spelling of boolean attributes", () => {
    const parsed = parse('<root flag="False"/>', { nativeTypeAttributes: true }).elements![0]!;
    expect(parsed.attributes!.flag).toBe(false);
    expect(attr(parsed, "flag")).toBe("False");
  });
});

describe("parse (custom)", () => {
  it("preserves leading/trailing spaces in text nodes that have content", () => {
    const result = parse('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
    expect(result.elements?.[0]?.elements?.[0]).toEqual({ type: "text", text: " PAGE " });
  });

  it("drops pure-whitespace text nodes (indentation) by default", () => {
    const result = parse("<root>\n  <a>1</a>\n</root>");
    const children = result.elements?.[0]?.elements ?? [];
    expect(children.every((e) => e.type === "element")).toBe(true);
  });

  it("keeps pure-whitespace nodes when captureSpaces is on", () => {
    const result = parse("<root>\n  <a>1</a>\n</root>", {
      captureSpacesBetweenElements: true,
    });
    const children = result.elements?.[0]?.elements ?? [];
    expect(children.some((e) => e.type === "text")).toBe(true);
  });

  it("unescapes attribute values (& < > \" ')", () => {
    const result = parse('<r a="x&lt;y&amp;z&quot;q&apos;w"/>');
    expect(result.elements?.[0]?.attributes?.a).toBe("x<y&z\"q'w");
  });

  it("decodes decimal and hex numeric character references", () => {
    const result = parse("<r>a&#65;b&#x42;c</r>");
    expect(result.elements?.[0]?.elements?.[0]).toEqual({ type: "text", text: "aAbBc" });
  });

  it("reassembles a CDATA section split on `]]>` into a single node", () => {
    // writeCdata serializes `a]]>b` as two adjacent CDATA sections.
    const result = parse("<r><![CDATA[a]]]]><![CDATA[>b]]></r>");
    const cdatas = (result.elements?.[0]?.elements ?? []).filter((e) => e.type === "cdata");
    expect(cdatas).toHaveLength(1);
    expect((cdatas[0] as { cdata?: string }).cdata).toBe("a]]>b");
  });

  it("preserves a whitespace-only text node under xml:space=preserve", () => {
    const result = parse('<w:t xml:space="preserve"> </w:t>');
    expect(result.elements?.[0]?.elements?.[0]).toEqual({ type: "text", text: " " });
  });
});
