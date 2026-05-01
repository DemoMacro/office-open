import { describe, expect, it } from "vitest";

import { xml2js } from "../src/parse";

describe("xml2js (parse)", () => {
    it("parses a simple element", () => {
        const result = xml2js("<w:t>Hello</w:t>", { compact: false });
        expect(result.elements?.[0]).toEqual({
            type: "element",
            name: "w:t",
            elements: [{ type: "text", text: "Hello" }],
        });
    });

    it("parses an element with attributes", () => {
        const result = xml2js('<w:pStyle w:val="Title"/>', { compact: false });
        expect(result.elements?.[0]).toEqual({
            type: "element",
            name: "w:pStyle",
            attributes: { "w:val": "Title" },
        });
    });

    it("parses nested elements", () => {
        const result = xml2js("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>", { compact: false });
        expect(result.elements?.[0].name).toBe("w:p");
        expect(result.elements?.[0].elements?.[0].name).toBe("w:r");
        expect(result.elements?.[0].elements?.[0].elements?.[0].name).toBe("w:t");
    });

    it("parses XML declaration", () => {
        const result = xml2js('<?xml version="1.0" encoding="UTF-8"?><root/>', { compact: false });
        expect(result.declaration).toEqual({
            attributes: { version: "1.0", encoding: "UTF-8" },
        });
    });

    it("parses standalone declaration", () => {
        const result = xml2js('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>', {
            compact: false,
        });
        expect(result.declaration?.attributes?.standalone).toBe("yes");
    });

    it("ignores whitespace between elements by default", () => {
        const result = xml2js("<root>\n  <a>1</a>\n  <b>2</b>\n</root>", { compact: false });
        expect(result.elements?.[0].elements?.length).toBe(2);
        // No whitespace text nodes
        expect(result.elements?.[0].elements?.every((e: any) => e.type === "element")).toBe(true);
    });

    it("captures spaces between elements when option is set", () => {
        const result = xml2js("<root>\n  <a>1</a>\n  <b>2</b>\n</root>", {
            compact: false,
            captureSpacesBetweenElements: true,
        });
        // Should have text nodes for whitespace
        expect(result.elements?.[0].elements?.some((e: any) => e.type === "text")).toBe(true);
    });

    it("parses self-closing tags", () => {
        const result = xml2js("<w:p/>", { compact: false });
        expect(result.elements?.[0]).toEqual({ type: "element", name: "w:p" });
    });

    it("parses CDATA sections", () => {
        const result = xml2js("<root><![CDATA[data]]></root>", { compact: false });
        expect(result.elements?.[0].elements?.[0]).toEqual({ type: "cdata", cdata: "data" });
    });

    it("parses comments", () => {
        const result = xml2js("<root><!-- comment --><a>1</a></root>", { compact: false });
        expect(result.elements?.[0].elements?.[0]).toEqual({
            type: "comment",
            comment: " comment ",
        });
    });

    it("parses attributes with single quotes", () => {
        const result = xml2js("<root attr='value'/>", { compact: false });
        expect(result.elements?.[0].attributes).toEqual({ attr: "value" });
    });

    it("parses empty element", () => {
        const result = xml2js("<root></root>", { compact: false });
        expect(result.elements?.[0].name).toBe("root");
        // xml-js does not set elements for empty non-self-closing tags
        expect(result.elements?.[0].elements).toBeUndefined();
    });

    it("parses mixed content (text + elements)", () => {
        const result = xml2js("<root>text<a>1</a>more</root>", { compact: false });
        const children = result.elements?.[0].elements;
        expect(children?.[0]).toEqual({ type: "text", text: "text" });
        expect(children?.[1]?.name).toBe("a");
        expect(children?.[2]).toEqual({ type: "text", text: "more" });
    });

    it("parses OOXML-style namespaced tags and attributes", () => {
        const result = xml2js(
            '<w:p w:val="1" xml:space="preserve"><w:r><w:t xml:space="preserve">Hello</w:t></w:r></w:p>',
            { compact: false },
        );
        expect(result.elements?.[0].name).toBe("w:p");
        expect(result.elements?.[0].attributes?.["w:val"]).toBe("1");
        expect(result.elements?.[0].attributes?.["xml:space"]).toBe("preserve");
    });
});
