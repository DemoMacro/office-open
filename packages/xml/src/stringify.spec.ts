import { describe, expect, it } from "vitest";

import { js2xml } from "../src/stringify";
import type { Element } from "../src/types";

describe("js2xml (stringify)", () => {
    it("stringifies a simple element", () => {
        const el: Element = {
            elements: [
                { type: "element", name: "w:t", elements: [{ type: "text", text: "Hello" }] },
            ],
        };
        expect(js2xml(el)).toBe("<w:t>Hello</w:t>");
    });

    it("stringifies an element with attributes", () => {
        const el: Element = {
            elements: [{ type: "element", name: "w:pStyle", attributes: { "w:val": "Title" } }],
        };
        expect(js2xml(el)).toBe('<w:pStyle w:val="Title"/>');
    });

    it("stringifies nested elements", () => {
        const el: Element = {
            elements: [
                {
                    type: "element",
                    name: "w:p",
                    attributes: { "w:val": "1" },
                    elements: [
                        {
                            type: "element",
                            name: "w:r",
                            elements: [
                                {
                                    type: "element",
                                    name: "w:t",
                                    elements: [{ type: "text", text: "Hello" }],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        expect(js2xml(el)).toBe('<w:p w:val="1"><w:r><w:t>Hello</w:t></w:r></w:p>');
    });

    it("writes XML declaration", () => {
        const el: Element = {
            declaration: { attributes: { version: "1.0", encoding: "UTF-8", standalone: "yes" } },
            elements: [{ type: "element", name: "root" }],
        };
        expect(js2xml(el)).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
    });

    it("handles empty elements (self-closing)", () => {
        const el: Element = { elements: [{ type: "element", name: "root" }] };
        expect(js2xml(el)).toBe("<root/>");
    });

    it("forces closing tag for xml:space=preserve", () => {
        const el: Element = {
            elements: [
                {
                    type: "element",
                    name: "w:t",
                    attributes: { "xml:space": "preserve" },
                },
            ],
        };
        expect(js2xml(el)).toBe('<w:t xml:space="preserve"></w:t>');
    });

    it("handles CDATA", () => {
        const el: Element = {
            elements: [
                { type: "element", name: "root", elements: [{ type: "cdata", cdata: "data" }] },
            ],
        };
        expect(js2xml(el)).toBe("<root><![CDATA[data]]></root>");
    });

    it("handles comments", () => {
        const el: Element = {
            elements: [
                { type: "element", name: "root", elements: [{ type: "comment", comment: "note" }] },
            ],
        };
        expect(js2xml(el)).toBe("<root><!--note--></root>");
    });

    it("handles text nodes with special characters", () => {
        const el: Element = {
            elements: [
                { type: "element", name: "w:t", elements: [{ type: "text", text: "a & b" }] },
            ],
        };
        expect(js2xml(el)).toBe("<w:t>a &amp; b</w:t>");
    });

    it("desanitizes text before sanitizing to avoid double-escape", () => {
        const el: Element = {
            elements: [
                { type: "element", name: "w:t", elements: [{ type: "text", text: "&amp;" }] },
            ],
        };
        expect(js2xml(el)).toBe("<w:t>&amp;</w:t>");
    });

    it("applies attributeValueFn", () => {
        const el: Element = {
            elements: [
                {
                    type: "element",
                    name: "w:p",
                    attributes: { "w:val": "a&b" },
                },
            ],
        };
        const result = js2xml(el, {
            attributeValueFn: (value) => value.replace(/&/g, "&amp;"),
        });
        expect(result).toBe('<w:p w:val="a&amp;b"/>');
    });

    it("converts numeric attribute values to string", () => {
        const el: Element = {
            elements: [{ type: "element", name: "w:val", attributes: { count: 42 } }],
        };
        expect(js2xml(el)).toBe('<w:val count="42"/>');
    });

    it("handles element with no name", () => {
        const el: Element = { elements: [{ type: "text", text: "just text" }] };
        expect(js2xml(el)).toBe("just text");
    });
});
