import { describe, expect, it } from "vitest";
import xmlOriginal from "xml";
import { xml2js as xml2jsOriginal, js2xml as js2xmlOriginal } from "xml-js";

import { toElement } from "../src/convert";
import { xml2js } from "../src/parse";
import { xml } from "../src/serialize";
import { js2xml } from "../src/stringify";
import type { Element } from "../src/types";

// ── xml (serialize): output equivalence ──

describe("xml: output matches original xml package", () => {
    it("simple text element", () => {
        const input = { "w:t": "Hello World" };
        expect(xml(input)).toBe(xmlOriginal(input));
    });

    it("nested elements with attributes", () => {
        const input = {
            "w:p": [
                { _attr: { "w:val": "1" } },
                { "w:r": [{ "w:t": "Bold text" }] },
                { "w:r": [{ "w:t": "Normal text" }] },
            ],
        };
        expect(xml(input)).toBe(xmlOriginal(input));
    });

    it("self-closing element", () => {
        const input = { "w:pStyle": null };
        expect(xml(input)).toBe(xmlOriginal(input));
    });

    it("element with CDATA", () => {
        const input = { "w:t": { _cdata: "raw <data>" } };
        expect(xml(input)).toBe(xmlOriginal(input));
    });

    it("multiple root elements", () => {
        const input = [{ "w:t": "A" }, { "w:t": "B" }];
        expect(xml(input)).toBe(xmlOriginal(input));
    });

    it("with indentation", () => {
        const input = {
            "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Bold" }] }],
        };
        expect(xml(input, { indent: true })).toBe(xmlOriginal(input, { indent: "    " }));
    });

    it("with custom indent string", () => {
        const input = { "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Bold" }] }] };
        expect(xml(input, { indent: "\t" })).toBe(xmlOriginal(input, { indent: "\t" }));
    });

    it("with declaration", () => {
        const input = { "w:document": [{ "w:body": [] }] };
        expect(xml(input, { declaration: { encoding: "UTF-8" } })).toBe(
            xmlOriginal(input, { declaration: { encoding: "UTF-8" } }),
        );
    });

    it("with declaration and standalone", () => {
        const input = { "w:document": [{ "w:body": [] }] };
        expect(xml(input, { declaration: { encoding: "UTF-8", standalone: "yes" } })).toBe(
            xmlOriginal(input, { declaration: { encoding: "UTF-8", standalone: "yes" } }),
        );
    });

    it("escaped special characters", () => {
        const input = { "w:t": "a & b < c > d \"e\" 'f'" };
        expect(xml(input)).toBe(xmlOriginal(input));
    });
});

// ── xml2js (parse): output equivalence ──

describe("xml2js: output matches original xml-js", () => {
    it("simple element", () => {
        const xmlStr = "<w:t>Hello</w:t>";
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("nested elements with attributes", () => {
        const xmlStr = '<w:p w:val="1"><w:r><w:t>Bold</w:t></w:r></w:p>';
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("self-closing element", () => {
        const xmlStr = '<w:pStyle w:val="Title"/>';
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("XML declaration", () => {
        const xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>';
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("CDATA section", () => {
        const xmlStr = "<w:t><![CDATA[raw <data>]]></w:t>";
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("comment", () => {
        const xmlStr = "<root><!-- comment --><child/></root>";
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("complex OOXML document", () => {
        const xmlStr =
            '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p></w:body></w:document>';
        const ours = xml2js(xmlStr, { compact: false });
        const original = xml2jsOriginal(xmlStr, { compact: false });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("captureSpacesBetweenElements", () => {
        const xmlStr = "<root>  <a/>  <b/></root>";
        const ours = xml2js(xmlStr, { compact: false, captureSpacesBetweenElements: true });
        const original = xml2jsOriginal(xmlStr, {
            compact: false,
            captureSpacesBetweenElements: true,
        });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("ignoreDeclaration", () => {
        const xmlStr = '<?xml version="1.0"?><root/>';
        const ours = xml2js(xmlStr, { compact: false, ignoreDeclaration: true });
        const original = xml2jsOriginal(xmlStr, { compact: false, ignoreDeclaration: true });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });

    it("trim", () => {
        const xmlStr = "<root>  text  </root>";
        const ours = xml2js(xmlStr, { compact: false, trim: true });
        const original = xml2jsOriginal(xmlStr, { compact: false, trim: true });
        expect(JSON.stringify(ours)).toBe(JSON.stringify(original));
    });
});

// ── js2xml (stringify): output equivalence ──

describe("js2xml: output matches original xml-js", () => {
    it("simple element", () => {
        const element: Element = {
            elements: [
                { type: "element", name: "w:t", elements: [{ type: "text", text: "Hello" }] },
            ],
        };
        expect(js2xml(element)).toBe(js2xmlOriginal(element));
    });

    it("attributes", () => {
        const element: Element = {
            elements: [{ type: "element", name: "w:p", attributes: { "w:val": "Title" } }],
        };
        expect(js2xml(element)).toBe(js2xmlOriginal(element));
    });

    it("nested elements", () => {
        const element: Element = {
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
        expect(js2xml(element)).toBe(js2xmlOriginal(element));
    });

    it("self-closing with preserve", () => {
        const element: Element = {
            elements: [{ type: "element", name: "w:p", attributes: { "xml:space": "preserve" } }],
        };
        expect(js2xml(element)).toBe(js2xmlOriginal(element));
    });

    it("with attributeValueFn", () => {
        const element: Element = {
            elements: [
                {
                    type: "element",
                    name: "w:p",
                    attributes: { "w:val": "1&2" },
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
        const fn = (value: string) => value.replace(/&/g, "&amp;");
        expect(js2xml(element, { attributeValueFn: fn })).toBe(
            js2xmlOriginal(element, { attributeValueFn: fn }),
        );
    });

    it("with spaces indentation", () => {
        const element: Element = {
            elements: [
                {
                    type: "element",
                    name: "w:p",
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
        expect(js2xml(element, { spaces: 2 })).toBe(js2xmlOriginal(element, { spaces: 2 }));
    });
});

// ── toElement: structural equivalence with bridge ──

describe("toElement: structure matches xml→xml2js bridge", () => {
    it("simple text element", () => {
        const input = { "w:t": "Hello" };
        const direct = toElement(input);
        const bridged = xml2js(xml(input), { compact: false });
        expect(direct.name).toBe(bridged.elements?.[0]?.name);
        expect(direct.elements?.[0]?.text).toBe(bridged.elements?.[0]?.elements?.[0]?.text);
    });

    it("nested elements", () => {
        const input = {
            "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Bold" }] }],
        };
        const direct = toElement(input);
        const bridged = xml2js(xml(input), { compact: false });
        expect(direct.name).toBe(bridged.elements?.[0]?.name);
        expect(direct.attributes).toEqual(bridged.elements?.[0]?.attributes);
        expect(direct.elements?.[0]?.name).toBe(bridged.elements?.[0]?.elements?.[0]?.name);
    });
});

// ── Round-trip: xml() → xml2js() → js2xml() ──

describe("Round-trip: xml → xml2js → js2xml", () => {
    it("preserves simple content", () => {
        const input = { "w:t": "Hello World" };
        const output = js2xml(xml2js(xml(input), { compact: false }));
        expect(output).toBe("<w:t>Hello World</w:t>");
    });

    it("preserves nested structure", () => {
        const input = {
            "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Bold text content" }] }],
        };
        const output = js2xml(xml2js(xml(input), { compact: false }));
        expect(output).toBe('<w:p w:val="1"><w:r><w:t>Bold text content</w:t></w:r></w:p>');
    });

    it("preserves escaped characters", () => {
        const input = { "w:t": "a & b < c > d" };
        const output = js2xml(xml2js(xml(input), { compact: false }));
        expect(output).toBe("<w:t>a &amp; b &lt; c &gt; d</w:t>");
    });
});
