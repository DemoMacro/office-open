import { describe, expect, it } from "vitest";

import { xml } from "../src/serialize";

describe("xml (serialize)", () => {
    it("serializes a simple element", () => {
        expect(xml({ "w:t": "Hello" })).toBe("<w:t>Hello</w:t>");
    });

    it("serializes an element with attributes", () => {
        expect(xml({ "w:pStyle": [{ _attr: { "w:val": "Title" } }] })).toBe(
            '<w:pStyle w:val="Title"/>',
        );
    });

    it("serializes nested elements", () => {
        const input = {
            "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Hello" }] }],
        };
        expect(xml(input)).toBe('<w:p w:val="1"><w:r><w:t>Hello</w:t></w:r></w:p>');
    });

    it("serializes multiple top-level elements", () => {
        expect(xml([{ a: "1" }, { b: "2" }])).toBe("<a>1</a><b>2</b>");
    });

    it("handles null values", () => {
        expect(xml({ "w:p": null })).toBe("<w:p/>");
    });

    it("handles empty array", () => {
        expect(xml({ "w:p": [] })).toBe("<w:p></w:p>");
    });

    it("escapes text content", () => {
        expect(xml({ "w:t": "a & b" })).toBe("<w:t>a &amp; b</w:t>");
    });

    it("escapes attribute values", () => {
        expect(xml({ "w:p": [{ _attr: { "w:val": 'a"b' } }] })).toBe('<w:p w:val="a&quot;b"/>');
    });

    it("handles CDATA", () => {
        const input = { "w:t": { _cdata: "data with ]]> inside" } };
        expect(xml(input)).toBe("<w:t><![CDATA[data with ]]]]><![CDATA[> inside]]></w:t>");
    });

    it("adds XML declaration", () => {
        const result = xml({ "w:p": [] }, { declaration: { encoding: "UTF-8" } });
        expect(result.startsWith("<?xml")).toBe(true);
        expect(result.includes('encoding="UTF-8"')).toBe(true);
    });

    it("adds standalone to declaration", () => {
        const result = xml(
            { "w:p": [] },
            { declaration: { encoding: "UTF-8", standalone: "yes" } },
        );
        expect(result).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p></w:p>');
    });

    it("handles indent as string", () => {
        const input = { "w:p": [{ "w:r": [{ "w:t": "Hello" }] }] };
        const result = xml(input, { indent: "  " });
        expect(result).toBe("<w:p>\n  <w:r>\n    <w:t>Hello</w:t>\n  </w:r>\n</w:p>");
    });

    it("handles indent as true (default 4 spaces)", () => {
        const input = { "w:p": [{ "w:r": [{ "w:t": "Hello" }] }] };
        const result = xml(input, true);
        expect(result).toContain("    ");
    });

    it("handles indent as false/undefined (no indent)", () => {
        const input = { "w:p": [{ "w:r": [{ "w:t": "Hello" }] }] };
        const result = xml(input, false);
        expect(result).toBe("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>");
    });

    it("handles numeric attribute values", () => {
        expect(xml({ "w:val": [{ _attr: { count: 42 } }] })).toBe('<w:val count="42"/>');
    });

    it("handles boolean attribute values", () => {
        expect(xml({ "w:val": [{ _attr: { active: true } }] })).toBe('<w:val active="true"/>');
    });
});
