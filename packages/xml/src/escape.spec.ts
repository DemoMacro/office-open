import { describe, expect, it } from "vitest";

import { escapeXml, escapeAttributeValue } from "../src/escape";

describe("escapeXml", () => {
    it("escapes all five XML special characters", () => {
        expect(escapeXml('Tom & Jerry <"test">')).toBe("Tom &amp; Jerry &lt;&quot;test&quot;&gt;");
    });

    it("escapes single quotes", () => {
        expect(escapeXml("it's")).toBe("it&apos;s");
    });

    it("returns empty string for empty input", () => {
        expect(escapeXml("")).toBe("");
    });

    it("does not modify strings without special characters", () => {
        expect(escapeXml("Hello World")).toBe("Hello World");
    });

    it("escapes numbers converted to strings", () => {
        expect(escapeXml(String(42))).toBe("42");
    });
});

describe("escapeAttributeValue", () => {
    it("escapes all five characters", () => {
        expect(escapeAttributeValue("a&b<c>d\"e'f")).toBe("a&amp;b&lt;c&gt;d&quot;e&apos;f");
    });

    it("preserves already-escaped entities", () => {
        expect(escapeAttributeValue("&amp;&lt;")).toBe("&amp;&lt;");
    });

    it("escapes unescaped ampersand but not named entities", () => {
        expect(escapeAttributeValue("a & b &lt; c")).toBe("a &amp; b &lt; c");
    });

    it("handles empty string", () => {
        expect(escapeAttributeValue("")).toBe("");
    });

    it("converts non-string values", () => {
        expect(escapeAttributeValue(String(42))).toBe("42");
    });
});
