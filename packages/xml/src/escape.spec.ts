import { describe, expect, it } from "vitest";

import { escapeXml, sanitizeXmlText } from "../src/escape";

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

describe("sanitizeXmlText", () => {
  it("preserves XML whitespace control characters", () => {
    expect(sanitizeXmlText("a\tb\nc\rd")).toBe("a\tb\nc\rd");
  });

  it("normalizes vertical tab and removes invalid control characters", () => {
    expect(sanitizeXmlText("a\u000Bb\u0000c\u0008d")).toBe("a\nbcd");
  });
});

describe("XML code points", () => {
  it("rejects invalid controls and lone surrogates during escaping", () => {
    expect(() => escapeXml("invalid\u0001")).toThrow(/U\+0001/);
    expect(() => escapeXml("surrogate\ud800")).toThrow(/U\+d800/);
  });
});
