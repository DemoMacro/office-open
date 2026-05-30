import { describe, expect, it } from "vitest";

import { escapeXml } from "../src/escape";

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
