import { describe, expect, it } from "vite-plus/test";

import { SharedStrings } from "./shared-strings";

const context = { stack: [] } as any;

describe("SharedStrings", () => {
  it("register() returns incrementing indices", () => {
    const ss = new SharedStrings();
    expect(ss.register("A")).toBe(0);
    expect(ss.register("B")).toBe(1);
    expect(ss.register("C")).toBe(2);
  });

  it("register() deduplicates identical strings", () => {
    const ss = new SharedStrings();
    ss.register("A");
    ss.register("B");
    expect(ss.register("A")).toBe(0); // same index
    expect(ss.register("B")).toBe(1); // same index
  });

  it("count reflects total registered strings (with dedup)", () => {
    const ss = new SharedStrings();
    ss.register("A");
    ss.register("B");
    ss.register("A"); // dup
    expect(ss.count).toBe(2);
  });

  // ── toXml path ──

  describe("toXml", () => {
    it("produces valid sst XML with namespace", () => {
      const ss = new SharedStrings();
      ss.register("Hello");
      const xml = ss.toXml(context);

      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
      expect(xml).toContain("<si><t>Hello</t></si>");
      expect(xml).toContain('count="1"');
      expect(xml).toContain('uniqueCount="1"');
    });

    it("escapes XML special characters", () => {
      const ss = new SharedStrings();
      ss.register("<b>&");
      const xml = ss.toXml(context);
      expect(xml).toContain("<t>&lt;b&gt;&amp;</t>");
    });
  });
});
