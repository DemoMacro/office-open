import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { SharedStrings, sharedStringsDesc } from "./shared-strings";

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

  it("round-trips the source reference count separately from unique entries", () => {
    const ss = new SharedStrings();
    ss.register("A");
    ss.register("B");
    ss.setSourceCount(7);
    expect(ss.toDescriptorOptions()).toEqual({ entries: ["A", "B"], count: 7 });
  });

  it("omits an absent source reference count", () => {
    const ss = new SharedStrings();
    ss.register("A");
    const xml = ss.toDescriptorOptions();
    expect(xml.count).toBeUndefined();
    expect(sharedStringsDesc.stringify(xml, {} as never)).not.toContain('count="1"');
    expect(xml.uniqueCount).toBeUndefined();
    expect(sharedStringsDesc.stringify(xml, {} as never)).not.toContain('uniqueCount="1"');
  });

  it("round-trips an explicit unique count separately from entries", () => {
    const ss = new SharedStrings();
    ss.register("A");
    ss.setSourceUniqueCount(9);
    expect(ss.toDescriptorOptions()).toEqual({ entries: ["A"], uniqueCount: 9 });
  });

  it("preserves the index of an empty shared-string item", () => {
    const doc = parseXml(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        "<si/><si><t>A</t></si></sst>",
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const result = sharedStringsDesc.parse(root, {} as never);
    expect(result.entries).toEqual([{}, "A"]);
    expect(sharedStringsDesc.stringify(result, {} as never)).toContain(
      "<si></si><si><t>A</t></si>",
    );
  });

  // ── toXml path ──

  describe("serialize", () => {
    it("produces valid sst XML with namespace", () => {
      const ss = new SharedStrings();
      ss.register("Hello");
      const xml = ss.serialize();

      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
      expect(xml).toContain("<si><t>Hello</t></si>");
      expect(xml).toContain('count="1"');
      expect(xml).toContain('uniqueCount="1"');
    });

    it("uses the source reference count for sst/@count", () => {
      const ss = new SharedStrings();
      ss.register("A");
      ss.setSourceCount(4);
      expect(ss.serialize()).toContain('count="4"');
      expect(ss.serialize()).toContain('uniqueCount="1"');
    });

    it("escapes XML special characters", () => {
      const ss = new SharedStrings();
      ss.register("<b>&");
      const xml = ss.serialize();
      expect(xml).toContain("<t>&lt;b&gt;&amp;</t>");
    });

    it("marks whitespace-edged text with xml:space=preserve", () => {
      const ss = new SharedStrings();
      ss.register(" padded ");
      const xml = ss.serialize();
      expect(xml).toContain('<t xml:space="preserve"> padded </t>');
    });

    it("leaves tight text without the preserve attribute", () => {
      const ss = new SharedStrings();
      ss.register("tight");
      const xml = ss.serialize();
      expect(xml).toContain("<t>tight</t>");
      expect(xml).not.toContain("xml:space");
    });
  });
});
