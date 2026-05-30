import { Formatter } from "@export/formatter";
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

  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("empty table produces sst with count=0", () => {
      const tree = new Formatter().format(new SharedStrings());
      expect(tree).to.deep.equal({
        sst: [
          {
            _attr: {
              xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              count: 0,
              uniqueCount: 0,
            },
          },
        ],
      });
    });

    it("produces si elements for each string", () => {
      const ss = new SharedStrings();
      ss.register("Hello");
      ss.register("World");

      const tree = new Formatter().format(ss);
      const sst = tree["sst"] as any[];

      // _attr + 2 si elements
      expect(sst).toHaveLength(3);
      expect(sst[1]).to.deep.equal({ si: [{ t: ["Hello"] }] });
      expect(sst[2]).to.deep.equal({ si: [{ t: ["World"] }] });
    });

    it("count and uniqueCount match", () => {
      const ss = new SharedStrings();
      ss.register("X");
      ss.register("Y");
      ss.register("X"); // dup

      const tree = new Formatter().format(ss);
      const sst = tree["sst"] as any[];
      expect(sst[0]["_attr"]["count"]).toBe(2);
      expect(sst[0]["_attr"]["uniqueCount"]).toBe(2);
    });
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
