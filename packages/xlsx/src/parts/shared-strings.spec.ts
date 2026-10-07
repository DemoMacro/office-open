import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { buildRPrXml, parseRPr, SharedStrings, sharedStringsDesc } from "./shared-strings";

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

  it("parses the legacy Excel 2005 SST compatibility vocabulary", () => {
    const doc = parseXml(
      '<sst xmlns="http://schemas.microsoft.com/office/excel/2005/8/sst" ' +
        'totalCount="3" uniqueCount="2"><sstItem><t>A</t></sstItem>' +
        "<sstItem><t>B</t></sstItem></sst>",
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const result = sharedStringsDesc.parse(root, {} as never);
    expect(result.entries).toEqual(["A", "B"]);
    expect(sharedStringsDesc.stringify(result, {} as never)).toContain("<si><t>A</t></si>");
  });

  it("round-trips legacy shared-string compatibility content", () => {
    const doc = parseXml(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<si xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
        'mc:Ignorable="w14" w14:attr="value" ' +
        'xmlns:w14="http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing">' +
        '<w14:placeholder mc:ProcessContent="w14:placeholder" mc:PreserveAttributes="w14:a w14:b">' +
        '<t w14:a="a" w14:b="b" w14:c="c">wrapped</t></w14:placeholder><w14:no/>' +
        "<t>value</t></si></sst>",
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const result = sharedStringsDesc.parse(root, {} as never);
    const entry = result.entries[0];
    if (typeof entry === "string") throw new Error("expected a rich-text entry");
    if (!entry) throw new Error("parsed shared string has no entry");
    expect(entry.wordDrawingExtension).toEqual({
      attribute: "value",
      placeholder: {
        processContent: "w14:placeholder",
        preserveAttributes: "w14:a w14:b",
        text: "wrapped",
        textAttributes: { a: "a", b: "b", c: "c" },
      },
      no: true,
    });

    const xml = sharedStringsDesc.stringify(result, {} as never)!;
    expect(xml).toContain('w14:attr="value"');
    expect(xml).toContain('mc:PreserveAttributes="w14:a w14:b"');
    expect(xml).toContain('<t w14:a="a" w14:b="b" w14:c="c">wrapped</t>');
    expect(xml).toContain("<w14:no/>");
    expect(xml.indexOf("<w14:placeholder")).toBeLessThan(xml.indexOf("<t>value</t>"));
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

    it("round-trips CT_RPrElt source child order", () => {
      const el = parseXml(
        '<rPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<rFont val="A"/><family val="2"/><charset val="1"/></rPr>',
      ).elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      const properties = parseRPr(el);
      expect(properties.propertyOrder).toEqual(["rFont", "family", "charset"]);
      expect(buildRPrXml(properties)).toContain(
        '<rFont val="A"/><family val="2"/><charset val="1"/>',
      );
    });
  });
});
