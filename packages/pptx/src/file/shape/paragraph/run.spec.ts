import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { TextRun } from "./run";

const context = { stack: [] } as any;

describe("TextRun", () => {
  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("empty constructor produces <a:r/>", () => {
      const tree = new Formatter().format(new TextRun());
      expect(tree).to.deep.equal({ "a:r": {} });
    });

    it("string constructor produces text content", () => {
      const tree = new Formatter().format(new TextRun("Hello"));
      expect(tree).to.deep.equal({ "a:r": [{ "a:t": ["Hello"] }] });
    });

    it("options with text produces text content", () => {
      const tree = new Formatter().format(new TextRun({ text: "World" }));
      expect(tree).to.deep.equal({ "a:r": [{ "a:t": ["World"] }] });
    });

    it("options with bold produces rPr + t", () => {
      const tree = new Formatter().format(new TextRun({ text: "Bold", bold: true }));
      expect(tree).to.deep.equal({
        "a:r": [{ "a:rPr": { _attr: { b: true } } }, { "a:t": ["Bold"] }],
      });
    });

    it("options with fontSize produces sz attribute (hundredths of pt)", () => {
      const tree = new Formatter().format(new TextRun({ text: "Big", fontSize: 24 }));
      expect(tree).to.deep.equal({
        "a:r": [{ "a:rPr": { _attr: { sz: 2400 } } }, { "a:t": ["Big"] }],
      });
    });

    it("options with multiple properties", () => {
      const tree = new Formatter().format(
        new TextRun({ text: "Styled", bold: true, italic: true, fontSize: 18 }),
      );
      expect(tree).to.deep.equal({
        "a:r": [{ "a:rPr": { _attr: { sz: 1800, b: true, i: true } } }, { "a:t": ["Styled"] }],
      });
    });

    it("properties without text produces rPr wrapped in a:r", () => {
      const tree = new Formatter().format(new TextRun({ bold: true }));
      expect(tree).to.deep.equal({
        "a:r": [{ "a:rPr": { _attr: { b: true } } }],
      });
    });
  });

  // ── toXml path ──

  describe("toXml", () => {
    it("empty constructor produces self-closing tag", () => {
      const xml = new TextRun().toXml(context);
      expect(xml).toBe("<a:r/>");
    });

    it("string constructor produces text element", () => {
      const xml = new TextRun("Hello").toXml(context);
      expect(xml).toBe("<a:r><a:t>Hello</a:t></a:r>");
    });

    it("options with bold produces rPr + t", () => {
      const xml = new TextRun({ text: "Bold", bold: true }).toXml(context);
      expect(xml).toContain('<a:rPr b="true"/>');
      expect(xml).toContain("<a:t>Bold</a:t>");
    });

    it("escapes XML special characters in text", () => {
      const xml = new TextRun("<b>&\"'</b>").toXml(context);
      expect(xml).toContain("<a:t>&lt;b&gt;&amp;&quot;&apos;&lt;/b&gt;</a:t>");
    });

    it("properties without text produces rPr wrapped in a:r", () => {
      const xml = new TextRun({ bold: true }).toXml(context);
      expect(xml).toContain('<a:rPr b="true"/>');
      expect(xml).toContain("<a:r>");
      expect(xml).toContain("</a:r>");
    });
  });
});
