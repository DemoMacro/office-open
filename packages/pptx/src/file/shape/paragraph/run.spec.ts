import { describe, expect, it } from "vite-plus/test";

import { TextRun } from "./run";

const context = { stack: [] };

describe("TextRun", () => {
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
      expect(xml).toContain('<a:rPr b="1"/>');
      expect(xml).toContain("<a:t>Bold</a:t>");
    });

    it("escapes XML special characters in text", () => {
      const xml = new TextRun("<b>&\"'</b>").toXml(context);
      expect(xml).toContain("<a:t>&lt;b&gt;&amp;&quot;&apos;&lt;/b&gt;</a:t>");
    });

    it("properties without text produces rPr wrapped in a:r", () => {
      const xml = new TextRun({ bold: true }).toXml(context);
      expect(xml).toContain('<a:rPr b="1"/>');
      expect(xml).toContain("<a:r>");
      expect(xml).toContain("</a:r>");
    });
  });
});
