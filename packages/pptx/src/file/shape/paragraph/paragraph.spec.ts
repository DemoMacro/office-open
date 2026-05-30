import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "./paragraph";
import { TextRun } from "./run";

const context = { stack: [] } as any;

describe("Paragraph", () => {
  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("empty constructor produces <a:p/> with endParaRPr", () => {
      const tree = new Formatter().format(new Paragraph());
      expect(tree).to.deep.equal({
        "a:p": [{ "a:endParaRPr": { _attr: { lang: "en-US" } } }],
      });
    });

    it("string constructor produces text run + endParaRPr", () => {
      const tree = new Formatter().format(new Paragraph("Hello"));
      expect(tree).to.deep.equal({
        "a:p": [
          { "a:r": [{ "a:t": ["Hello"] }] },
          { "a:endParaRPr": { _attr: { lang: "en-US" } } },
        ],
      });
    });

    it("children with TextRun instances", () => {
      const tree = new Formatter().format(
        new Paragraph({ children: [new TextRun("First"), new TextRun("Second")] }),
      );
      expect(tree).to.deep.equal({
        "a:p": [
          { "a:r": [{ "a:t": ["First"] }] },
          { "a:r": [{ "a:t": ["Second"] }] },
          { "a:endParaRPr": { _attr: { lang: "en-US" } } },
        ],
      });
    });

    it("children with RunOptions objects (auto-wrapped to TextRun)", () => {
      const tree = new Formatter().format(new Paragraph({ children: [{ text: "Auto" }] }));
      expect(tree).to.deep.equal({
        "a:p": [{ "a:r": [{ "a:t": ["Auto"] }] }, { "a:endParaRPr": { _attr: { lang: "en-US" } } }],
      });
    });

    it("with alignment property", () => {
      const tree = new Formatter().format(
        new Paragraph({ text: "Center", properties: { alignment: "center" } }),
      );
      expect(tree).to.deep.equal({
        "a:p": [
          { "a:pPr": { _attr: { algn: "ctr" } } },
          { "a:r": [{ "a:t": ["Center"] }] },
          { "a:endParaRPr": { _attr: { lang: "en-US" } } },
        ],
      });
    });

    it("with indentLevel property", () => {
      const tree = new Formatter().format(
        new Paragraph({ text: "Indent", properties: { indentLevel: 2 } }),
      );
      expect(tree).to.deep.equal({
        "a:p": [
          { "a:pPr": { _attr: { lvl: 2 } } },
          { "a:r": [{ "a:t": ["Indent"] }] },
          { "a:endParaRPr": { _attr: { lang: "en-US" } } },
        ],
      });
    });

    it("always appends endParaRPr at the end", () => {
      const tree = new Formatter().format(new Paragraph("Test"));
      const pChildren = tree["a:p"] as any[];
      const last = pChildren[pChildren.length - 1];
      expect(last).to.deep.equal({ "a:endParaRPr": { _attr: { lang: "en-US" } } });
    });
  });

  // ── toXml path ──

  describe("toXml", () => {
    it("empty constructor produces <a:p/>", () => {
      const xml = new Paragraph().toXml(context);
      expect(xml).toBe('<a:p><a:endParaRPr lang="en-US"/></a:p>');
    });

    it("string constructor produces text + endParaRPr", () => {
      const xml = new Paragraph("Hello").toXml(context);
      expect(xml).toContain("<a:t>Hello</a:t>");
      expect(xml).toContain('<a:endParaRPr lang="en-US"/>');
    });

    it("with alignment produces pPr", () => {
      const xml = new Paragraph({
        text: "Center",
        properties: { alignment: "center" },
      }).toXml(context);
      expect(xml).toContain('<a:pPr algn="ctr"/>');
    });

    it("children produce multiple runs", () => {
      const xml = new Paragraph({
        children: [new TextRun("A"), new TextRun("B")],
      }).toXml(context);
      expect(xml).toContain("<a:t>A</a:t>");
      expect(xml).toContain("<a:t>B</a:t>");
    });
  });
});
