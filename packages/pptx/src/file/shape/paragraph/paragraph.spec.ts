import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "./paragraph";
import { TextRun } from "./run";

const context = { stack: [] } as any;

describe("Paragraph", () => {
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
