import { Formatter } from "@export/formatter";
import { beforeEach, describe, expect, it } from "vite-plus/test";

import { Shape } from "./shape";

const context = { stack: [] } as any;

// Shape uses a static nextId counter, so we reset it between tests
beforeEach(() => {
  (Shape as any).nextId = 2;
});

describe("Shape", () => {
  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("default constructor produces p:sp with auto-generated id", () => {
      const tree = new Formatter().format(new Shape());
      const sp = tree["p:sp"] as any[];
      // p:nvSpPr, p:spPr, p:txBody
      expect(sp).toHaveLength(3);

      // nvSpPr contains cNvPr with id=2, name="Shape 2"
      const nvSpPr = sp[0]["p:nvSpPr"];
      expect(nvSpPr[0]["p:cNvPr"]["_attr"]).to.deep.equal({ id: 2, name: "Shape 2" });
    });

    it("custom id and name", () => {
      const tree = new Formatter().format(new Shape({ id: 10, name: "MyShape" }));
      const sp = tree["p:sp"] as any[];
      const nvSpPr = sp[0]["p:nvSpPr"];
      expect(nvSpPr[0]["p:cNvPr"]["_attr"]).to.deep.equal({ id: 10, name: "MyShape" });
    });

    it("auto-increments id when not provided", () => {
      new Shape(); // id=2
      const shape2 = new Shape(); // id=3
      expect(shape2.shapeId).toBe(3);
    });

    it("with placeholder produces p:ph element", () => {
      const tree = new Formatter().format(new Shape({ placeholder: "title" }));
      const sp = tree["p:sp"] as any[];
      const nvSpPr = sp[0]["p:nvSpPr"];
      const nvPr = nvSpPr[2]["p:nvPr"];
      expect(nvPr).to.deep.equal([{ "p:ph": { _attr: { type: "title" } } }]);
    });

    it("with placeholder and index produces p:ph with idx", () => {
      const tree = new Formatter().format(new Shape({ placeholder: "body", placeholderIndex: 1 }));
      const sp = tree["p:sp"] as any[];
      const nvSpPr = sp[0]["p:nvSpPr"];
      const nvPr = nvSpPr[2]["p:nvPr"];
      expect(nvPr).to.deep.equal([{ "p:ph": { _attr: { type: "body", idx: 1 } } }]);
    });

    it("with textBody produces txBody with paragraphs", () => {
      const tree = new Formatter().format(new Shape({ textBody: { text: "Hello" } }));
      const sp = tree["p:sp"] as any[];
      // Last child should be txBody
      const txBody = sp[sp.length - 1]["p:txBody"];
      expect(txBody).toBeDefined();
    });

    it("shapeId is a public readonly property", () => {
      const shape = new Shape({ id: 42 });
      expect(shape.shapeId).toBe(42);
    });

    it("animation property is stored but not in XML", () => {
      const shape = new Shape({ animation: { triggers: [] } as any });
      expect(shape.animation).toBeDefined();
    });
  });

  // ── toXml path ──

  describe("toXml", () => {
    it("default constructor produces valid XML", () => {
      const xml = new Shape().toXml(context);
      expect(xml).toContain("<p:sp>");
      expect(xml).toContain("</p:sp>");
      expect(xml).toContain('name="Shape 2"');
    });

    it("custom id and name in XML", () => {
      const xml = new Shape({ id: 5, name: "Custom" }).toXml(context);
      expect(xml).toContain('id="5"');
      expect(xml).toContain('name="Custom"');
    });

    it("placeholder in XML produces p:ph", () => {
      const xml = new Shape({ placeholder: "title" }).toXml(context);
      expect(xml).toContain('<p:ph type="title"/>');
    });

    it("textBody in XML produces p:txBody with text", () => {
      const xml = new Shape({ textBody: { text: "Hi" } }).toXml(context);
      expect(xml).toContain("<a:t>Hi</a:t>");
    });
  });
});
