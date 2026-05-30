import { beforeEach, describe, expect, it } from "vite-plus/test";

import { Shape } from "./shape";

const context = { stack: [] } as any;

// Shape uses a static nextId counter, so we reset it between tests
beforeEach(() => {
  (Shape as any).nextId = 2;
});

describe("Shape", () => {
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
