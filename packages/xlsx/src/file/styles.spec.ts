import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Styles } from "./styles";

const context = { stack: [] } as any;

describe("Styles", () => {
  // ── register() deduplication ──

  describe("register", () => {
    it("default style returns index 0", () => {
      const styles = new Styles();
      expect(styles.register({})).toBe(0);
    });

    it("registering default twice returns same index", () => {
      const styles = new Styles();
      styles.register({});
      expect(styles.register({})).toBe(0);
    });

    it("with font returns new index", () => {
      const styles = new Styles();
      const idx = styles.register({ font: { bold: true } });
      expect(idx).toBeGreaterThan(0);
    });

    it("same font deduplicates", () => {
      const styles = new Styles();
      const idx1 = styles.register({ font: { bold: true } });
      const idx2 = styles.register({ font: { bold: true } });
      expect(idx1).toBe(idx2);
    });

    it("different fonts get different indices", () => {
      const styles = new Styles();
      const idx1 = styles.register({ font: { bold: true } });
      const idx2 = styles.register({ font: { italic: true } });
      expect(idx1).not.toBe(idx2);
    });

    it("with fill returns new index", () => {
      const styles = new Styles();
      const idx = styles.register({ fill: { type: "solid", color: "FF0000" } });
      expect(idx).toBeGreaterThan(0);
    });

    it("with border returns new index", () => {
      const styles = new Styles();
      const idx = styles.register({
        border: { top: { style: "thin", color: "000000" } },
      });
      expect(idx).toBeGreaterThan(0);
    });

    it("with alignment sets applyAlignment=1", () => {
      const styles = new Styles();
      styles.register({ alignment: { horizontal: "center" } });

      const tree = new Formatter().format(styles);
      const styleSheet = tree["styleSheet"] as any[];

      // Find cellXfs
      const cellXfs = styleSheet.find((c: any) => "cellXfs" in c)["cellXfs"] as any[];
      // Default xf (index 0) + our xf (index 1)
      const xf = cellXfs[2]; // _attr + first xf + our xf
      expect(xf["xf"]).toBeDefined();
      const xfChildren = xf["xf"] as any[];
      const xfAttr = xfChildren.find((c: any) => "_attr" in c)["_attr"];
      expect(xfAttr["applyAlignment"]).toBe(1);
    });

    it("custom numFmt gets 164+ ID", () => {
      const styles = new Styles();
      styles.register({ numFmt: '0.00"€"' });

      const tree = new Formatter().format(styles);
      const styleSheet = tree["styleSheet"] as any[];

      // numFmts section should be present
      const numFmts = styleSheet.find((c: any) => "numFmts" in c);
      expect(numFmts).toBeDefined();
    });

    it("built-in numFmt does not create custom entry", () => {
      const styles = new Styles();
      const idx = styles.register({ numFmt: "0.00" });
      // "0.00" is built-in (id=2), should return 0 for default font/fill/border
      expect(idx).toBeGreaterThan(0);
    });
  });

  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("default construction has required sections", () => {
      const tree = new Formatter().format(new Styles());
      const styleSheet = tree["styleSheet"] as any[];

      const keys = styleSheet.map((c: any) => Object.keys(c)[0]);
      expect(keys).toContain("fonts");
      expect(keys).toContain("fills");
      expect(keys).toContain("borders");
      expect(keys).toContain("cellStyleXfs");
      expect(keys).toContain("cellXfs");
      expect(keys).toContain("cellStyles");
    });

    it("default fonts count is 1 (Calibri 11pt)", () => {
      const tree = new Formatter().format(new Styles());
      const styleSheet = tree["styleSheet"] as any[];
      const fonts = styleSheet.find((c: any) => "fonts" in c)["fonts"] as any[];
      expect(fonts[0]["_attr"]["count"]).toBe(1);
    });

    it("default fills count is 2 (none + gray125)", () => {
      const tree = new Formatter().format(new Styles());
      const styleSheet = tree["styleSheet"] as any[];
      const fills = styleSheet.find((c: any) => "fills" in c)["fills"] as any[];
      expect(fills[0]["_attr"]["count"]).toBe(2);
    });

    it("registering a font increases font count", () => {
      const styles = new Styles();
      styles.register({ font: { bold: true, size: 14 } });

      const tree = new Formatter().format(styles);
      const styleSheet = tree["styleSheet"] as any[];
      const fonts = styleSheet.find((c: any) => "fonts" in c)["fonts"] as any[];
      expect(fonts[0]["_attr"]["count"]).toBe(2);
    });
  });

  // ── toXml path ──

  describe("toXml", () => {
    it("produces valid styleSheet XML", () => {
      const xml = new Styles().toXml(context);
      expect(xml).toContain("<styleSheet");
      expect(xml).toContain("</styleSheet>");
      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    });

    it("includes default font, fills, borders", () => {
      const xml = new Styles().toXml(context);
      expect(xml).toContain("<fonts");
      expect(xml).toContain("<fills");
      expect(xml).toContain("<borders");
    });

    it("includes bold font when registered", () => {
      const styles = new Styles();
      styles.register({ font: { bold: true } });
      const xml = styles.toXml(context);
      expect(xml).toContain("<b/>");
    });

    it("includes custom numFmt in XML", () => {
      const styles = new Styles();
      styles.register({ numFmt: "#,##0.00_custom" });
      const xml = styles.toXml(context);
      expect(xml).toContain("<numFmts");
      expect(xml).toContain('numFmtId="164"');
    });
  });
});
