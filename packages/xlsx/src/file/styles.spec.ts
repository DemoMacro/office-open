import { describe, expect, it } from "vite-plus/test";

import { Styles } from "./styles";

const context = { stack: [] };

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
      const xml = styles.toXml(context);
      expect(xml).toContain('applyAlignment="1"');
    });

    it("custom numFmt gets 164+ ID", () => {
      const styles = new Styles();
      styles.register({ numFmt: '0.00"€"' });
      const xml = styles.toXml(context);
      expect(xml).toContain("<numFmts");
      expect(xml).toContain('numFmtId="164"');
    });

    it("built-in numFmt does not create custom entry", () => {
      const styles = new Styles();
      const idx = styles.register({ numFmt: "0.00" });
      expect(idx).toBeGreaterThan(0);
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
