import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { Styles, stylesDesc } from "./styles";
import type { StylesParseResult } from "./styles";

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
      const xml = styles.serialize();
      expect(xml).toContain('applyAlignment="1"');
    });

    it("custom numFmt gets 164+ ID", () => {
      const styles = new Styles();
      styles.register({ numFmt: '0.00"€"' });
      const xml = styles.serialize();
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

  describe("serialize", () => {
    it("produces valid styleSheet XML", () => {
      const xml = new Styles().serialize();
      expect(xml).toContain("<styleSheet");
      expect(xml).toContain("</styleSheet>");
      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    });

    it("includes default font, fills, borders", () => {
      const xml = new Styles().serialize();
      expect(xml).toContain("<fonts");
      expect(xml).toContain("<fills");
      expect(xml).toContain("<borders");
    });

    it("includes bold font when registered", () => {
      const styles = new Styles();
      styles.register({ font: { bold: true } });
      const xml = styles.serialize();
      expect(xml).toContain("<b/>");
    });

    it("includes custom numFmt in XML", () => {
      const styles = new Styles();
      styles.register({ numFmt: "#,##0.00_custom" });
      const xml = styles.serialize();
      expect(xml).toContain("<numFmts");
      expect(xml).toContain('numFmtId="164"');
    });
  });

  // ── DXFs (differential formats) ──

  describe("registerDxf", () => {
    it("returns incrementing indices starting at 0", () => {
      const styles = new Styles();
      const id1 = styles.registerDxf({ font: { bold: true } });
      const id2 = styles.registerDxf({ fill: { color: "FF0000" } });
      expect(id1).toBe(0);
      expect(id2).toBe(1);
    });

    it("outputs dxfs with count when registered", () => {
      const styles = new Styles();
      styles.registerDxf({ font: { bold: true, color: "FF0000" } });
      const xml = styles.serialize();
      expect(xml).toContain('<dxfs count="1">');
      expect(xml).toContain("<dxf>");
      expect(xml).toContain("<b/>");
    });

    it("outputs dxfs count=0 when none registered", () => {
      const xml = new Styles().serialize();
      expect(xml).toContain('dxfs count="0"');
    });

    it("includes fill bgColor for dxf", () => {
      const styles = new Styles();
      styles.registerDxf({ fill: { color: "00FF00" } });
      const xml = styles.serialize();
      expect(xml).toContain('bgColor rgb="FF00FF00"');
    });

    it("includes numFmt in dxf", () => {
      const styles = new Styles();
      styles.registerDxf({ numFmt: "0.00" });
      const xml = styles.serialize();
      expect(xml).toContain('formatCode="0.00"');
    });
  });

  // ── stylesDesc round-trip (stringify → parse) ──

  describe("stylesDesc round-trip", () => {
    function roundTripStyles(styles: Styles): StylesParseResult {
      const xml = stylesDesc.stringify({ styles }, undefined as unknown as WriteContext)!;
      const doc = parseXml(xml);
      const el = doc.elements![0];
      return stylesDesc.parse(
        el,
        undefined as unknown as ReadContext,
      ) as unknown as StylesParseResult;
    }

    it("round-trips alignment with relativeIndent/shrinkToFit/readingOrder/justifyLastLine", () => {
      const styles = new Styles();
      styles.register({
        alignment: {
          horizontal: "center",
          relativeIndent: 3,
          shrinkToFit: true,
          readingOrder: 2,
          justifyLastLine: true,
        },
      });
      const result = roundTripStyles(styles);
      const xf = result.cellXfs!.find((x) => x.alignment);
      expect(xf?.alignment?.horizontal).toBe("center");
      expect(xf?.alignment?.relativeIndent).toBe(3);
      expect(xf?.alignment?.shrinkToFit).toBe(true);
      expect(xf?.alignment?.readingOrder).toBe(2);
      expect(xf?.alignment?.justifyLastLine).toBe(true);
    });

    it("round-trips gradient fill with left/right/top/bottom positions", () => {
      const styles = new Styles();
      styles.register({
        fill: {
          type: "gradient",
          gradientType: "path",
          gradientLeft: 10,
          gradientRight: 90,
          gradientTop: 20,
          gradientBottom: 80,
          stops: [
            { position: 0, color: "FF0000" },
            { position: 1, color: "0000FF" },
          ],
        },
      });
      const result = roundTripStyles(styles);
      const grad = result.fills!.find((f) => f.type === "gradient");
      expect(grad).toBeDefined();
      expect(grad!.gradientType).toBe("path");
      expect(grad!.gradientLeft).toBe(10);
      expect(grad!.gradientRight).toBe(90);
      expect(grad!.gradientTop).toBe(20);
      expect(grad!.gradientBottom).toBe(80);
      expect(grad!.stops).toHaveLength(2);
      expect(grad!.stops![0].color).toBe("FF0000");
      expect(grad!.stops![1].color).toBe("0000FF");
    });
  });
});
