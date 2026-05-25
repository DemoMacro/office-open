import { Formatter } from "@export/formatter";
import { toElement } from "@office-open/xml";
import { beforeEach, describe, expect, it } from "vite-plus/test";

import { Paragraph } from "../paragraph";
import { ColumnBreak, PageBreak, PageBreakBefore } from "./break";

describe("PageBreak", () => {
  let pageBreak: PageBreak;

  beforeEach(() => {
    pageBreak = new PageBreak();
  });

  describe("#constructor()", () => {
    it("should create a Page Break with correct attributes", () => {
      const tree = new Formatter().format(pageBreak);
      expect(tree).to.deep.equal({
        "w:r": [
          {
            "w:br": {
              _attr: {
                "w:type": "page",
              },
            },
          },
        ],
      });
    });
  });
});

describe("ColumnBreak", () => {
  let columnBreak: ColumnBreak;

  beforeEach(() => {
    columnBreak = new ColumnBreak();
  });

  describe("#constructor()", () => {
    it("should create a Column Break with correct attributes", () => {
      const tree = new Formatter().format(columnBreak);
      expect(tree).to.deep.equal({
        "w:r": [
          {
            "w:br": {
              _attr: {
                "w:type": "column",
              },
            },
          },
        ],
      });
    });
  });
});

describe("PageBreakBefore", () => {
  it("should create page break before", () => {
    const pageBreakBefore = new PageBreakBefore();
    const tree = new Formatter().format(pageBreakBefore);
    expect(tree).to.deep.equal({
      "w:pageBreakBefore": {},
    });
  });
});

// ── Parse round-trip tests ──────────────────────────────────────────────────

describe("Break parse round-trip", () => {
  it("should round-trip pageBreak via Paragraph JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "Before" }, { pageBreak: true }, { text: "After" }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    const brRun = root.elements?.find((e) =>
      e.elements?.some((c) => c.name === "w:br" && c.attributes?.["w:type"] === "page"),
    );
    expect(brRun).toBeDefined();
  });

  it("should round-trip columnBreak via Paragraph JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ text: "Col1" }, { columnBreak: true }, { text: "Col2" }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    const brRun = root.elements?.find((e) =>
      e.elements?.some((c) => c.name === "w:br" && c.attributes?.["w:type"] === "column"),
    );
    expect(brRun).toBeDefined();
  });
});
