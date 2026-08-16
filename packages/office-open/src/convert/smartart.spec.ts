import type { SmartArtOptions as DocxSmartArt } from "@office-open/docx";
import type { SmartArtOptions as PptxSmartArt } from "@office-open/pptx";
import { describe, expect, it } from "vitest";

import { toDocxSmartArt, toPptxSmartArt } from "./smartart";

describe("toDocxSmartArt (pptx → docx)", () => {
  it("maps absolute position to offset and passes nodes through", () => {
    const pptx: PptxSmartArt = {
      nodes: [{ text: "Root", children: [{ text: "Child" }] }],
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      layout: "process1",
      style: "moderate1",
      color: "colorful1",
    };
    const docx = toDocxSmartArt(pptx);
    expect(docx.nodes).toEqual([{ text: "Root", children: [{ text: "Child" }] }]);
    expect(docx.transformation.width).toBe(300);
    expect(docx.transformation.height).toBe(400);
    expect(docx.transformation.offset).toEqual({ left: 100, top: 200 });
    expect(docx.layout).toBe("process1");
    expect(docx.style).toBe("moderate1");
    expect(docx.color).toBe("colorful1");
  });

  it("omits offset when pptx has no x/y", () => {
    const pptx: PptxSmartArt = { nodes: [], width: 10, height: 20 };
    expect(toDocxSmartArt(pptx).transformation.offset).toBeUndefined();
  });
});

describe("toPptxSmartArt (docx → pptx)", () => {
  it("maps transformation to absolute x/y and rebuilds nodes", () => {
    const docx: DocxSmartArt = {
      nodes: [{ text: "A", children: [{ text: "B" }] }],
      transformation: { width: 50, height: 60, offset: { left: 10, top: 20 } },
      layout: "hierarchy1",
    };
    const pptx = toPptxSmartArt(docx);
    expect(pptx.nodes).toEqual([{ text: "A", children: [{ text: "B" }] }]);
    expect(pptx.x).toBe(10);
    expect(pptx.y).toBe(20);
    expect(pptx.width).toBe(50);
    expect(pptx.height).toBe(60);
    expect(pptx.layout).toBe("hierarchy1");
  });
});

describe("round-trip pptx → docx → pptx", () => {
  it("preserves nodes, dimensions, and layout/style/color", () => {
    const pptx: PptxSmartArt = {
      nodes: [{ text: "Root", children: [{ text: "C1" }, { text: "C2" }] }],
      x: 50,
      y: 60,
      width: 700,
      height: 800,
      layout: "process1",
      style: "simple1",
      color: "accent1_2",
    };
    const back = toPptxSmartArt(toDocxSmartArt(pptx));
    expect(back.nodes).toEqual(pptx.nodes);
    expect(back.width).toBe(700);
    expect(back.height).toBe(800);
    expect(back.x).toBe(50);
    expect(back.y).toBe(60);
    expect(back.layout).toBe("process1");
    expect(back.style).toBe("simple1");
    expect(back.color).toBe("accent1_2");
  });
});

describe("cross-format cNvPr (alt text) preservation", () => {
  const name = "Flow";
  const description = "Process flow";
  const title = "Flow title";
  const hidden = true;

  const pptxSrc: PptxSmartArt = {
    nodes: [],
    width: 10,
    height: 10,
    name,
    description,
    title,
    hidden,
  };
  const docxSrc: DocxSmartArt = {
    nodes: [],
    transformation: { width: 10, height: 10 },
    altText: { name, description, title, hidden },
  };

  it("pptx → docx: cNvPr lands on altText", () => {
    expect(toDocxSmartArt(pptxSrc).altText).toEqual({ name, description, title, hidden });
  });

  it("docx → pptx: altText flows to top-level cNvPr", () => {
    const pptx = toPptxSmartArt(docxSrc);
    expect(pptx.name).toBe(name);
    expect(pptx.description).toBe(description);
    expect(pptx.title).toBe(title);
    expect(pptx.hidden).toBe(hidden);
  });

  it("docx defaults altText.name to SmartArt when source has other cNvPr but no name", () => {
    const docx = toDocxSmartArt({ ...pptxSrc, name: undefined });
    expect(docx.altText?.name).toBe("SmartArt");
    expect(docx.altText?.description).toBe(description);
  });

  it("omits docx altText when source carries no cNvPr at all", () => {
    const pptx: PptxSmartArt = { nodes: [], width: 10, height: 10 };
    expect(toDocxSmartArt(pptx).altText).toBeUndefined();
  });
});
