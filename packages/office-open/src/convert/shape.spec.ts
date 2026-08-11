import type { ShapeOptions as PptxShapeOptions } from "@office-open/pptx";
import type { ShapeOptions as XlsxShapeOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vitest";

import { toDocxShape, toPptxShape, toXlsxShape } from "./shape";

// Coordinates picked as integer multiples of the heuristic cell sizes
// (DEFAULT_COL_EMU = 609600, DEFAULT_ROW_EMU = 190500) so xlsx ↔ {pptx,docx}
// round-trips position losslessly.
const RECT_PPTX: PptxShapeOptions = {
  x: 609600,
  y: 190500,
  width: 1219200,
  height: 381000,
  geometry: "rect",
  fill: "4472C4",
  outline: { type: "solidFill", color: { value: "ED7D31" }, width: 9525 },
  textBody: { text: "Hello" },
  name: "My Shape",
};

describe("toXlsxShape (pptx → xlsx)", () => {
  it("maps absolute EMU origin to cell anchor + spPr.xfrm", () => {
    const x = toXlsxShape(RECT_PPTX);
    expect(x.col).toBe(2);
    expect(x.row).toBe(2);
    expect(x.toCol).toBe(4);
    expect(x.toRow).toBe(4);
    expect(x.spPr.width).toBe(1219200);
    expect(x.spPr.height).toBe(381000);
    expect(x.spPr.geometry).toBe("rect");
    expect(x.spPr.fill).toBe("4472C4");
    expect(x.name).toBe("My Shape");
  });

  it("carries the outline verbatim onto spPr", () => {
    const x = toXlsxShape(RECT_PPTX);
    expect(x.spPr.outline).toEqual({ type: "solidFill", color: { value: "ED7D31" }, width: 9525 });
  });

  it("passes the text body through (a:p shared by pptx/xlsx)", () => {
    const x = toXlsxShape(RECT_PPTX);
    expect(x.textBody).toEqual({ text: "Hello" });
  });
});

describe("toPptxShape (xlsx → pptx)", () => {
  it("reconstructs top-level x/y/w/h from anchor + spPr", () => {
    const xlsx: XlsxShapeOptions = {
      col: 2,
      row: 2,
      toCol: 4,
      toRow: 4,
      spPr: {
        x: 0,
        y: 0,
        width: 1219200,
        height: 381000,
        geometry: "rect",
        fill: "4472C4",
      },
    };
    const p = toPptxShape(xlsx);
    expect(p.x).toBe(609600);
    expect(p.y).toBe(190500);
    expect(p.width).toBe(1219200);
    expect(p.height).toBe(381000);
    expect(p.geometry).toBe("rect");
    expect(p.fill).toBe("4472C4");
  });
});

describe("toDocxShape (pptx → docx)", () => {
  it("maps position to transformation and textBody to w:p children", () => {
    const d = toDocxShape(RECT_PPTX);
    expect(d.transformation.offset).toEqual({ left: 609600, top: 190500 });
    expect(d.transformation.width).toBe(1219200);
    expect(d.transformation.height).toBe(381000);
    // text shorthand → one string paragraph (docx children accept strings)
    expect(d.children).toHaveLength(1);
    expect(d.children[0]).toBe("Hello");
    // docx geometry rejects the bare-string shorthand → preset object
    expect(d.presetGeometry).toEqual({ preset: "rect" });
    expect(d.fill).toBe("4472C4");
  });
});

describe("toPptxShape (docx → pptx)", () => {
  it("reconstructs x/y/w/h from transformation and textBody from children", () => {
    const d = toDocxShape(RECT_PPTX);
    const back = toPptxShape(d);
    expect(back.x).toBe(609600);
    expect(back.y).toBe(190500);
    expect(back.width).toBe(1219200);
    expect(back.height).toBe(381000);
    expect(back.geometry).toEqual({ preset: "rect" });
    expect(back.fill).toBe("4472C4");
    expect(back.textBody?.paragraphs?.[0]).toBe("Hello");
  });
});

describe("round-trip pptx → xlsx → pptx", () => {
  it("preserves position (integer cell multiples) + content", () => {
    const back = toPptxShape(toXlsxShape(RECT_PPTX));
    expect(back.x).toBe(609600);
    expect(back.y).toBe(190500);
    expect(back.width).toBe(1219200);
    expect(back.height).toBe(381000);
    expect(back.geometry).toBe("rect");
    expect(back.fill).toBe("4472C4");
    expect(back.outline).toEqual({ type: "solidFill", color: { value: "ED7D31" }, width: 9525 });
    expect(back.textBody).toEqual({ text: "Hello" });
    expect(back.name).toBe("My Shape");
  });
});

describe("round-trip pptx → docx → pptx", () => {
  it("preserves position + fill + text (geometry normalizes to preset object)", () => {
    const back = toPptxShape(toDocxShape(RECT_PPTX));
    expect(back.x).toBe(609600);
    expect(back.y).toBe(190500);
    expect(back.width).toBe(1219200);
    expect(back.height).toBe(381000);
    expect(back.fill).toBe("4472C4");
    expect(back.textBody?.paragraphs?.[0]).toBe("Hello");
    // geometry string → docx preset → pptx preset object (lossless value, different shape)
    expect(back.geometry).toEqual({ preset: "rect" });
  });
});
