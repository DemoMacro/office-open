import type {
  GroupShapeOptions as PptxGroupOptions,
  ShapeOptions as PptxShapeOptions,
} from "@office-open/pptx";
import { describe, expect, it, vi } from "vitest";

import { toDocxGroup, toPptxGroup, toXlsxGroup } from "./group";

const PPTX_GROUP: PptxGroupOptions = {
  x: 609600,
  y: 190500,
  width: 2438400,
  height: 762000,
  children: [
    {
      shape: {
        x: 609600,
        y: 190500,
        width: 1219200,
        height: 381000,
        geometry: "rect",
        fill: "4472C4",
        textBody: { text: "Child" },
      },
    },
  ],
};

describe("toXlsxGroup (pptx → xlsx)", () => {
  it("maps container to grpSpPr + anchor and recurses into shape children", () => {
    const x = toXlsxGroup(PPTX_GROUP);
    expect(x.col).toBe(2);
    expect(x.row).toBe(2);
    expect(x.grpSpPr.width).toBe(2438400);
    expect(x.grpSpPr.height).toBe(762000);
    expect(x.shapes).toHaveLength(1);
    expect(x.shapes?.[0].spPr.geometry).toBe("rect");
    expect(x.shapes?.[0].spPr.fill).toBe("4472C4");
    expect(x.shapes?.[0].textBody).toEqual({ text: "Child" });
  });

  it("warns and drops children xlsx groups cannot host", () => {
    const spy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const x = toXlsxGroup({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      children: [
        { picture: { type: "png", data: new Uint8Array(), x: 0, y: 0, width: 1, height: 1 } },
      ],
    });
    expect(x.shapes).toBeUndefined();
    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });
});

describe("toDocxGroup (pptx → docx)", () => {
  it("emits wpg children as wps ShapeMediaData", () => {
    const d = toDocxGroup(PPTX_GROUP);
    expect(d.transformation.offset).toEqual({ left: 609600, top: 190500 });
    expect(d.transformation.width).toBe(2438400);
    expect(d.children).toHaveLength(1);
    expect(d.children[0].type).toBe("wps");
  });
});

describe("round-trip pptx → xlsx → pptx", () => {
  it("preserves container + shape child", () => {
    const back = toPptxGroup(toXlsxGroup(PPTX_GROUP));
    expect(back.x).toBe(609600);
    expect(back.y).toBe(190500);
    expect(back.width).toBe(2438400);
    expect(back.height).toBe(762000);
    expect(back.children).toHaveLength(1);
    const shape = (back.children[0] as { shape: PptxShapeOptions }).shape;
    expect(shape.geometry).toBe("rect");
    expect(shape.fill).toBe("4472C4");
    expect(shape.textBody).toEqual({ text: "Child" });
  });
});

describe("round-trip pptx → docx → pptx", () => {
  it("preserves container + shape child (text normalizes to paragraphs)", () => {
    const back = toPptxGroup(toDocxGroup(PPTX_GROUP));
    expect(back.x).toBe(609600);
    expect(back.y).toBe(190500);
    expect(back.width).toBe(2438400);
    expect(back.height).toBe(762000);
    expect(back.children).toHaveLength(1);
    const shape = (back.children[0] as { shape: PptxShapeOptions }).shape;
    expect(shape.fill).toBe("4472C4");
    expect(shape.textBody?.paragraphs?.[0]).toBe("Child");
  });
});
