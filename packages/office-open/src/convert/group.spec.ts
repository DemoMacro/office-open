import type { GroupOptions as DocxGroupOptions, MediaTransformation } from "@office-open/docx";
import type {
  GroupOptions as PptxGroupOptions,
  ShapeOptions as PptxShapeOptions,
} from "@office-open/pptx";
import type { GroupOptions as XlsxGroupOptions } from "@office-open/xlsx";
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
        properties: { geometry: "rect", fill: "4472C4" },
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
    expect(x.properties.width).toBe(2438400);
    expect(x.properties.height).toBe(762000);
    expect(x.shapes).toHaveLength(1);
    expect(x.shapes?.[0].properties.geometry).toBe("rect");
    expect(x.shapes?.[0].properties.fill).toBe("4472C4");
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

  it("carries container rotation and vertical flip onto grpSpPr", () => {
    const x = toXlsxGroup({ ...PPTX_GROUP, rotation: 45, flipVertical: true });
    expect(x.properties.rotation).toBe(45);
    expect(x.properties.flipVertical).toBe(true);
  });
});

describe("toDocxGroup (pptx → docx)", () => {
  it("emits wpg children as wps ShapeMediaData", () => {
    const d = toDocxGroup(PPTX_GROUP);
    const transformation = d.transformation as MediaTransformation;
    expect(transformation.offset).toEqual({ x: 609600, y: 190500 });
    expect(transformation.width).toBe(2438400);
    expect(transformation.height).toBe(762000);
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
    expect(shape.properties?.geometry).toBe("rect");
    expect(shape.properties?.fill).toBe("4472C4");
    expect(shape.textBody).toEqual({ text: "Child" });
  });

  it("preserves a different child coordinate system", () => {
    const scaled = {
      ...PPTX_GROUP,
      childOffsetX: 0,
      childOffsetY: 0,
      childExtentWidth: 1219200,
      childExtentHeight: 381000,
    };
    const xlsx = toXlsxGroup(scaled);
    expect(xlsx.properties.childOffsetX).toBe(0);
    expect(xlsx.properties.childOffsetY).toBe(0);
    expect(xlsx.properties.childExtentWidth).toBe(1219200);
    expect(xlsx.properties.childExtentHeight).toBe(381000);
    const docx = toDocxGroup(scaled);
    expect(docx.childOffsetX).toBe(0);
    expect(docx.childOffsetY).toBe(0);
    expect(docx.childExtentWidth).toBe(1219200);
    expect(docx.childExtentHeight).toBe(381000);
    const back = toPptxGroup(xlsx);
    expect(back.childOffsetX).toBe(0);
    expect(back.childOffsetY).toBe(0);
    expect(back.childExtentWidth).toBe(1219200);
    expect(back.childExtentHeight).toBe(381000);
  });

  it("preserves container rotation and flips", () => {
    const source: PptxGroupOptions = {
      ...PPTX_GROUP,
      rotation: 45,
      flipHorizontal: true,
      flipVertical: true,
    };
    expect(toPptxGroup(toXlsxGroup(source))).toMatchObject({
      rotation: 45,
      flipHorizontal: true,
      flipVertical: true,
    });
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
    expect(shape.properties?.fill).toBe("4472C4");
    expect(shape.textBody?.paragraphs?.[0]).toBe("Child");
  });
});

describe("cross-format container cNvPr preservation", () => {
  const name = "Diagram group";
  const description = "Flow diagram";
  const title = "Group title";
  const hidden = true;

  it("pptx → xlsx: cNvPr passes straight through", () => {
    const x = toXlsxGroup({ ...PPTX_GROUP, name, description, title, hidden });
    expect(x.name).toBe(name);
    expect(x.description).toBe(description);
    expect(x.title).toBe(title);
    expect(x.hidden).toBe(hidden);
  });

  it("pptx → docx: cNvPr bridges to altText", () => {
    const d = toDocxGroup({ ...PPTX_GROUP, name, description, title, hidden });
    expect(d.altText?.name).toBe(name);
    expect(d.altText?.description).toBe(description);
    expect(d.altText?.title).toBe(title);
    expect(d.altText?.hidden).toBe(hidden);
  });

  it("xlsx → pptx: cNvPr passes straight through", () => {
    const x = toXlsxGroup(PPTX_GROUP);
    const back = toPptxGroup({ ...(x as XlsxGroupOptions), name, description, title, hidden });
    expect(back.name).toBe(name);
    expect(back.description).toBe(description);
    expect(back.title).toBe(title);
    expect(back.hidden).toBe(hidden);
  });

  it("docx → pptx: altText bridges back to cNvPr", () => {
    const d = toDocxGroup(PPTX_GROUP);
    const back = toPptxGroup({
      ...(d as DocxGroupOptions),
      altText: { name, description, title, hidden },
    });
    expect(back.name).toBe(name);
    expect(back.description).toBe(description);
    expect(back.title).toBe(title);
    expect(back.hidden).toBe(hidden);
  });

  it("docx → xlsx: altText bridges to cNvPr", () => {
    const d = toDocxGroup(PPTX_GROUP);
    const x = toXlsxGroup({
      ...(d as DocxGroupOptions),
      altText: { name, description, title, hidden },
    });
    expect(x.name).toBe(name);
    expect(x.description).toBe(description);
    expect(x.title).toBe(title);
    expect(x.hidden).toBe(hidden);
  });

  it("omits altText when no cNvPr field is authored", () => {
    const d = toDocxGroup(PPTX_GROUP);
    expect(d.altText).toBeUndefined();
  });
});

describe("cross-format child cNvPr preservation", () => {
  const childName = "Box";
  const childDescription = "Process box";
  const childTitle = "Box title";
  const childHidden = true;

  it("pptx → xlsx: child shape cNvPr carries all fields (not just name)", () => {
    const x = toXlsxGroup({
      ...PPTX_GROUP,
      children: [
        {
          shape: {
            ...(PPTX_GROUP.children[0] as { shape: PptxShapeOptions }).shape,
            name: childName,
            description: childDescription,
            title: childTitle,
            hidden: childHidden,
          },
        },
      ],
    });
    expect(x.shapes?.[0].name).toBe(childName);
    expect(x.shapes?.[0].description).toBe(childDescription);
    expect(x.shapes?.[0].title).toBe(childTitle);
    expect(x.shapes?.[0].hidden).toBe(childHidden);
  });

  it("pptx → docx: child shape cNvPr → nonVisualProperties (all fields)", () => {
    const d = toDocxGroup({
      ...PPTX_GROUP,
      children: [
        {
          shape: {
            ...(PPTX_GROUP.children[0] as { shape: PptxShapeOptions }).shape,
            name: childName,
            description: childDescription,
          },
        },
      ],
    });
    const nv = (
      d.children[0] as { data: { nonVisualProperties?: { name?: string; description?: string } } }
    ).data.nonVisualProperties;
    expect(nv?.name).toBe(childName);
    expect(nv?.description).toBe(childDescription);
  });

  it("xlsx → pptx: child shape cNvPr carries all fields", () => {
    const x = toXlsxGroup(PPTX_GROUP);
    const back = toPptxGroup({
      ...(x as XlsxGroupOptions),
      shapes: [
        {
          properties: x.shapes![0].properties,
          name: childName,
          description: childDescription,
          title: childTitle,
          hidden: childHidden,
        },
      ],
    });
    const shape = (back.children[0] as { shape: PptxShapeOptions }).shape;
    expect(shape.name).toBe(childName);
    expect(shape.description).toBe(childDescription);
    expect(shape.title).toBe(childTitle);
    expect(shape.hidden).toBe(childHidden);
  });
});
