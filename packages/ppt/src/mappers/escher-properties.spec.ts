import { describe, expect, it } from "vitest";

import { PptParseError } from "../errors";
import type { RecordNode } from "../records";
import { SHAPE_TYPE_PRESETS, SHAPE_TYPE_WARPS, readShapeProperties } from "./escher-properties";

const PROPERTY_TABLE_OFFSET = 8;

function createRecord(instance: number, length: number): { record: RecordNode; view: DataView } {
  const buffer = new ArrayBuffer(length);
  const view = new DataView(buffer);
  view.setUint16(0, 3 | (instance << 4), true);
  view.setUint16(2, 0xf00b, true);
  view.setUint32(4, length - 8, true);
  return {
    view,
    record: {
      offset: 0,
      version: 3,
      instance,
      type: 0xf00b,
      length: length - 8,
      end: length,
      children: [],
    },
  };
}

function writeProperty(view: DataView, index: number, id: number, value: number, flags = 0): void {
  view.setUint16(PROPERTY_TABLE_OFFSET + index * 6, id | flags, true);
  view.setUint32(PROPERTY_TABLE_OFFSET + index * 6 + 2, value, true);
}

describe("SHAPE_TYPE_PRESETS", () => {
  it("uses the authoritative MSO_SPT mappings", () => {
    expect(SHAPE_TYPE_PRESETS).toMatchObject({
      1: "rect",
      2: "roundRect",
      3: "ellipse",
      4: "diamond",
      5: "triangle",
      6: "rtTriangle",
      7: "parallelogram",
      8: "trapezoid",
      9: "hexagon",
      10: "octagon",
      11: "plus",
    });
  });
});

describe("SHAPE_TYPE_WARPS", () => {
  it("maps exact legacy WordArt presets", () => {
    expect(SHAPE_TYPE_WARPS).toMatchObject({
      24: "textPlain",
      137: "textStop",
      146: "textCircle",
      156: "textWave1",
      167: "textDeflateInflateDeflate",
      175: "textCanDown",
    });
  });
});

describe("readShapeProperties", () => {
  it("projects complex freeform vertices and segments", () => {
    const { record, view } = createRecord(6, 74);
    writeProperty(view, 0, 320, 0);
    writeProperty(view, 1, 321, 0);
    writeProperty(view, 2, 322, 100);
    writeProperty(view, 3, 323, 80);
    writeProperty(view, 4, 325, 14, 0x8000);
    writeProperty(view, 5, 326, 16, 0x8000);

    const vertices = 44;
    view.setUint32(vertices, 2, true);
    view.setUint16(vertices + 4, 4, true);
    view.setInt16(vertices + 6, 0, true);
    view.setInt16(vertices + 8, 0, true);
    view.setInt16(vertices + 10, 100, true);
    view.setInt16(vertices + 12, 80, true);

    const segments = 58;
    view.setUint16(segments, 3, true);
    view.setUint16(segments + 2, 0, true);
    view.setUint16(segments + 4, 2 << 13, true);
    view.setUint16(segments + 8, 1, true);
    view.setUint16(segments + 12, 4 << 13, true);

    const style = readShapeProperties(view, record);
    expect(style.customGeometry).toEqual({
      pathList: [
        {
          w: 100,
          h: 80,
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            { command: "lineTo", point: { x: "100", y: "80" } },
          ],
        },
      ],
    });
  });

  it("rejects truncated complex property data", () => {
    const { record, view } = createRecord(1, 20);
    writeProperty(view, 0, 325, 16, 0x8000);
    view.setUint32(14, 2, true);
    view.setUint16(18, 4, true);

    expect(() => readShapeProperties(view, record)).toThrow(PptParseError);
    try {
      readShapeProperties(view, record);
    } catch (error) {
      expect(error).toBeInstanceOf(PptParseError);
      expect((error as PptParseError).context.reason).toBe("truncated-complex-property");
    }
  });

  it("projects a segment-less freeform as a closed vertex path", () => {
    const { record, view } = createRecord(4, 64);
    writeProperty(view, 0, 325, 16, 0x8000);
    writeProperty(view, 1, 326, 8, 0x8000);

    const vertices = 32;
    view.setUint32(vertices, 2, true);
    view.setUint16(vertices + 4, 0xfff0, true);
    view.setInt16(vertices + 6, 10, true);
    view.setInt16(vertices + 8, 20, true);
    view.setInt16(vertices + 10, 50, true);
    view.setInt16(vertices + 12, 80, true);

    const segments = 48;
    view.setUint16(segments, 0, true);
    view.setUint16(segments + 2, 0, true);

    const style = readShapeProperties(view, record);
    expect(style.customGeometry).toEqual({
      pathList: [
        {
          w: 40,
          h: 60,
          commands: [
            { command: "moveTo", point: { x: "10", y: "20" } },
            { command: "lineTo", point: { x: "50", y: "80" } },
            { command: "close" },
          ],
        },
      ],
    });
  });
});
