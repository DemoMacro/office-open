import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { lineShapeDesc, connectorShapeDesc } from "./line";
import type { LineShapeDescriptorOptions, ConnectorShapeDescriptorOptions } from "./line";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTripLine(opts: LineShapeDescriptorOptions) {
  const xml = lineShapeDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return lineShapeDesc.parse(el, readCtx);
}

function roundTripConnector(opts: ConnectorShapeDescriptorOptions) {
  const xml = connectorShapeDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return connectorShapeDesc.parse(el, readCtx);
}

describe("lineShapeDesc round-trip", () => {
  it("round-trips basic line coordinates", () => {
    const opts: LineShapeDescriptorOptions = {
      id: 10,
      name: "Test Line",
      x1: 0,
      y1: 0,
      x2: 200,
      y2: 100,
    };
    const result = roundTripLine(opts);

    expect(result.id).toBe(10);
    expect(result.name).toBe("Test Line");
    // Coordinates go through pixel->EMU->pixel conversion, expect rounding
    expect(result.x1).toBeCloseTo(0, 0);
    expect(result.y1).toBeCloseTo(0, 0);
    expect(result.x2).toBeCloseTo(200, 0);
    expect(result.y2).toBeCloseTo(100, 0);
  });

  it("round-trips line with default endpoints", () => {
    const opts: LineShapeDescriptorOptions = {
      id: 5,
    };
    const result = roundTripLine(opts);

    expect(result.id).toBe(5);
    expect(result.name).toBe("Line 5");
  });

  it("round-trips line with reversed coordinates (flip)", () => {
    const opts: LineShapeDescriptorOptions = {
      id: 3,
      x1: 200,
      y1: 150,
      x2: 50,
      y2: 10,
    };
    const result = roundTripLine(opts);

    expect(result.x1).toBeCloseTo(200, 0);
    expect(result.y1).toBeCloseTo(150, 0);
    expect(result.x2).toBeCloseTo(50, 0);
    expect(result.y2).toBeCloseTo(10, 0);
  });

  it("round-trips line with outline", () => {
    const opts: LineShapeDescriptorOptions = {
      id: 4,
      outline: { width: 2, color: "FF0000" },
    };
    const result = roundTripLine(opts);

    expect(result.outline).toBeDefined();
    const outline = result.outline as Record<string, unknown>;
    expect(outline.width).toBe(2);
  });
});

describe("connectorShapeDesc round-trip", () => {
  it("round-trips basic connector", () => {
    const opts: ConnectorShapeDescriptorOptions = {
      id: 20,
      name: "Test Connector",
      x1: 10,
      y1: 20,
      x2: 300,
      y2: 200,
    };
    const result = roundTripConnector(opts);

    expect(result.id).toBe(20);
    expect(result.name).toBe("Test Connector");
    expect(result.x1).toBeCloseTo(10, 0);
    expect(result.y1).toBeCloseTo(20, 0);
    expect(result.x2).toBeCloseTo(300, 0);
    expect(result.y2).toBeCloseTo(200, 0);
  });

  it("round-trips connector with arrowheads", () => {
    const opts: ConnectorShapeDescriptorOptions = {
      id: 21,
      x1: 0,
      y1: 0,
      x2: 100,
      y2: 0,
      endArrowhead: "triangle",
      beginArrowhead: "open",
    };
    const result = roundTripConnector(opts);

    // "triangle" maps identity; "open" -> "arrow" in XML, reversed back to "open" on parse
    expect(result.endArrowhead).toBe("triangle");
    expect(result.beginArrowhead).toBe("open");
  });

  it("round-trips connector with outline", () => {
    const opts: ConnectorShapeDescriptorOptions = {
      id: 22,
      outline: { width: 3, color: "00FF00" },
    };
    const result = roundTripConnector(opts);

    expect(result.outline).toBeDefined();
    const outline = result.outline as Record<string, unknown>;
    expect(outline.width).toBe(3);
  });
});
