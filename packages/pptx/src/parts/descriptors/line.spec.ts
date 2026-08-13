import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { ConnectorOptions, LineShapeOptions } from "@shared/shape/line-shape";
import { describe, expect, it } from "vite-plus/test";

import { lineShapeDesc, connectorShapeDesc } from "./line";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTripLine(opts: LineShapeOptions) {
  const xml = lineShapeDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return lineShapeDesc.parse(el, readCtx);
}

function roundTripConnector(opts: ConnectorOptions) {
  const xml = connectorShapeDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return connectorShapeDesc.parse(el, readCtx);
}

describe("lineShapeDesc round-trip", () => {
  it("round-trips basic line coordinates", () => {
    const opts: LineShapeOptions = {
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
    const opts: LineShapeOptions = {
      id: 5,
    };
    const result = roundTripLine(opts);

    expect(result.id).toBe(5);
    expect(result.name).toBe("Line 5");
  });

  it("round-trips line with reversed coordinates (flip)", () => {
    const opts: LineShapeOptions = {
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
    const opts: LineShapeOptions = {
      id: 4,
      outline: { type: "solidFill", color: { value: "FF0000" }, width: 2 },
    };
    const result = roundTripLine(opts);

    expect(result.outline).toBeDefined();
    const outline = result.outline as Record<string, unknown>;
    expect(outline.width).toBe(2);
  });
});

describe("connectorShapeDesc round-trip", () => {
  it("round-trips basic connector", () => {
    const opts: ConnectorOptions = {
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
    const opts: ConnectorOptions = {
      id: 21,
      x1: 0,
      y1: 0,
      x2: 100,
      y2: 0,
      outline: {
        headEnd: { type: "triangle" },
        tailEnd: { type: "arrow" },
      },
    };
    const result = roundTripConnector(opts);

    expect(result.outline?.headEnd?.type).toBe("triangle");
    expect(result.outline?.tailEnd?.type).toBe("arrow");
  });

  it("round-trips connector with outline", () => {
    const opts: ConnectorOptions = {
      id: 22,
      outline: { type: "solidFill", color: { value: "00FF00" }, width: 3 },
    };
    const result = roundTripConnector(opts);

    expect(result.outline).toBeDefined();
    const outline = result.outline as Record<string, unknown>;
    expect(outline.width).toBe(3);
  });
});
