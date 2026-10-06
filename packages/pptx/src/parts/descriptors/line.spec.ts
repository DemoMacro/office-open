import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { ConnectorOptions, LineShapeOptions } from "@shared/shape/line-shape";
import { describe, expect, it } from "vite-plus/test";

import { lineShapeDesc, connectorShapeDesc } from "./line";

const writeCtx = { registerShapeId() {} } as unknown as WriteContext;
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
      properties: { outline: { type: "solidFill", color: { value: "FF0000" }, width: 2 } },
    };
    const result = roundTripLine(opts);

    expect(result.properties?.outline).toBeDefined();
    const outline = result.properties!.outline as Record<string, unknown>;
    expect(outline.width).toBe(2);
  });

  it("round-trips rotated black-and-white text boxes", () => {
    const result = roundTripLine({
      id: 7,
      x1: 10,
      y1: 20,
      x2: 110,
      y2: 120,
      rotation: 90,
      blackWhiteMode: "auto",
      textBox: true,
    });

    expect(result.rotation).toBe(90);
    expect(result.blackWhiteMode).toBe("auto");
    expect(result.textBox).toBe(true);
  });
});

describe("connectorShapeDesc round-trip", () => {
  it("round-trips flipV on a zero-height connector", () => {
    const writeXml = connectorShapeDesc.stringify(
      { id: 21, x1: 10, y1: 20, x2: 10, y2: 20, flipVertical: true },
      writeCtx,
    );
    expect(writeXml).toContain('flipV="1"');
    const element = parseXml(writeXml!)?.elements?.[0];
    if (!element) throw new Error("parsed document has no root element");
    expect(connectorShapeDesc.parse(element, readCtx).flipVertical).toBe(true);

    const explicitFalseXml = connectorShapeDesc.stringify(
      { id: 22, x1: 10, y1: 20, x2: 20, y2: 30, flipVertical: false },
      writeCtx,
    );
    expect(explicitFalseXml).toContain('flipV="0"');
    const falseElement = parseXml(explicitFalseXml!)?.elements?.[0];
    if (!falseElement) throw new Error("parsed document has no root element");
    expect(connectorShapeDesc.parse(falseElement, readCtx).flipVertical).toBe(false);
  });

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
      properties: {
        outline: {
          headEnd: { type: "triangle" },
          tailEnd: { type: "arrow" },
        },
      },
    };
    const result = roundTripConnector(opts);

    expect(result.properties?.outline?.headEnd?.type).toBe("triangle");
    expect(result.properties?.outline?.tailEnd?.type).toBe("arrow");
  });

  it("round-trips connector userDrawn marking", () => {
    const opts: ConnectorOptions = { id: 23, userDrawn: true };
    const xml = connectorShapeDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain('<p:nvPr userDrawn="1"/>');
    expect(roundTripConnector(opts).userDrawn).toBe(true);
  });

  it("round-trips connector with outline", () => {
    const opts: ConnectorOptions = {
      id: 22,
      properties: { outline: { type: "solidFill", color: { value: "00FF00" }, width: 3 } },
    };
    const result = roundTripConnector(opts);

    expect(result.properties?.outline).toBeDefined();
    const outline = result.properties!.outline as Record<string, unknown>;
    expect(outline.width).toBe(3);
  });

  it("round-trips connector rotation and black-and-white mode", () => {
    const result = roundTripConnector({
      id: 23,
      x1: 10,
      y1: 20,
      x2: 110,
      y2: 120,
      rotation: 270,
      blackWhiteMode: "gray",
    });

    expect(result.rotation).toBe(270);
    expect(result.blackWhiteMode).toBe("gray");
  });
});
