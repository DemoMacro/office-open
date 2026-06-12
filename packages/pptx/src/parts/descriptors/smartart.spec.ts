import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { smartArtDesc } from "./smartart";
import type { SmartArtDescriptorOptions } from "./smartart";

// ── Mock contexts ──

const smartArtRegistry = new Map<string, unknown>();

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  nextSmartArtKey: () => "smartart_1024",
  addSmartArt: (key: string, data: unknown) => smartArtRegistry.set(key, data),
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: SmartArtDescriptorOptions) {
  const xml = smartArtDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return smartArtDesc.parse(el, readCtx);
}

describe("smartArtDesc round-trip", () => {
  it("round-trips basic position and name", () => {
    const opts: SmartArtDescriptorOptions = {
      id: 100,
      name: "Test Diagram",
      x: 50,
      y: 60,
      width: 200,
      height: 150,
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("Test Diagram");
    expect(result.x).toBe(50);
    expect(result.y).toBe(60);
    expect(result.width).toBe(200);
    expect(result.height).toBe(150);
  });

  it("round-trips position with defaults", () => {
    const opts: SmartArtDescriptorOptions = {
      id: 200,
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("Diagram 200");
    expect(result.x).toBe(0);
    expect(result.y).toBe(0);
    expect(result.width).toBe(100);
    expect(result.height).toBe(100);
  });

  it("registers SmartArt data in context when nodes provided", () => {
    smartArtRegistry.clear();
    const opts: SmartArtDescriptorOptions = {
      id: 300,
      smartArtKey: "smartart_test",
      nodes: [{ text: "Root", children: [{ text: "Child" }] }],
      layout: "process1",
      style: "moderate1",
      color: "colorful1",
    };
    smartArtDesc.stringify(opts, writeCtx);

    expect(smartArtRegistry.has("smartart_test")).toBe(true);
  });

  it("round-trips without nodes", () => {
    const opts: SmartArtDescriptorOptions = {
      id: 400,
      name: "Empty Diagram",
      x: 10,
      y: 20,
      width: 300,
      height: 200,
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("Empty Diagram");
    expect(result.x).toBe(10);
    expect(result.y).toBe(20);
    expect(result.width).toBe(300);
    expect(result.height).toBe(200);
  });

  it("round-trips large EMU values correctly", () => {
    const opts: SmartArtDescriptorOptions = {
      id: 500,
      x: 1024,
      y: 768,
      width: 1920,
      height: 1080,
    };
    const result = roundTrip(opts);

    expect(result.x).toBe(1024);
    expect(result.y).toBe(768);
    expect(result.width).toBe(1920);
    expect(result.height).toBe(1080);
  });
});
