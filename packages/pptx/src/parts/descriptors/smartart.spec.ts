import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import {
  getColorXml,
  getLayoutXml,
  getStyleXml,
  stringifyColorDefinitionPart,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { SmartArtOptions } from "../smartart";
import { smartArtDesc } from "./smartart";

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

function roundTrip(opts: SmartArtOptions) {
  const xml = smartArtDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return smartArtDesc.parse(el, readCtx);
}

describe("smartArtDesc round-trip", () => {
  it("round-trips basic position and name", () => {
    const opts: SmartArtOptions = {
      id: 100,
      nodes: [],
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
    const opts: SmartArtOptions = {
      id: 200,
      nodes: [],
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("Diagram 200");
    expect(result.x).toBe(0);
    expect(result.y).toBe(0);
    // default 100px = 952500 EMU
    expect(result.width).toBe(952500);
    expect(result.height).toBe(952500);
  });

  it("registers SmartArt data in context when nodes provided", () => {
    smartArtRegistry.clear();
    const opts: SmartArtOptions = {
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

  it("round-trips with empty nodes", () => {
    const opts: SmartArtOptions = {
      id: 400,
      nodes: [],
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
    const opts: SmartArtOptions = {
      id: 500,
      nodes: [],
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

describe("smartArtDesc custom definitions", () => {
  const customLayout: LayoutDefinitionOptions = {
    uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/customSteps",
    layoutNode: { name: "diagram", children: [{ algorithm: { type: "snake" } }] },
  };
  const customStyle: StyleDefinitionOptions = {
    uniqueId: "urn:microsoft.com/office/officeart/2005/8/quickstyle/customSteps",
    styleLabels: [{ name: "step" }],
  };
  const customColors: ColorDefinitionOptions = {
    uniqueId: "urn:microsoft.com/office/officeart/2005/8/colors/customSteps",
    styleLabels: [{ name: "step", fillColorList: { colors: [{ value: "accent1" }] } }],
  };

  it("registers custom definitions verbatim in the SmartArt entry", () => {
    smartArtRegistry.clear();
    const opts: SmartArtOptions = {
      id: 600,
      nodes: [{ text: "Step" }],
      layout: customLayout,
      style: customStyle,
      color: customColors,
    };
    smartArtDesc.stringify(opts, writeCtx);
    const entry = smartArtRegistry.get("smartart_1024") as {
      layout: unknown;
      style: unknown;
      color: unknown;
    };
    expect(entry.layout).toBe(customLayout);
    expect(entry.style).toBe(customStyle);
    expect(entry.color).toBe(customColors);
  });

  it("parses custom definition parts back into structured options", () => {
    const xml = smartArtDesc.stringify(
      {
        id: 601,
        nodes: [{ text: "Step" }],
        layout: customLayout,
        style: customStyle,
        color: customColors,
      },
      writeCtx,
    );
    if (!xml) throw new Error("stringify returned undefined");
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const partFor = (body: string) => parseXml(body).elements?.[0];
    const ctx = {
      resolveRelationship: (rid: string) => {
        if (rid.startsWith("{smartart-lo:")) return "ppt/diagrams/layout1.xml";
        if (rid.startsWith("{smartart-qs:")) return "ppt/diagrams/quickStyle1.xml";
        if (rid.startsWith("{smartart-cs:")) return "ppt/diagrams/colors1.xml";
        return undefined;
      },
      getPart: (path: string) => {
        if (path.endsWith("layout1.xml"))
          return partFor(stringifyLayoutDefinitionPart(customLayout));
        if (path.endsWith("quickStyle1.xml"))
          return partFor(stringifyStyleDefinitionPart(customStyle));
        if (path.endsWith("colors1.xml"))
          return partFor(stringifyColorDefinitionPart(customColors));
        return undefined;
      },
      getRaw: () => undefined,
    } as unknown as ReadContext;

    const result = smartArtDesc.parse(el, ctx);
    expect(result.layout).toEqual(customLayout);
    expect(result.style).toEqual(customStyle);
    expect(result.color).toEqual(customColors);
  });

  it("folds built-in definition stubs back to their id string", () => {
    const xml = smartArtDesc.stringify(
      {
        id: 602,
        nodes: [{ text: "Step" }],
        layout: "process1",
        style: "simple1",
        color: "accent1_2",
      },
      writeCtx,
    );
    if (!xml) throw new Error("stringify returned undefined");
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const partFor = (body: string) => parseXml(body).elements?.[0];
    const ctx = {
      resolveRelationship: (rid: string) => {
        if (rid.startsWith("{smartart-lo:")) return "ppt/diagrams/layout1.xml";
        if (rid.startsWith("{smartart-qs:")) return "ppt/diagrams/quickStyle1.xml";
        if (rid.startsWith("{smartart-cs:")) return "ppt/diagrams/colors1.xml";
        return undefined;
      },
      getPart: (path: string) => {
        if (path.endsWith("layout1.xml")) return partFor(getLayoutXml("process1"));
        if (path.endsWith("quickStyle1.xml")) return partFor(getStyleXml("simple1"));
        if (path.endsWith("colors1.xml")) return partFor(getColorXml("accent1_2"));
        return undefined;
      },
      getRaw: () => undefined,
    } as unknown as ReadContext;

    const result = smartArtDesc.parse(el, ctx);
    expect(result.layout).toBe("process1");
    expect(result.style).toBe("simple1");
    expect(result.color).toBe("accent1_2");
  });
});
