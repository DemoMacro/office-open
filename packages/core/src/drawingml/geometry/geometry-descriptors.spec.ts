import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import type { CustomGeometryOptions } from "./custom-geometry";
import { presetGeometryDesc, customGeometryDesc } from "./geometry-descriptors";
import type { PresetGeometryOptions } from "./preset-geometry";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("presetGeometryDesc", () => {
  it("round-trips basic preset geometry", () => {
    const opts: PresetGeometryOptions = { preset: "rect" };
    const result = roundTrip(presetGeometryDesc, opts);
    expect(result.preset).toBe("rect");
    expect(result.adjustmentValues).toBeUndefined();
  });

  it("round-trips preset geometry with adjustment values", () => {
    const opts: PresetGeometryOptions = {
      preset: "roundRect",
      adjustmentValues: [{ name: "adj", formula: "val 16667" }],
    };
    const result = roundTrip(presetGeometryDesc, opts);
    expect(result.preset).toBe("roundRect");
    expect(result.adjustmentValues).toBeDefined();
    expect(result.adjustmentValues).toHaveLength(1);
    expect(result.adjustmentValues![0].name).toBe("adj");
    expect(result.adjustmentValues![0].formula).toBe("val 16667");
  });

  it("round-trips preset geometry with multiple adjustment values", () => {
    const opts: PresetGeometryOptions = {
      preset: "chevron",
      adjustmentValues: [
        { name: "adj1", formula: "val 50000" },
        { name: "adj2", formula: "val 25000" },
      ],
    };
    const result = roundTrip(presetGeometryDesc, opts);
    expect(result.preset).toBe("chevron");
    expect(result.adjustmentValues).toHaveLength(2);
  });
});

describe("customGeometryDesc", () => {
  it("round-trips simple custom geometry (triangle)", () => {
    const opts: CustomGeometryOptions = {
      pathList: [
        {
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            { command: "lineTo", point: { x: "100000", y: "0" } },
            { command: "lineTo", point: { x: "50000", y: "100000" } },
            { command: "close" },
          ],
        },
      ],
    };
    const result = roundTrip(customGeometryDesc, opts);
    expect(result.pathList).toBeDefined();
    expect(result.pathList).toHaveLength(1);
    expect(result.pathList![0].commands).toHaveLength(4);
    expect(result.pathList![0].commands![0].command).toBe("moveTo");
    expect(result.pathList![0].commands![3].command).toBe("close");
  });

  it("round-trips custom geometry with all features", () => {
    const opts: CustomGeometryOptions = {
      adjustmentValues: [{ name: "adj", formula: "val 50000" }],
      guides: [{ name: "gd1", formula: "val 10000" }],
      adjustHandles: [
        {
          type: "xy",
          position: { x: "50000", y: "0" },
          guideRefX: "adj",
          minX: "0",
          maxX: "100000",
        },
      ],
      connectionSites: [{ angle: "0", position: { x: "50000", y: "0" } }],
      textRectangle: { left: "10000", top: "10000", right: "90000", bottom: "90000" },
      pathList: [
        {
          w: 100000,
          h: 100000,
          fill: "normal",
          stroke: true,
          extrusionOk: false,
          commands: [
            { command: "moveTo", point: { x: "50000", y: "0" } },
            { command: "lineTo", point: { x: "100000", y: "100000" } },
            { command: "lineTo", point: { x: "0", y: "100000" } },
            { command: "close" },
          ],
        },
      ],
    };
    const result = roundTrip(customGeometryDesc, opts);
    expect(result.adjustmentValues).toHaveLength(1);
    expect(result.guides).toHaveLength(1);
    expect(result.adjustHandles).toHaveLength(1);
    expect(result.adjustHandles![0].type).toBe("xy");
    expect(result.connectionSites).toHaveLength(1);
    expect(result.textRectangle).toEqual({
      left: "10000",
      top: "10000",
      right: "90000",
      bottom: "90000",
    });
    expect(result.pathList).toHaveLength(1);
    expect(result.pathList![0].w).toBe(100000);
    expect(result.pathList![0].h).toBe(100000);
  });

  it("round-trips custom geometry with arc and bezier commands", () => {
    const opts: CustomGeometryOptions = {
      pathList: [
        {
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            {
              command: "arcTo",
              widthRadius: "50000",
              heightRadius: "50000",
              startAngle: "0",
              sweepAngle: "5400000",
            },
            {
              command: "quadBezTo",
              points: [
                { x: "25000", y: "0" },
                { x: "50000", y: "25000" },
              ],
            },
            {
              command: "cubicBezTo",
              points: [
                { x: "25000", y: "0" },
                { x: "50000", y: "25000" },
                { x: "50000", y: "50000" },
              ],
            },
          ],
        },
      ],
    };
    const result = roundTrip(customGeometryDesc, opts);
    const cmds = result.pathList![0].commands!;
    expect(cmds[0].command).toBe("moveTo");
    expect(cmds[1].command).toBe("arcTo");
    expect(cmds[2].command).toBe("quadBezTo");
    expect(cmds[3].command).toBe("cubicBezTo");
  });

  it("round-trips custom geometry with polar adjust handle", () => {
    const opts: CustomGeometryOptions = {
      adjustHandles: [
        {
          type: "polar",
          position: { x: "50000", y: "50000" },
          guideRefRadius: "adj1",
          minRadius: "0",
          maxRadius: "100000",
          guideRefAngle: "adj2",
          minAngle: "0",
          maxAngle: "21600000",
        },
      ],
      pathList: [{ commands: [{ command: "close" }] }],
    };
    const result = roundTrip(customGeometryDesc, opts);
    expect(result.adjustHandles).toHaveLength(1);
    expect(result.adjustHandles![0].type).toBe("polar");
  });
});
