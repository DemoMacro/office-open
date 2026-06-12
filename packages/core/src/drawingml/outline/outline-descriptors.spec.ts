import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import type { OutlineOptions } from "./outline";
import { outlineDesc } from "./outline-descriptors";

function roundTrip(opts: OutlineOptions): OutlineOptions {
  const xml = stringify(outlineDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(outlineDesc, el, {} as any);
}

describe("outlineDesc", () => {
  it("round-trips outline with noFill", () => {
    const opts: OutlineOptions = { type: "noFill" };
    const result = roundTrip(opts);
    expect(result.type).toBe("noFill");
  });

  it("round-trips outline with solidFill color", () => {
    const opts: OutlineOptions = {
      type: "solidFill",
      color: { value: "FF0000" },
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("solidFill");
    expect(result.color).toEqual({ value: "FF0000" });
  });

  it("round-trips outline with gradientFill", () => {
    const opts: OutlineOptions = {
      type: "gradFill",
      gradientFill: {
        stops: [
          { position: 0, color: { value: "FF0000" } },
          { position: 100000, color: { value: "0000FF" } },
        ],
        shade: { angle: 5400000, scaled: true },
      },
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("gradFill");
    expect(result.gradientFill).toBeDefined();
    expect(result.gradientFill!.stops).toHaveLength(2);
    expect(result.gradientFill!.stops[0].position).toBe(0);
    expect(result.gradientFill!.stops[1].position).toBe(100000);
  });

  it("round-trips width, cap, compoundLine, align", () => {
    const opts: OutlineOptions = {
      type: "noFill",
      width: 9525,
      cap: "round",
      compoundLine: "double",
      align: "center",
    };
    const result = roundTrip(opts);
    expect(result.width).toBe(9525);
    expect(result.cap).toBe("round");
    expect(result.compoundLine).toBe("double");
    expect(result.align).toBe("center");
  });

  it("round-trips preset dash", () => {
    const opts: OutlineOptions = {
      type: "noFill",
      dash: "dash",
    };
    const result = roundTrip(opts);
    expect(result.dash).toBe("dash");
  });

  it("round-trips custom dash", () => {
    const opts: OutlineOptions = {
      type: "noFill",
      customDash: [
        { d: "100000", sp: "50000" },
        { d: "200000", sp: "100000" },
      ],
    };
    const result = roundTrip(opts);
    expect(result.customDash).toBeDefined();
    expect(result.customDash).toHaveLength(2);
    expect(result.customDash![0].d).toBe("100000");
    expect(result.customDash![0].sp).toBe("50000");
  });

  it("round-trips join styles", () => {
    const result1 = roundTrip({ type: "noFill", join: "round" });
    expect(result1.join).toBe("round");

    const result2 = roundTrip({ type: "noFill", join: "bevel" });
    expect(result2.join).toBe("bevel");

    const result3 = roundTrip({ type: "noFill", join: "miter", miterLimit: 80000 });
    expect(result3.join).toBe("miter");
    expect(result3.miterLimit).toBe(80000);
  });

  it("round-trips headEnd and tailEnd", () => {
    const opts: OutlineOptions = {
      type: "noFill",
      headEnd: { type: "stealth", width: "medium", length: "medium" },
      tailEnd: { type: "arrow" },
    };
    const result = roundTrip(opts);
    expect(result.headEnd).toBeDefined();
    expect(result.headEnd!.type).toBe("stealth");
    expect(result.headEnd!.width).toBe("medium");
    expect(result.headEnd!.length).toBe("medium");
    expect(result.tailEnd).toBeDefined();
    expect(result.tailEnd!.type).toBe("arrow");
  });
});
