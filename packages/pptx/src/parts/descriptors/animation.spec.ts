import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { timingDesc } from "./animation";
import type { TimingDescriptorOptions } from "./animation";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TimingDescriptorOptions) {
  const xml = timingDesc.stringify(opts, writeCtx)!;
  if (!xml) return { entries: [] };
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return timingDesc.parse(el, readCtx);
}

describe("timingDesc round-trip", () => {
  it("round-trips empty entries", () => {
    const opts: TimingDescriptorOptions = { entries: [] };
    const result = roundTrip(opts);
    expect(result.entries).toHaveLength(0);
  });

  it("round-trips single fade animation", () => {
    const opts: TimingDescriptorOptions = {
      entries: [
        {
          spid: 2,
          options: { type: "fade", trigger: "onClick", duration: 500 },
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.entries).toHaveLength(1);
    expect(result.entries[0].spid).toBe(2);
    expect(result.entries[0].options.type).toBe("fade");
    expect(result.entries[0].options.trigger).toBe("onClick");
    expect(result.entries[0].options.duration).toBe(500);
  });

  it("round-trips appear animation", () => {
    const opts: TimingDescriptorOptions = {
      entries: [
        {
          spid: 3,
          options: { type: "appear", trigger: "withPrevious" },
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.entries).toHaveLength(1);
    expect(result.entries[0].options.type).toBe("appear");
    expect(result.entries[0].options.trigger).toBe("withPrevious");
  });

  it("round-trips animation with direction", () => {
    const opts: TimingDescriptorOptions = {
      entries: [
        {
          spid: 4,
          options: { type: "wipe", direction: "left", trigger: "afterPrevious", duration: 700 },
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.entries[0].options.type).toBe("wipe");
    expect(result.entries[0].options.direction).toBe("left");
    expect(result.entries[0].options.duration).toBe(700);
  });
});
