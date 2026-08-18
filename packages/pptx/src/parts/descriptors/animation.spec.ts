import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { AnimationEntry } from "@shared/animation/timing";
import { describe, expect, it } from "vite-plus/test";

import { timingDesc } from "./animation";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function parseTimingXml(xml: string) {
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return timingDesc.parse(el, readCtx);
}

function roundTrip(entries: AnimationEntry[]): AnimationEntry[] {
  const xml = timingDesc.stringify(entries, writeCtx)!;
  if (!xml) return [];
  const result = parseTimingXml(xml);
  if (!Array.isArray(result)) {
    throw new Error(`expected structured entries, got verbatim: ${result.slice(0, 120)}`);
  }
  return result;
}

describe("timingDesc round-trip", () => {
  it("round-trips empty entries", () => {
    const entries: AnimationEntry[] = [];
    const result = roundTrip(entries);
    expect(result).toHaveLength(0);
  });

  it("round-trips single fade animation", () => {
    const entries: AnimationEntry[] = [
      { shapeId: 2, type: "fade", trigger: "onClick", duration: 500 },
    ];
    const result = roundTrip(entries);
    expect(result).toHaveLength(1);
    const [entry] = result;
    expect(entry?.shapeId).toBe(2);
    expect(entry?.type).toBe("fade");
    expect(entry?.trigger).toBe("onClick");
    expect(entry?.duration).toBe(500);
  });

  it("round-trips appear animation", () => {
    const entries: AnimationEntry[] = [{ shapeId: 3, type: "appear", trigger: "withPrevious" }];
    const result = roundTrip(entries);
    expect(result).toHaveLength(1);
    const [entry] = result;
    expect(entry?.type).toBe("appear");
    expect(entry?.trigger).toBe("withPrevious");
  });

  it("round-trips animation with direction", () => {
    const entries: AnimationEntry[] = [
      { shapeId: 4, type: "wipe", direction: "left", trigger: "afterPrevious", duration: 700 },
    ];
    const result = roundTrip(entries);
    const [entry] = result;
    expect(entry?.type).toBe("wipe");
    expect(entry?.direction).toBe("left");
    expect(entry?.duration).toBe(700);
  });

  it("round-trips animate behavior attributes (calcmode/valueType/from/to/by)", () => {
    const entries: AnimationEntry[] = [
      {
        shapeId: 4,
        attributeName: "ppt_w",
        calcMode: "lin",
        valueType: "num",
        from: "0",
        to: "1",
        animBy: "0.5",
        duration: 500,
      },
    ];
    const result = roundTrip(entries);
    const [entry] = result;
    expect(entry?.shapeId).toBe(4);
    expect(entry?.calcMode).toBe("lin");
    expect(entry?.valueType).toBe("num");
    expect(entry?.from).toBe("0");
    expect(entry?.to).toBe("1");
    expect(entry?.animBy).toBe("0.5");
    expect(entry?.attributeName).toBe("ppt_w");
  });

  it("round-trips animate-color color space and command behavior", () => {
    const entries: AnimationEntry[] = [
      {
        shapeId: 7,
        class: "emph",
        emphasisType: "colorChange",
        colorSpace: "hsl",
        colorTo: "FF0000",
        duration: 300,
        trigger: "withPrevious",
      },
      {
        shapeId: 7,
        commandType: "call",
        command: "play",
        duration: 300,
        trigger: "withPrevious",
      },
    ];
    const result = roundTrip(entries);
    expect(result).toHaveLength(2);
    const [colorEntry, cmdEntry] = result;
    expect(colorEntry?.colorSpace).toBe("hsl");
    expect(colorEntry?.emphasisType).toBe("colorChange");
    expect(cmdEntry?.commandType).toBe("call");
    expect(cmdEntry?.command).toBe("play");
  });

  it("round-trips iterate container attributes", () => {
    const entries: AnimationEntry[] = [
      {
        shapeId: 9,
        attributeName: "ppt_w",
        iterate: { type: "lt", backwards: true, interval: 200 },
        duration: 500,
      },
    ];
    const result = roundTrip(entries);
    const [entry] = result;
    expect(entry?.shapeId).toBe(9);
    expect(entry?.iterate).toEqual({ type: "lt", backwards: true, interval: 200 });
  });

  it("falls back to verbatim inner XML when the model cannot rebuild the tree", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:audio><p:cMediaNode vol="80000"><p:cTn id="3" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="5"/></p:tgtEl></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    expect(typeof result).toBe("string");
    expect(result).toContain('<p:audio><p:cMediaNode vol="80000">');
    // The verbatim form re-emits byte-identical timing.
    expect(timingDesc.stringify(result, writeCtx)).toBe(xml);
  });

  it("falls back to verbatim inner XML for tmRoot-only timing", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    expect(typeof result).toBe("string");
    expect(timingDesc.stringify(result, writeCtx)).toBe(xml);
  });
});
