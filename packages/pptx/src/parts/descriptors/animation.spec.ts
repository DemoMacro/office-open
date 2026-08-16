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

function roundTrip(entries: AnimationEntry[]) {
  const xml = timingDesc.stringify(entries, writeCtx)!;
  if (!xml) return [];
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return timingDesc.parse(el, readCtx);
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

  it("parses animate behavior attributes (calcmode/valueType/from/to/by)", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="clickEffect" presetClass="entr" presetID="10"><p:childTnLst><p:anim calcmode="lin" valueType="num" from="0" to="1" by="0.5"><p:cBhvr><p:cTn dur="500"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result;
    expect(entry?.shapeId).toBe(4);
    expect(entry?.calcMode).toBe("lin");
    expect(entry?.valueType).toBe("num");
    expect(entry?.from).toBe("0");
    expect(entry?.to).toBe("1");
    expect(entry?.animBy).toBe("0.5");
    expect(entry?.attributeName).toBe("ppt_w");
  });

  it("parses animate-color color space and command behavior", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="withEffect" presetClass="emph" presetID="29"><p:childTnLst><p:animClr clrSpc="hsl"><p:cBhvr><p:cTn dur="300"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:animClr><p:cmd type="call" cmd="play"><p:cBhvr><p:cTn dur="300"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result;
    expect(entry?.colorSpace).toBe("hsl");
    expect(entry?.emphasisType).toBe("colorChange");
    expect(entry?.commandType).toBe("call");
    expect(entry?.command).toBe("play");
  });

  it("parses iterate container attributes", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="clickEffect" presetClass="entr" presetID="10"><p:childTnLst><p:anim calcmode="lin"><p:cBhvr><p:cTn dur="500"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:anim><p:iterate type="lt" backwards="1"><p:tmAbs tm="200"/></p:iterate></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result;
    expect(entry?.shapeId).toBe(9);
    expect(entry?.iterate).toEqual({ type: "lt", backwards: true });
  });
});
