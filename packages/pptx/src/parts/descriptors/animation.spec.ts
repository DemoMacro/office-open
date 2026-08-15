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

function parseTimingXml(xml: string) {
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return timingDesc.parse(el, readCtx);
}

function roundTrip(opts: TimingDescriptorOptions) {
  const xml = timingDesc.stringify(opts, writeCtx)!;
  if (!xml) return { entries: [] };
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
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
          shapeId: 2,
          options: { type: "fade", trigger: "onClick", duration: 500 },
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.entries).toHaveLength(1);
    const [entry] = result.entries;
    expect(entry?.shapeId).toBe(2);
    expect(entry?.options.type).toBe("fade");
    expect(entry?.options.trigger).toBe("onClick");
    expect(entry?.options.duration).toBe(500);
  });

  it("round-trips appear animation", () => {
    const opts: TimingDescriptorOptions = {
      entries: [
        {
          shapeId: 3,
          options: { type: "appear", trigger: "withPrevious" },
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.entries).toHaveLength(1);
    const [entry] = result.entries;
    expect(entry?.options.type).toBe("appear");
    expect(entry?.options.trigger).toBe("withPrevious");
  });

  it("round-trips animation with direction", () => {
    const opts: TimingDescriptorOptions = {
      entries: [
        {
          shapeId: 4,
          options: { type: "wipe", direction: "left", trigger: "afterPrevious", duration: 700 },
        },
      ],
    };
    const result = roundTrip(opts);
    const [entry] = result.entries;
    expect(entry?.options.type).toBe("wipe");
    expect(entry?.options.direction).toBe("left");
    expect(entry?.options.duration).toBe(700);
  });

  it("parses animate behavior attributes (calcmode/valueType/from/to/by)", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="clickEffect" presetClass="entr" presetID="10"><p:childTnLst><p:anim calcmode="lin" valueType="num" from="0" to="1" by="0.5"><p:cBhvr><p:cTn dur="500"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result.entries;
    expect(entry?.shapeId).toBe(4);
    expect(entry?.options.calcMode).toBe("lin");
    expect(entry?.options.valueType).toBe("num");
    expect(entry?.options.from).toBe("0");
    expect(entry?.options.to).toBe("1");
    expect(entry?.options.animBy).toBe("0.5");
    expect(entry?.options.attributeName).toBe("ppt_w");
  });

  it("parses animate-color color space and command behavior", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="withEffect" presetClass="emph" presetID="29"><p:childTnLst><p:animClr clrSpc="hsl"><p:cBhvr><p:cTn dur="300"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:animClr><p:cmd type="call" cmd="play"><p:cBhvr><p:cTn dur="300"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result.entries;
    expect(entry?.options.colorSpace).toBe("hsl");
    expect(entry?.options.emphasisType).toBe("colorChange");
    expect(entry?.options.commandType).toBe("call");
    expect(entry?.options.command).toBe("play");
  });

  it("parses iterate container attributes", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst><p:par><p:cTn><p:childTnLst><p:par><p:cTn nodeType="clickEffect" presetClass="entr" presetID="10"><p:childTnLst><p:anim calcmode="lin"><p:cBhvr><p:cTn dur="500"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:anim><p:iterate type="lt" backwards="1"><p:tmAbs tm="200"/></p:iterate></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const result = parseTimingXml(xml);
    const [entry] = result.entries;
    expect(entry?.shapeId).toBe(9);
    expect(entry?.options.iterate).toEqual({ type: "lt", backwards: true });
  });
});
