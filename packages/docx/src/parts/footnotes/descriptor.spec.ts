import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { footnotesDesc } from "./descriptor";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  fileData: {},
  file: {},
  viewWrapper: { relationships: { toXml: () => "" } },
  stringifyChild: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(notes: Map<number, string[]>) {
  const xml = footnotesDesc.stringify({ notes }, writeCtx as any)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return footnotesDesc.parse(el, readCtx);
}

describe("footnotesDesc round-trip", () => {
  it("round-trips footnote with text paragraphs", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["First footnote"]);
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(1);
    expect(result.notes.has(1)).toBe(true);
    const paragraphs = result.notes.get(1)!;
    expect(paragraphs).toHaveLength(1);
  });

  it("round-trips multiple footnotes", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["Footnote one"]);
    notes.set(2, ["Footnote two"]);
    notes.set(3, ["Footnote three"]);
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(3);
    expect(result.notes.has(1)).toBe(true);
    expect(result.notes.has(2)).toBe(true);
    expect(result.notes.has(3)).toBe(true);
  });

  it("round-trips empty footnotes", () => {
    const notes = new Map<number, string[]>();
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(0);
  });

  it("skips system footnotes (separator, continuation) on parse", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["User footnote"]);
    const result = roundTrip(notes);
    // Only user footnotes should be in the result (not separator id=-1 or continuation id=0)
    for (const id of result.notes.keys()) {
      expect(id).toBeGreaterThanOrEqual(1);
    }
  });
});
