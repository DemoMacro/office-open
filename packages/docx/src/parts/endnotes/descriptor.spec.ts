import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { endnotesDesc } from "./descriptor";

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
  const xml = endnotesDesc.stringify({ notes }, writeCtx as any)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return endnotesDesc.parse(el, readCtx);
}

describe("endnotesDesc round-trip", () => {
  it("round-trips endnote with text paragraphs", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["First endnote"]);
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(1);
    expect(result.notes.has(1)).toBe(true);
    const paragraphs = result.notes.get(1)!;
    expect(paragraphs).toHaveLength(1);
  });

  it("round-trips multiple endnotes", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["Endnote one"]);
    notes.set(2, ["Endnote two"]);
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(2);
  });

  it("round-trips empty endnotes", () => {
    const notes = new Map<number, string[]>();
    const result = roundTrip(notes);
    expect(result.notes.size).toBe(0);
  });

  it("skips system endnotes on parse", () => {
    const notes = new Map<number, string[]>();
    notes.set(1, ["User endnote"]);
    const result = roundTrip(notes);
    for (const id of result.notes.keys()) {
      expect(id).toBeGreaterThanOrEqual(1);
    }
  });
});
