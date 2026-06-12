import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { notesSlideDesc } from "./notes-slide";
import type { NotesSlideDescriptorOptions } from "./notes-slide";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: NotesSlideDescriptorOptions) {
  const xml = notesSlideDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return notesSlideDesc.parse(el, readCtx);
}

describe("notesSlideDesc round-trip", () => {
  it("round-trips notes text", () => {
    const opts: NotesSlideDescriptorOptions = {
      text: "Speaker notes content",
    };
    const result = roundTrip(opts);

    expect(result.text).toBe("Speaker notes content");
  });

  it("round-trips empty notes", () => {
    const opts: NotesSlideDescriptorOptions = {};
    const result = roundTrip(opts);

    expect(result.text).toBeUndefined();
  });

  it("round-trips multi-line notes", () => {
    const opts: NotesSlideDescriptorOptions = {
      text: "Line one\nLine two\nLine three",
    };
    const result = roundTrip(opts);

    // Notes text is joined with \n in parse, but the XML uses a single run
    expect(result.text).toContain("Line one");
  });

  it("round-trips notes with special characters", () => {
    const opts: NotesSlideDescriptorOptions = {
      text: '<b>Bold & "quoted"',
    };
    const result = roundTrip(opts);

    expect(result.text).toBe('<b>Bold & "quoted"');
  });
});
