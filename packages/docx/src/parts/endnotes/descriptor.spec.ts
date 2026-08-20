import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyBodyChild } from "../../body";
import { parseSectionChild } from "../../parse/body";
import { setNotesParseChild } from "../notes/shared";
import { setTableParseChild } from "../table/descriptor";
import { endnotesDesc } from "./descriptor";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  fileData: {},
  file: {},
  viewWrapper: { relationships: { toXml: () => "" } },
  stringifyChild: undefined,
};
(writeCtx as any).stringifyChild = (child: Parameters<typeof stringifyBodyChild>[0]) =>
  stringifyBodyChild(child, writeCtx as any);

setNotesParseChild(parseSectionChild);
setTableParseChild(parseSectionChild);

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(notes: Map<number, string[]>) {
  const xml = endnotesDesc.stringify({ notes }, writeCtx as any)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
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

  it("round-trips a block-level table before a paragraph", () => {
    const source =
      '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:endnote w:id="1"><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
      "<w:p><w:r><w:t>after</w:t></w:r></w:p></w:endnote></w:endnotes>";
    const root = parseXml(source).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const parsed = endnotesDesc.parse(root, readCtx);
    expect(parsed.notes.get(1)?.[0]).toMatchObject({ table: {} });

    const xml = endnotesDesc.stringify(parsed, writeCtx as any)!;
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain('<w:t xml:space="preserve">cell</w:t>');
    expect(xml).not.toContain("<w:endnoteRef/>");
  });
});
