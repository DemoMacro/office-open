import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyBodyChild } from "../../body";
import { parseSectionChild } from "../../parse/body";
import { setNotesParseChild } from "../notes/shared";
import { setTableParseChild } from "../table/descriptor";
import { footnotesDesc } from "./descriptor";

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
  const xml = footnotesDesc.stringify({ notes }, writeCtx as any)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
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

  it("round-trips tables and only injects a reference mark into a first paragraph", () => {
    const source =
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
      "<w:p><w:r><w:t>after</w:t></w:r></w:p></w:footnote></w:footnotes>";
    const root = parseXml(source).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const parsed = footnotesDesc.parse(root, readCtx);
    expect(parsed.notes.get(1)?.[0]).toMatchObject({ table: {} });

    const xml = footnotesDesc.stringify(parsed, writeCtx as any)!;
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain("<w:t>cell</w:t>");
    expect(xml).not.toContain("<w:footnoteRef/>");
  });

  it("preserves absent and custom footnote reference marks", () => {
    const absent =
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>custom note</w:t></w:r></w:p></w:footnote>' +
      "</w:footnotes>";
    const absentRoot = parseXml(absent).elements?.[0];
    if (!absentRoot) throw new Error("parsed document has no root element");
    const parsedAbsent = footnotesDesc.parse(absentRoot, readCtx);
    const absentXml = footnotesDesc.stringify(parsedAbsent, writeCtx as never)!;
    expect(absentXml).not.toContain("<w:footnoteRef/>");

    const freshXml = footnotesDesc.stringify(
      { notes: new Map([[1, ["fresh note"]]]) },
      writeCtx as never,
    )!;
    expect(freshXml).toContain("<w:footnoteRef/>");

    const custom =
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="2"><w:p><w:r><w:rPr><w:b/></w:rPr><w:footnoteRef/></w:r>' +
      "<w:r><w:t>†</w:t></w:r></w:p></w:footnote></w:footnotes>";
    const customRoot = parseXml(custom).elements?.[0];
    if (!customRoot) throw new Error("parsed document has no root element");
    const parsedCustom = footnotesDesc.parse(customRoot, readCtx);
    const customXml = footnotesDesc.stringify(parsedCustom, writeCtx as never)!;
    expect(customXml.match(/<w:footnoteRef\/>/g)).toHaveLength(1);
    expect(customXml).toContain("<w:b/>");
  });
});
