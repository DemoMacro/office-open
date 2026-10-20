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
    expect(xml).toContain("<w:t>cell</w:t>");
    expect(xml).not.toContain("<w:endnoteRef/>");
  });

  it("injects the reference run after the paragraph properties", () => {
    // CT_P orders w:pPr before every run: an endnoteRef injected right after
    // the paragraph open tag makes the part schema-invalid.
    const source =
      '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
      "<w:r><w:t>note text</w:t></w:r></w:p></w:endnote></w:endnotes>";
    const root = parseXml(source).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const parsed = endnotesDesc.parse(root, readCtx);
    parsed.autoReferenceMarks?.set(1, true);

    const xml = endnotesDesc.stringify(parsed, writeCtx as any)!;
    expect(xml).toContain(
      '</w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>',
    );
    expect(xml.indexOf("<w:endnoteRef/>")).toBeGreaterThan(xml.indexOf("</w:pPr>"));
  });

  it("preserves absent and custom endnote reference marks", () => {
    const absent =
      '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:endnote w:id="1"><w:p><w:r><w:t>custom note</w:t></w:r></w:p></w:endnote>' +
      "</w:endnotes>";
    const absentRoot = parseXml(absent).elements?.[0];
    if (!absentRoot) throw new Error("parsed document has no root element");
    const parsedAbsent = endnotesDesc.parse(absentRoot, readCtx);
    const absentXml = endnotesDesc.stringify(parsedAbsent, writeCtx as never)!;
    expect(absentXml).not.toContain("<w:endnoteRef/>");

    const freshXml = endnotesDesc.stringify(
      { notes: new Map([[1, ["fresh note"]]]) },
      writeCtx as never,
    )!;
    expect(freshXml).toContain("<w:endnoteRef/>");

    const custom =
      '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:endnote w:id="2"><w:p><w:r><w:rPr><w:b/></w:rPr><w:endnoteRef/></w:r>' +
      "<w:r><w:t>§</w:t></w:r></w:p></w:endnote></w:endnotes>";
    const customRoot = parseXml(custom).elements?.[0];
    if (!customRoot) throw new Error("parsed document has no root element");
    const parsedCustom = endnotesDesc.parse(customRoot, readCtx);
    const customXml = endnotesDesc.stringify(parsedCustom, writeCtx as never)!;
    expect(customXml.match(/<w:endnoteRef\/>/g)).toHaveLength(1);
    expect(customXml).toContain("<w:b/>");
  });
});
