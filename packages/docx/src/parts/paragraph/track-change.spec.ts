import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Track changes never touch the contexts, so empty mocks suffice.
const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function roundTrip(inner: string): string {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return stringifyParagraph(parseParagraph(el, readCtx), writeCtx);
}

describe("fresh revision markers without an id", () => {
  it("assigns distinct document-unique ids on emit", () => {
    // Fresh input may omit the marker id (nothing references it): the library
    // allocates one, and two markers in one document never share a value.
    const out1 = stringifyParagraph(
      {
        children: [
          {
            insertion: { author: "Alice", date: "2026-01-01T00:00:00Z", children: [{ text: "a" }] },
          },
        ],
      } as never,
      writeCtx,
    );
    const out2 = stringifyParagraph(
      {
        children: [
          {
            deletion: { author: "Alice", date: "2026-01-01T00:00:00Z", children: [{ text: "b" }] },
          },
        ],
      } as never,
      writeCtx,
    );
    const idOf = (xml: string) => xml.match(/w:id="(\d+)"/)![1];
    const [a, b] = [idOf(out1), idOf(out2)];
    expect(a).not.toBe("undefined");
    expect(b).not.toBe("undefined");
    expect(a).not.toBe(b);
  });

  it("emits a numeric pPrChange id when the revision omits it", () => {
    const out = stringifyParagraph(
      {
        children: [{ text: "x" }],
        revision: { author: "Alice", date: "2026-01-01T00:00:00Z", alignment: "center" },
      } as never,
      writeCtx,
    );
    expect(out).toMatch(/<w:pPrChange [^>]*w:id="\d+"/);
  });
});

describe("track-change round-trip", () => {
  it("round-trips w14:noSpellErr on the paragraph", () => {
    const doc = parseXml(`<w:p ${W_NS} w14:noSpellErr="1"><w:r><w:t>text</w:t></w:r></w:p>`);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const xml = stringifyParagraph(parseParagraph(el, readCtx), writeCtx);
    expect(xml).toContain('w14:noSpellErr="1"');
  });

  it("round-trips w14 conflictIns and conflictDel run wrappers", () => {
    const xml = roundTrip(
      '<w14:conflictIns w:id="7" w:author="Bob" w:date="2021-02-03T00:00:00Z">' +
        "<w:r><w:t>kept</w:t></w:r>" +
        "</w14:conflictIns>" +
        '<w14:conflictDel w:id="8" w:author="Bob" w:date="2021-02-03T00:00:00Z">' +
        "<w:r><w:delText>dropped</w:delText></w:r>" +
        "</w14:conflictDel>",
    );
    expect(xml).toContain(
      '<w14:conflictIns w:id="7" w:author="Bob" w:date="2021-02-03T00:00:00Z">',
    );
    expect(xml).toContain("kept");
    expect(xml).toContain(
      '<w14:conflictDel w:id="8" w:author="Bob" w:date="2021-02-03T00:00:00Z">',
    );
    expect(xml).toContain("dropped");
  });

  it("round-trips w14 customXml conflict range markers", () => {
    const xml = roundTrip(
      '<w14:customXmlConflictInsRangeStart w:id="9" w:author="Cara" w:date="2021-03-04T00:00:00Z"/>' +
        "<w:r><w:t>body</w:t></w:r>" +
        '<w14:customXmlConflictInsRangeEnd w:id="9"/>' +
        '<w14:customXmlConflictDelRangeStart w:id="10"/>' +
        '<w14:customXmlConflictDelRangeEnd w:id="10"/>',
    );
    expect(xml).toContain(
      '<w14:customXmlConflictInsRangeStart w:id="9" w:author="Cara" w:date="2021-03-04T00:00:00Z"/>',
    );
    expect(xml).toContain('<w14:customXmlConflictInsRangeEnd w:id="9"/>');
    expect(xml).toContain('<w14:customXmlConflictDelRangeStart w:id="10"/>');
    expect(xml).toContain('<w14:customXmlConflictDelRangeEnd w:id="10"/>');
  });

  it("re-emits a plain deleted text run as w:delText", () => {
    // The dispatch text fast path emits w:t; a run inside w:del must keep the
    // delText spelling regardless of how the run shape is serialized.
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">gone</w:delText></w:r>' +
        "</w:del>",
    );
    expect(xml).toContain('<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">');
    expect(xml).toContain("<w:rPr><w:b/></w:rPr>");
    expect(xml).toContain('<w:delText xml:space="preserve">gone</w:delText>');
    expect(xml).not.toContain("<w:t");
  });

  it("keeps w:t for the same run shape inside w:ins", () => {
    const xml = roundTrip(
      '<w:ins w:id="2" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:delText xml:space="preserve">stray</w:delText></w:r>' +
        "</w:ins>",
    );
    expect(xml).toContain("<w:ins ");
    expect(xml).toContain("<w:t");
    expect(xml).not.toContain("delText");
  });

  it("keeps run identity attributes on tracked comment references", () => {
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD" w:rsidDel="DDEEFF00">' +
        '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
        '<w:commentReference w:id="7"/></w:r></w:del>',
    );
    expect(xml).toContain('<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD" w:rsidDel="DDEEFF00">');
  });
});
