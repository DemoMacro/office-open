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

function movedFromChildren(inner: string) {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  const paragraph = parseParagraph(el, readCtx);
  const child = paragraph.children?.find(
    (candidate) => typeof candidate === "object" && "movedFrom" in candidate,
  );
  if (!child || !("movedFrom" in child)) throw new Error("parsed paragraph has no movedFrom");
  return child.movedFrom.children;
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
  it("round-trips hyperlinks and nested insertions in both nesting directions", () => {
    const xml = roundTrip(
      '<w:ins w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        "<w:r><w:t>before</w:t></w:r>" +
        '<w:hyperlink w:anchor="section" w:history="1">' +
        "<w:r><w:t>link</w:t></w:r>" +
        '<w:ins w:id="2" w:author="Bob" w:date="2020-01-02T00:00:00Z">' +
        "<w:r><w:t>inserted</w:t></w:r></w:ins>" +
        "</w:hyperlink></w:ins>",
    );

    expect(xml.indexOf("<w:r><w:t>before</w:t></w:r>")).toBeLessThan(
      xml.indexOf('<w:hyperlink w:history="1" w:anchor="section">'),
    );
    expect(xml.indexOf("<w:r><w:t>link</w:t></w:r>")).toBeLessThan(
      xml.indexOf('<w:ins w:id="2" w:author="Bob" w:date="2020-01-02T00:00:00Z">'),
    );
    expect(xml).toContain("<w:r><w:t>inserted</w:t></w:r></w:ins></w:hyperlink></w:ins>");
  });

  it("preserves per-stage rsids on a deleted complex field", () => {
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r w:rsidR="00000001" w:rsidRPr="00000002" w:rsidDel="00000003">' +
        '<w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r w:rsidR="00000004" w:rsidRPr="00000005" w:rsidDel="00000006">' +
        '<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>' +
        '<w:r w:rsidR="00000007" w:rsidRPr="00000008" w:rsidDel="00000009">' +
        '<w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r w:rsidR="0000000A" w:rsidRPr="0000000B" w:rsidDel="0000000C">' +
        "<w:delText>1</w:delText></w:r>" +
        '<w:r w:rsidR="0000000D" w:rsidRPr="0000000E" w:rsidDel="0000000F">' +
        '<w:fldChar w:fldCharType="end"/></w:r></w:del>',
    );
    for (const id of ["00000001", "00000004", "00000007", "0000000A", "0000000D"]) {
      expect(xml).toContain(`w:rsidR="${id}"`);
    }
    for (const id of ["00000003", "00000006", "00000009", "0000000C", "0000000F"]) {
      expect(xml).toContain(`w:rsidDel="${id}"`);
    }
  });

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

  it("round-trips mixed content inside a deleted run", () => {
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        "<w:r><w:rPr><w:b/></w:rPr><w:delText>gone</w:delText><w:tab/>" +
        '<w:footnoteReference w:id="4"/><w:lastRenderedPageBreak/></w:r>' +
        "</w:del>",
    );
    expect(xml).toContain(
      "<w:r><w:rPr><w:b/></w:rPr><w:delText>gone</w:delText><w:tab/>" +
        '<w:footnoteReference w:id="4"/><w:lastRenderedPageBreak/></w:r>',
    );
    expect(xml).not.toContain("<w:t>");
  });

  it("preserves explicit deleted text inside w:ins", () => {
    const xml = roundTrip(
      '<w:ins w:id="2" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:delText xml:space="preserve">stray</w:delText></w:r>' +
        "</w:ins>",
    );
    expect(xml).toContain("<w:ins ");
    expect(xml).toContain('<w:delText xml:space="preserve">stray</w:delText>');
  });

  it("keeps bookmark ranges anchored inside a deletion", () => {
    const xml = roundTrip(
      '<w:del w:id="3" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:bookmarkStart w:id="167" w:name="Deleted"/>' +
        "<w:r><w:delText>gone</w:delText></w:r>" +
        '<w:bookmarkEnd w:id="167"/>' +
        "</w:del>",
    );
    expect(xml).toContain('<w:bookmarkStart w:id="167" w:name="Deleted"/>');
    expect(xml).toContain('<w:bookmarkEnd w:id="167"/>');
  });

  it("keeps moved-from run text as w:delText", () => {
    const xml = roundTrip(
      '<w:moveFrom w:id="4" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:delText xml:space="preserve">moved away</w:delText></w:r>' +
        "</w:moveFrom>",
    );
    expect(xml).toContain('<w:moveFrom w:id="4" w:author="Alice" w:date="2020-01-01T00:00:00Z">');
    expect(xml).toContain('<w:delText xml:space="preserve">moved away</w:delText>');
    expect(xml).not.toContain(">moved away</w:t>");
  });

  it("keeps moved-from plain run text as w:t", () => {
    const xml = roundTrip(
      '<w:moveFrom w:id="5" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:t xml:space="preserve">moved text</w:t></w:r>' +
        "</w:moveFrom>",
    );
    expect(xml).toContain('<w:moveFrom w:id="5" w:author="Alice" w:date="2020-01-01T00:00:00Z">');
    expect(xml).toContain('<w:t xml:space="preserve">moved text</w:t>');
    expect(xml).not.toContain("<w:delText");
  });

  it("preserves an unterminated nested field inside a move revision", () => {
    const xml = roundTrip(
      '<w:moveFrom w:id="6" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> TOC </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        "</w:moveFrom>",
    );
    expect(xml).toContain('<w:moveFrom w:id="6"');
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(2);
    expect(xml).toContain("<w:instrText> TOC </w:instrText>");
    expect(xml).toContain("<w:instrText> PAGE </w:instrText>");
    expect(xml).toContain("<w:t>1</w:t>");
  });

  it("keeps a completed nested field when an outer field is unterminated", () => {
    const children = movedFromChildren(
      '<w:moveFrom w:id="7" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> TOC </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGEREF _Toc1 </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:moveFrom>",
    );

    expect(children).toHaveLength(2);
    const header = children.find((child) => typeof child === "object" && "rawXml" in child);
    const nestedField = children.find(
      (child) => typeof child === "object" && "complexField" in child,
    );
    expect(header && "rawXml" in header && header.rawXml).toContain('w:fldCharType="begin"');
    expect(nestedField && "complexField" in nestedField).toBe(true);
    if (nestedField && "complexField" in nestedField) {
      expect(nestedField.complexField.instruction).toBe(" PAGEREF _Toc1 ");
      expect(nestedField.complexField.result).toBe("1");
    }

    const xml = roundTrip(
      '<w:moveFrom w:id="7" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> TOC </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        "<w:r><w:instrText> PAGEREF _Toc1 </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>1</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:moveFrom>",
    );
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1);
    expect(xml).toContain("<w:instrText> PAGEREF _Toc1 </w:instrText>");
    expect(xml).toContain("<w:t>1</w:t>");
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

  it("keeps every comment reference in a tracked run", () => {
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r w:rsidR="00112233">' +
        '<w:commentReference w:id="7"/><w:commentReference w:id="9"/>' +
        "</w:r></w:del>",
    );
    expect(xml).toContain('<w:commentReference w:id="7"/>');
    expect(xml).toContain('<w:commentReference w:id="9"/>');
    expect(xml).toContain('<w:r w:rsidR="00112233">');
  });

  it("does not inject a comment-reference style into deleted mixed runs", () => {
    const xml = roundTrip(
      '<w:del w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z">' +
        '<w:r><w:commentReference w:id="7"/></w:r></w:del>',
    );
    expect(xml).toContain('<w:commentReference w:id="7"/>');
    expect(xml).not.toContain('w:val="CommentReference"');
  });
});
