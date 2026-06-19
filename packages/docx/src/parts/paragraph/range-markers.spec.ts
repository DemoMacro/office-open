import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Range markers never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  return parseParagraph(doc.elements![0], readCtx) as { children?: unknown[] };
}

function childrenOf(opts: { children?: unknown[] }): Record<string, unknown>[] {
  return (opts.children ?? []) as Record<string, unknown>[];
}

describe("range markers parse", () => {
  it("parses proofErr", () => {
    const cs = childrenOf(parseParagraphXml(`<w:proofErr w:type="spellStart"/>`));
    expect(cs).toHaveLength(1);
    expect(cs[0].proofErr).toBe("spellStart");
  });

  it("parses positionalTab (w:ptab)", () => {
    const cs = childrenOf(
      parseParagraphXml(`<w:ptab w:alignment="center" w:leader="dot" w:relativeTo="margin"/>`),
    );
    expect(cs[0].positionalTab).toEqual({
      alignment: "center",
      leader: "dot",
      relativeTo: "margin",
    });
  });

  it("parses permStart with all attrs and permEnd", () => {
    const cs = childrenOf(
      parseParagraphXml(
        `<w:permStart w:id="1" w:ed="user1" w:edGrp="editors" w:colFirst="0" w:colLast="1"/>` +
          `<w:permEnd w:id="1"/>`,
      ),
    );
    expect(cs[0].permStart).toMatchObject({
      id: "1",
      ed: "user1",
      editGroup: "editors",
      colFirst: 0,
      colLast: 1,
    });
    expect(cs[1].permEnd).toBe("1");
  });

  it("parses moveFromRangeStart/End and moveToRangeStart/End", () => {
    const cs = childrenOf(
      parseParagraphXml(
        `<w:moveFromRangeStart w:id="1" w:name="move1" w:author="A" w:date="2020"/>` +
          `<w:moveFromRangeEnd w:id="1"/>` +
          `<w:moveToRangeStart w:id="2" w:name="move2"/>` +
          `<w:moveToRangeEnd w:id="2"/>`,
      ),
    );
    expect(cs[0].moveFromRangeStart).toEqual({ id: 1, name: "move1", author: "A", date: "2020" });
    expect(cs[1].moveFromRangeEnd).toBe(1);
    expect(cs[2].moveToRangeStart).toEqual({ id: 2, name: "move2" });
    expect(cs[3].moveToRangeEnd).toBe(2);
  });

  it("parses the four customXml range marker pairs", () => {
    const cs = childrenOf(
      parseParagraphXml(
        `<w:customXmlInsRangeStart w:id="1" w:author="A" w:date="2020"/><w:customXmlInsRangeEnd w:id="1"/>` +
          `<w:customXmlDelRangeStart w:id="2" w:author="B"/><w:customXmlDelRangeEnd w:id="2"/>` +
          `<w:customXmlMoveFromRangeStart w:id="3"/><w:customXmlMoveFromRangeEnd w:id="3"/>` +
          `<w:customXmlMoveToRangeStart w:id="4" w:author="D" w:date="2021"/><w:customXmlMoveToRangeEnd w:id="4"/>`,
      ),
    );
    expect(cs[0].customXmlInsRangeStart).toEqual({ id: 1, author: "A", date: "2020" });
    expect(cs[1].customXmlInsRangeEnd).toBe(1);
    expect(cs[2].customXmlDelRangeStart).toEqual({ id: 2, author: "B" });
    expect(cs[3].customXmlDelRangeEnd).toBe(2);
    expect(cs[4].customXmlMoveFromRangeStart).toEqual({ id: 3 });
    expect(cs[5].customXmlMoveFromRangeEnd).toBe(3);
    expect(cs[6].customXmlMoveToRangeStart).toEqual({ id: 4, author: "D", date: "2021" });
    expect(cs[7].customXmlMoveToRangeEnd).toBe(4);
  });

  it("drops a marker missing its required w:id", () => {
    const cs = childrenOf(parseParagraphXml(`<w:moveToRangeEnd/>`));
    expect(cs).toHaveLength(0);
  });
});

describe("move revision runs parse", () => {
  it("parses movedFrom (w:moveFrom wrapping a run)", () => {
    const cs = childrenOf(
      parseParagraphXml(
        `<w:moveFrom w:id="1" w:author="A" w:date="2020"><w:r><w:t>moved text</w:t></w:r></w:moveFrom>`,
      ),
    );
    expect(cs[0].movedFrom).toMatchObject({
      id: 1,
      author: "A",
      date: "2020",
      children: [{ text: "moved text" }],
    });
  });

  it("parses movedTo (w:moveTo wrapping a run)", () => {
    const cs = childrenOf(
      parseParagraphXml(
        `<w:moveTo w:id="2" w:author="B" w:date="2021"><w:r><w:t>target</w:t></w:r></w:moveTo>`,
      ),
    );
    expect(cs[0].movedTo).toMatchObject({
      id: 2,
      author: "B",
      date: "2021",
      children: [{ text: "target" }],
    });
  });
});
