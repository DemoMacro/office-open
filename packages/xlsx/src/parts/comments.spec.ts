import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { commentsDesc, mergeNoteAnchors, vmlNotesDesc } from "./comments";
import type { CommentsDocOptions } from "./comments";

// ── Minimal context stubs ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: CommentsDocOptions) {
  const xml = commentsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return commentsDesc.parse(el, readCtx);
}

// ── Tests ──

describe("commentsDesc round-trip", () => {
  it("empty comments list returns undefined from stringify", () => {
    const xml = commentsDesc.stringify({ comments: [] }, writeCtx);
    expect(xml).toBeUndefined();
  });

  it("round-trips basic comment with author, cell, and text", () => {
    const opts: CommentsDocOptions = {
      comments: [
        { cell: "A1", author: "Alice", text: "Hello world" },
        { cell: "B2", author: "Bob", text: "Another comment" },
      ],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments).toHaveLength(2);
    expect(comments[0]?.cell).toBe("A1");
    expect(comments[0]?.author).toBe("Alice");
    expect(comments[0]?.text).toBe("Hello world");
    expect(comments[1]?.cell).toBe("B2");
    expect(comments[1]?.author).toBe("Bob");
    expect(comments[1]?.text).toBe("Another comment");
  });

  it("round-trips multiple authors correctly", () => {
    const opts: CommentsDocOptions = {
      comments: [
        { cell: "A1", author: "Alice", text: "First" },
        { cell: "A2", author: "Bob", text: "Second" },
        { cell: "A3", author: "Alice", text: "Third" },
      ],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments).toHaveLength(3);
    // Authors are deduplicated in the XML, but parse resolves them back
    expect(comments[0]?.author).toBe("Alice");
    expect(comments[1]?.author).toBe("Bob");
    expect(comments[2]?.author).toBe("Alice");
  });

  it("round-trips special characters in text", () => {
    const opts: CommentsDocOptions = {
      comments: [{ cell: "A1", author: "Test", text: '<b>&"quotes"' }],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments[0]?.text).toBe('<b>&"quotes"');
  });

  it("round-trips single author with no text", () => {
    const opts: CommentsDocOptions = {
      comments: [{ cell: "C3", author: "Empty", text: "" }],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments).toHaveLength(1);
    expect(comments[0]?.cell).toBe("C3");
    expect(comments[0]?.author).toBe("Empty");
  });

  it("round-trips rich text runs with per-run formatting", () => {
    const opts: CommentsDocOptions = {
      comments: [
        {
          cell: "A1",
          author: "Alice",
          text: {
            runs: [
              {
                text: "bold ",
                properties: {
                  bold: true,
                  italic: false,
                  size: 12,
                  color: "FF0000",
                  font: "Calibri",
                },
              },
              {
                text: "italic",
                properties: {
                  underline: "single",
                  strike: true,
                },
              },
            ],
          },
        },
      ],
    };
    const result = roundTrip(opts);
    const text = result.comments![0]?.text;

    expect(typeof text).toBe("object");
    expect(text).not.toBeNull();
    const runs = (text as { runs: unknown[] }).runs;
    expect(runs).toHaveLength(2);
    expect((runs[0] as { text: string }).text).toBe("bold ");
    const props0 = (runs[0] as { properties: Record<string, unknown> }).properties;
    expect(props0.bold).toBe(true);
    expect(props0.size).toBe(12);
    expect(props0.color).toBe("FF0000");
    expect(props0.font).toBe("Calibri");
    const props1 = (runs[1] as { properties: Record<string, unknown> }).properties;
    expect(props1.underline).toBe("single");
    expect(props1.strike).toBe(true);
  });

  it("parses commentPr but never re-emits it", () => {
    // stringify drops commentPr: Excel refuses to open it beside the VML note
    // drawing the compiler always writes. Parse keeps the fields for inspection.
    const emitted = commentsDesc.stringify(
      {
        comments: [
          {
            cell: "A1",
            author: "Alice",
            text: "note",
            commentPr: { locked: false, print: false, textHAlign: "center" },
          },
        ],
      },
      writeCtx,
    )!;
    expect(emitted).not.toContain("commentPr");

    const xml =
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ` xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">` +
      `<authors><author>Alice</author></authors><commentList>` +
      `<comment ref="A1" authorId="0"><text><t>note</t></text>` +
      `<commentPr locked="0" print="0" textHAlign="center">` +
      `<anchor moveWithCells="1" sizeWithCells="0">` +
      `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>` +
      `</anchor></commentPr></comment></commentList></comments>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = commentsDesc.parse(el, readCtx);
    const pr = result.comments[0]!.commentPr!;
    expect(pr.locked).toBe(false);
    expect(pr.print).toBe(false);
    expect(pr.textHAlign).toBe("center");
    expect(pr.anchor?.moveWithCells).toBe(true);
    expect(pr.anchor?.sizeWithCells).toBe(false);
    // parseMarker normalizes omitted offsets to 0
    expect(pr.anchor?.from).toEqual({ col: 1, row: 1, colOff: 0, rowOff: 0 });
    expect(pr.anchor?.to).toEqual({ col: 3, row: 4, colOff: 0, rowOff: 0 });
  });
});

describe("vmlNotesDesc stringify multi-column cell refs", () => {
  it("anchors AA1 and AB10 with correct 0-based column/row", () => {
    const xml = vmlNotesDesc.stringify(
      {
        comments: [
          { cell: "AA1", author: "A", text: "x" },
          { cell: "AB10", author: "B", text: "y" },
        ],
      },
      writeCtx,
    )!;
    // AA = col 26, AB = col 27 (0-based); rows are 1-based in refs → 0-based
    expect(xml).toContain("<x:Column>26</x:Column>");
    expect(xml).toContain("<x:Row>0</x:Row>");
    expect(xml).toContain("<x:Column>27</x:Column>");
    expect(xml).toContain("<x:Row>9</x:Row>");
  });
});

describe("vmlNotesDesc parse round-trips note placement", () => {
  function roundTripVml(opts: CommentsDocOptions) {
    const xml = vmlNotesDesc.stringify(opts, writeCtx)!;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    return vmlNotesDesc.parse(el, readCtx);
  }

  it("round-trips default anchor, hidden state, and default size", () => {
    const anchors = roundTripVml({ comments: [{ cell: "B3", author: "A", text: "x" }] });
    expect(anchors).toEqual([
      {
        row: 2,
        column: 1,
        anchor: [1, 0, 2, 0, 3, 0, 4, 0],
        visible: false,
        width: 108,
        height: 59.25,
      },
    ]);
  });

  it("round-trips custom anchor, visible state, and size", () => {
    const anchors = roundTripVml({
      comments: [
        {
          cell: "A1",
          author: "A",
          text: "x",
          anchor: [0, 15, 0, 12, 2, 32, 4, 4],
          visible: true,
          size: { width: 200, height: 100 },
        },
      ],
    });
    expect(anchors[0]?.anchor).toEqual([0, 15, 0, 12, 2, 32, 4, 4]);
    expect(anchors[0]?.visible).toBe(true);
    expect(anchors[0]?.width).toBe(200);
    expect(anchors[0]?.height).toBe(100);
  });

  it("parses a source-style vmlDrawing with newline-separated anchor", () => {
    const src =
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;width:108pt;height:59.25pt;visibility:visible">' +
      '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
      "<x:Anchor>2, 15, 1, 12, 4, 32, 0, 4</x:Anchor>" +
      "<x:Row>1</x:Row><x:Column>2</x:Column>" +
      "</x:ClientData>" +
      "</v:shape></xml>";
    const el = parseXml(src).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const anchors = vmlNotesDesc.parse(el, readCtx);
    expect(anchors).toEqual([
      {
        row: 1,
        column: 2,
        anchor: [2, 15, 1, 12, 4, 32, 0, 4],
        visible: true,
        width: 108,
        height: 59.25,
      },
    ]);
  });
});

describe("mergeNoteAnchors", () => {
  it("merges placement into the matching comment by cell", () => {
    const comments: CommentsDocOptions["comments"] = [
      { cell: "A1", author: "A", text: "x" },
      { cell: "B2", author: "B", text: "y" },
    ];
    mergeNoteAnchors(comments, [
      {
        row: 1,
        column: 1,
        anchor: [1, 0, 1, 0, 3, 0, 3, 0],
        visible: true,
        width: 200,
        height: 90,
      },
    ]);
    expect(comments[0]).toEqual({ cell: "A1", author: "A", text: "x" });
    expect(comments[1]?.anchor).toEqual([1, 0, 1, 0, 3, 0, 3, 0]);
    expect(comments[1]?.visible).toBe(true);
    expect(comments[1]?.size).toEqual({ width: 200, height: 90 });
  });
});
