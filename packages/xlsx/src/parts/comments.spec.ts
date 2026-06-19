import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { commentsDesc, vmlNotesDesc } from "./comments";
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
  const el = doc.elements![0];
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
    expect(comments[0].cell).toBe("A1");
    expect(comments[0].author).toBe("Alice");
    expect(comments[0].text).toBe("Hello world");
    expect(comments[1].cell).toBe("B2");
    expect(comments[1].author).toBe("Bob");
    expect(comments[1].text).toBe("Another comment");
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
    expect(comments[0].author).toBe("Alice");
    expect(comments[1].author).toBe("Bob");
    expect(comments[2].author).toBe("Alice");
  });

  it("round-trips special characters in text", () => {
    const opts: CommentsDocOptions = {
      comments: [{ cell: "A1", author: "Test", text: '<b>&"quotes"' }],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments[0].text).toBe('<b>&"quotes"');
  });

  it("round-trips single author with no text", () => {
    const opts: CommentsDocOptions = {
      comments: [{ cell: "C3", author: "Empty", text: "" }],
    };
    const result = roundTrip(opts);
    const comments = result.comments!;

    expect(comments).toHaveLength(1);
    expect(comments[0].cell).toBe("C3");
    expect(comments[0].author).toBe("Empty");
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
    const text = result.comments![0].text;

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

  it("round-trips commentPr with properties and anchor", () => {
    const opts: CommentsDocOptions = {
      comments: [
        {
          cell: "A1",
          author: "Alice",
          text: "note",
          commentPr: {
            locked: false,
            print: false,
            textHAlign: "center",
            anchor: { moveWithCells: true, sizeWithCells: false },
          },
        },
      ],
    };
    const result = roundTrip(opts);
    const pr = result.comments![0].commentPr!;
    expect(pr.locked).toBe(false);
    expect(pr.print).toBe(false);
    expect(pr.textHAlign).toBe("center");
    expect(pr.anchor?.moveWithCells).toBe(true);
    expect(pr.anchor?.sizeWithCells).toBe(false);
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
