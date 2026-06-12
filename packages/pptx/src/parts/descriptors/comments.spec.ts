import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { AuthorEntry, CommentEntry } from "@parts/comment";
import { describe, expect, it } from "vite-plus/test";

import { commentAuthorsDesc, slideCommentsDesc } from "./comments";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

// ── commentAuthorsDesc ──

describe("commentAuthorsDesc round-trip", () => {
  function roundTrip(authors: AuthorEntry[]) {
    const xml = commentAuthorsDesc.stringify(authors, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    return commentAuthorsDesc.parse(el, readCtx) as AuthorEntry[];
  }

  it("round-trips single author", () => {
    const authors: AuthorEntry[] = [{ id: 0, name: "Alice", initials: "A", clrIdx: 0, lastIdx: 0 }];
    const result = roundTrip(authors);

    expect(result).toHaveLength(1);
    expect(result[0].id).toBe(0);
    expect(result[0].name).toBe("Alice");
    expect(result[0].initials).toBe("A");
    expect(result[0].clrIdx).toBe(0);
    expect(result[0].lastIdx).toBe(0);
  });

  it("round-trips multiple authors", () => {
    const authors: AuthorEntry[] = [
      { id: 0, name: "Alice", initials: "AL", clrIdx: 0, lastIdx: 1 },
      { id: 1, name: "Bob", initials: "BO", clrIdx: 1, lastIdx: 2 },
    ];
    const result = roundTrip(authors);

    expect(result).toHaveLength(2);
    expect(result[0].name).toBe("Alice");
    expect(result[1].name).toBe("Bob");
    expect(result[1].id).toBe(1);
    expect(result[1].clrIdx).toBe(1);
  });
});

// ── slideCommentsDesc ──

describe("slideCommentsDesc round-trip", () => {
  function roundTrip(comments: CommentEntry[]) {
    const xml = slideCommentsDesc.stringify(comments, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    return slideCommentsDesc.parse(el, readCtx) as CommentEntry[];
  }

  it("round-trips basic comment", () => {
    const comments: CommentEntry[] = [{ authorId: 0, idx: 1, x: 100, y: 200, text: "Hello" }];
    const result = roundTrip(comments);

    expect(result).toHaveLength(1);
    expect(result[0].authorId).toBe(0);
    expect(result[0].idx).toBe(1);
    expect(result[0].x).toBe(100);
    expect(result[0].y).toBe(200);
    expect(result[0].text).toBe("Hello");
  });

  it("round-trips comment with date", () => {
    const comments: CommentEntry[] = [
      { authorId: 0, idx: 1, date: "2024-01-15T10:30:00Z", x: 0, y: 0, text: "Dated" },
    ];
    const result = roundTrip(comments);

    expect(result[0].date).toBe("2024-01-15T10:30:00Z");
  });

  it("round-trips comment with modified flag", () => {
    const comments: CommentEntry[] = [
      { authorId: 0, idx: 1, modified: true, x: 0, y: 0, text: "Modified" },
    ];
    const result = roundTrip(comments);

    expect(result[0].modified).toBe(true);
  });

  it("round-trips multiple comments", () => {
    const comments: CommentEntry[] = [
      { authorId: 0, idx: 1, x: 100, y: 100, text: "First" },
      { authorId: 1, idx: 2, x: 200, y: 200, text: "Second" },
    ];
    const result = roundTrip(comments);

    expect(result).toHaveLength(2);
    expect(result[1].authorId).toBe(1);
    expect(result[1].text).toBe("Second");
  });

  it("round-trips special characters in text", () => {
    const comments: CommentEntry[] = [
      { authorId: 0, idx: 1, x: 0, y: 0, text: '<Tag> & "quotes"' },
    ];
    const result = roundTrip(comments);

    expect(result[0].text).toBe('<Tag> & "quotes"');
  });
});
