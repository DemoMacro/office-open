import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import { commentsDesc } from "./comments";
import type { CommentsOptions } from "./paragraph/run/comment-run";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) => String(child),
  fileData: {} as never,
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: CommentsOptions): CommentsOptions {
  const xml = commentsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return commentsDesc.parse(el, readCtx);
}

describe("commentsDesc round-trip", () => {
  it("round-trips comment with author and date", () => {
    const result = roundTrip({
      children: [{ id: 1, author: "John", date: "2024-01-15T10:30:00Z", children: [] }],
    });
    expect(result.children).toHaveLength(1);
    expect(result.children[0].id).toBe(1);
    expect(result.children[0].author).toBe("John");
    expect(result.children[0].date).toBe("2024-01-15T10:30:00Z");
  });

  it("round-trips comment with initials", () => {
    const result = roundTrip({
      children: [
        { id: 2, author: "Jane", initials: "JD", date: "2024-02-01T12:00:00Z", children: [] },
      ],
    });
    expect(result.children[0].initials).toBe("JD");
  });

  it("round-trips multiple comments", () => {
    const result = roundTrip({
      children: [
        { id: 1, author: "A", date: "2024-01-01T00:00:00Z", children: [] },
        { id: 2, author: "B", date: "2024-02-01T00:00:00Z", children: [] },
      ],
    });
    expect(result.children).toHaveLength(2);
    expect(result.children[0].author).toBe("A");
    expect(result.children[1].author).toBe("B");
  });

  it("round-trips empty comments", () => {
    const result = roundTrip({ children: [] });
    expect(result.children).toHaveLength(0);
  });
});
