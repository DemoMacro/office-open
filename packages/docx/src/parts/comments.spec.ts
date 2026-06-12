import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import { commentsDesc } from "./comments";

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

function roundTrip(opts: { children: Record<string, unknown>[] }) {
  const xml = commentsDesc.stringify(opts as any, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return commentsDesc.parse(el, readCtx) as unknown as Record<string, unknown>;
}

describe("commentsDesc round-trip", () => {
  it("round-trips comment with author and date", () => {
    const result = roundTrip({
      children: [{ id: 1, author: "John", date: "2024-01-15T10:30:00Z", children: [] }],
    });
    const children = result.children as Record<string, unknown>[];
    expect(children).toHaveLength(1);
    expect(children[0].id).toBe(1);
    expect(children[0].author).toBe("John");
    expect(children[0].date).toBe("2024-01-15T10:30:00Z");
  });

  it("round-trips comment with initials", () => {
    const result = roundTrip({
      children: [
        { id: 2, author: "Jane", initials: "JD", date: "2024-02-01T12:00:00Z", children: [] },
      ],
    });
    const children = result.children as Record<string, unknown>[];
    expect(children[0].initials).toBe("JD");
  });

  it("round-trips multiple comments", () => {
    const result = roundTrip({
      children: [
        { id: 1, author: "A", date: "2024-01-01T00:00:00Z", children: [] },
        { id: 2, author: "B", date: "2024-02-01T00:00:00Z", children: [] },
      ],
    });
    const children = result.children as Record<string, unknown>[];
    expect(children).toHaveLength(2);
    expect(children[0].author).toBe("A");
    expect(children[1].author).toBe("B");
  });

  it("round-trips empty comments", () => {
    const result = roundTrip({ children: [] });
    const children = result.children as Record<string, unknown>[];
    expect(children).toHaveLength(0);
  });
});
