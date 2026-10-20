import { describe, expect, it } from "vite-plus/test";

import { buildCommentData } from "./slides";

describe("buildCommentData", () => {
  it("preserves a source lastIdx above existing comment ordinals", () => {
    const { authors } = buildCommentData(
      [{ comments: [{ author: "Alice", idx: 2, x: 0, y: 0, date: "2026-01-01", text: "x" }] }],
      [{ id: 0, name: "Alice", initials: "A", clrIdx: 0, lastIdx: 7 }],
    );
    expect(authors?.[0]?.lastIdx).toBe(7);
  });

  it("does not reset a captured author when comments arrive later", () => {
    const { authors } = buildCommentData(
      [
        {
          comments: [
            { author: "Alice", authorId: 0, idx: 1, x: 0, y: 0, date: "2026-01-01", text: "a" },
          ],
        },
        {
          comments: [
            { author: "Alice", authorId: 0, idx: 2, x: 0, y: 0, date: "2026-01-02", text: "b" },
          ],
        },
      ],
      [{ id: 0, name: "Alice", initials: "A", clrIdx: 0, lastIdx: 9 }],
    );
    expect(authors?.[0]?.lastIdx).toBe(9);
  });

  it("preserves duplicate author names as distinct author ids", () => {
    const { authors } = buildCommentData(
      [
        {
          comments: [
            { author: "Same", authorId: 1, idx: 1, x: 0, y: 0, text: "first" },
            { author: "Same", authorId: 5, idx: 1, x: 0, y: 0, text: "second" },
          ],
        },
      ],
      [
        { id: 1, name: "Same", initials: "S1", clrIdx: 1, lastIdx: 1 },
        { id: 5, name: "Same", initials: "S5", clrIdx: 5, lastIdx: 1 },
      ],
    );
    expect(authors?.map((author) => author.id)).toEqual([1, 5]);
  });
});
