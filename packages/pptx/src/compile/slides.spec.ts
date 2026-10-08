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
});
