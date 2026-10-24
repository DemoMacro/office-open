import { describe, expect, it } from "vite-plus/test";

import { sanitizeOptionsText } from "./text";

describe("sanitizeOptionsText", () => {
  it("normalizes strings recursively while preserving XML whitespace", () => {
    expect(sanitizeOptionsText({ children: ["a\tb\u000Bc", { text: "d\u0000e" }] })).toEqual({
      children: ["a\tb\nc", { text: "de" }],
    });
  });

  it("preserves dates and binary payloads", () => {
    const date = new Date(0);
    const data = new Uint8Array([1, 2, 3]);
    const result = sanitizeOptionsText({ date, data });
    expect(result.date).toBe(date);
    expect(result.data).toBe(data);
  });
});
