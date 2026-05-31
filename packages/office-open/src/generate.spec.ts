import { describe, expect, it } from "vite-plus/test";

import { parseInput } from "./generate";

describe("parseInput", () => {
  it("should parse a JSON string", async () => {
    const result = await parseInput('{"sections": []}');
    expect(result).toEqual({ sections: [] });
  });

  it("should parse a JSON string with nested objects", async () => {
    const input = JSON.stringify({ worksheets: [{ rows: [{ cells: [{ value: 1 }] }] }] });
    const result = await parseInput(input);
    expect(result).toEqual({ worksheets: [{ rows: [{ cells: [{ value: 1 }] }] }] });
  });

  it("should parse a JSON string starting with [", async () => {
    const input = JSON.stringify([{ foo: "bar" }]);
    const result = await parseInput(input);
    expect(result).toEqual([{ foo: "bar" }]);
  });

  it("should throw on invalid JSON string", async () => {
    await expect(parseInput("not json")).rejects.toThrow();
  });

  it("should trim whitespace before parsing", async () => {
    const result = await parseInput('  {"sections": []}  ');
    expect(result).toEqual({ sections: [] });
  });

  it("should handle empty object", async () => {
    const result = await parseInput("{}");
    expect(result).toEqual({});
  });

  it("should handle empty array", async () => {
    const result = await parseInput("[]");
    expect(result).toEqual([]);
  });
});
