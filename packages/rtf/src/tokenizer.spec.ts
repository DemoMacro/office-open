import { describe, expect, it } from "vite-plus/test";

import { tokenizeRtf } from "./tokenizer";

describe("tokenizeRtf", () => {
  it("separates groups, control words, parameters, and text", () => {
    const tokens = tokenizeRtf(String.raw`{\rtf1\fs24 Hello}`);
    expect(tokens[0]).toEqual({ kind: "group-start", position: 0 });
    expect(tokens[1]).toEqual({ kind: "control", word: "rtf", param: 1, position: 1 });
    expect(tokens[2]).toEqual({ kind: "control", word: "fs", param: 24, position: 6 });
    expect(tokens[3]).toMatchObject({ kind: "text", value: "Hello", position: 12 });
    expect(tokens.at(-1)).toEqual({ kind: "group-end", position: 17 });
  });

  it("keeps control symbols and escaped delimiters as controls", () => {
    const tokens = tokenizeRtf(String.raw`\{\}\\`);
    expect(tokens.map((token) => token.kind)).toEqual(["control", "control", "control"]);
    expect(tokens.map((token) => token.kind === "control" && token.symbol)).toEqual([
      "{",
      "}",
      "\\",
    ]);
  });

  it("reads hexadecimal escapes", () => {
    expect(tokenizeRtf("\\'41")).toEqual([{ kind: "hex", value: "41", position: 0 }]);
  });

  it("reports malformed escapes with positions", () => {
    expect(() => tokenizeRtf("\\'4")).toThrow("requires two hexadecimal digits");
    expect(() => tokenizeRtf("\\")).toThrow("ends at end of input");
  });

  it("treats a parameterless binary control as empty", () => {
    const tokens = tokenizeRtf(String.raw`{\rtf1{\pict\bin hex}}`);
    expect(tokens.some((token) => token.kind === "text" && token.value === "hex")).toBe(true);
  });

  it("materializes binary payload bytes and reports invalid lengths", () => {
    expect(tokenizeRtf(String.raw`\bin4 ABCD`)).toEqual([
      { kind: "binary", byteLength: 4, data: new Uint8Array([65, 66, 67, 68]), position: 0 },
    ]);
    expect(tokenizeRtf(String.raw`\bin0`)).toEqual([
      { kind: "binary", byteLength: 0, data: new Uint8Array(0), position: 0 },
    ]);
    expect(() => tokenizeRtf("\\bin-1")).toThrow("non-negative byte count");
    expect(() => tokenizeRtf("\\bin4 AB")).toThrow("extends past end");
  });
});
