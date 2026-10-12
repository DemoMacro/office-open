import { describe, expect, it } from "vite-plus/test";

import { convertOutput, type OutputType } from "./output";

describe("convertOutput", () => {
  it("converts binary data to a numeric array", () => {
    expect(convertOutput(new Uint8Array([1, 2, 3]), "array")).toEqual([1, 2, 3]);
  });

  it("fails loudly when a new output type is not dispatched", () => {
    const unhandledType = "future-output" as OutputType;
    expect(() => convertOutput(new Uint8Array([1]), unhandledType)).toThrow(TypeError);
  });
});
