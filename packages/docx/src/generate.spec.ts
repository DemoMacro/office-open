import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "./generate";

describe("generateDocument entry guards", () => {
  it("names the missing sections array instead of dying in the compiler", () => {
    expect(() => generateDocument({} as never)).toThrow(/sections is required/);
  });
});
