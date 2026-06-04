import { describe, expect, it } from "vite-plus/test";

import { createColorTransforms } from "./color-transform";

describe("createColorTransforms", () => {
  it("should create empty array when no transforms specified", () => {
    const transforms = createColorTransforms({});
    expect(transforms).to.deep.equal([]);
  });
});
