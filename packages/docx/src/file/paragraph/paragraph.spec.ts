import { assert, describe, it } from "vite-plus/test";

import { Paragraph } from "./paragraph";

describe("Paragraph", () => {
  describe("#constructor()", () => {
    it("should create valid JSON", () => {
      const paragraph = new Paragraph("");
      const stringifiedJson = JSON.stringify(paragraph);

      try {
        JSON.parse(stringifiedJson);
      } catch {
        assert.isTrue(false);
      }
      assert.isTrue(true);
    });
  });
});
