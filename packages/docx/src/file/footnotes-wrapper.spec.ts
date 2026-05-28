import { describe, expect, it } from "vite-plus/test";

import { FootnotesWrapper } from "./footnotes-wrapper";

describe("FootnotesWrapper", () => {
  describe("#constructor", () => {
    it("should create", () => {
      const file = new FootnotesWrapper();

      expect(file.view).to.be.ok;
      expect(file.relationships).to.be.ok;
    });
  });
});
