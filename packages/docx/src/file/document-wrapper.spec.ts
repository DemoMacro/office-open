import { describe, expect, it } from "vite-plus/test";

import { DocumentWrapper } from "./document-wrapper";

describe("DocumentWrapper", () => {
  describe("#constructor", () => {
    it("should create", () => {
      const file = new DocumentWrapper({ background: {} });

      expect(file.view).to.be.ok;
      expect(file.relationships).to.be.ok;
    });
  });
});
