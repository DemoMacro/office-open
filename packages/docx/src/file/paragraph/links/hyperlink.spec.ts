import { describe, expect, it } from "vite-plus/test";

import { TextRun } from "../run";
import { ExternalHyperlink } from "./hyperlink";

describe("ExternalHyperlink", () => {
  describe("#constructor()", () => {
    it("should create", () => {
      const externalHyperlink = new ExternalHyperlink({
        children: [new TextRun("test")],
        link: "http://www.google.com",
      });

      expect(externalHyperlink.options.link).to.equal("http://www.google.com");
    });
  });
});
