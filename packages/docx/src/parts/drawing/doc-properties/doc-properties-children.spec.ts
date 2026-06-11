import { describe, expect, it } from "vite-plus/test";

import { buildHyperlinkClickObj, buildHyperlinkHoverObj } from "./doc-properties-children";

describe("Document Properties Children", () => {
  describe("#buildHyperlinkClickObj", () => {
    it("should create a Hyperlink Click object", () => {
      expect(buildHyperlinkClickObj("1", false)).to.deep.equal({
        "a:hlinkClick": {
          _attr: {
            "r:id": "rId1",
          },
        },
      });
    });

    it("should create a Hyperlink Click object with xmlns:a", () => {
      expect(buildHyperlinkClickObj("1", true)).to.deep.equal({
        "a:hlinkClick": {
          _attr: {
            "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r:id": "rId1",
          },
        },
      });
    });
  });

  describe("#buildHyperlinkHoverObj", () => {
    it("should create a Hyperlink Hover object", () => {
      expect(buildHyperlinkHoverObj("1", false)).to.deep.equal({
        "a:hlinkHover": {
          _attr: {
            "r:id": "rId1",
          },
        },
      });
    });

    it("should create a Hyperlink Hover object with xmlns:a", () => {
      expect(buildHyperlinkHoverObj("1", true)).to.deep.equal({
        "a:hlinkHover": {
          _attr: {
            "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r:id": "rId1",
          },
        },
      });
    });
  });
});
