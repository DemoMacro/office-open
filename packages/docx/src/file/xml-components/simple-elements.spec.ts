import { Formatter } from "@export/formatter";
import { BuilderElement, stringEnumValObj } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

describe("BuilderElement", () => {
  describe("#constructor()", () => {
    it("should create a simple BuilderElement", () => {
      const element = new BuilderElement({
        name: "test",
      });

      const tree = new Formatter().format(element);
      expect(tree).to.deep.equal({
        test: {},
      });
    });

    it("should create a simple BuilderElement with attributes", () => {
      const element = new BuilderElement<{ readonly testAttr: string }>({
        attributes: {
          testAttr: {
            key: "w:testAttr",
            value: "test",
          },
        },
        name: "test",
      });

      const tree = new Formatter().format(element);
      expect(tree).to.deep.equal({
        test: {
          _attr: {
            "w:testAttr": "test",
          },
        },
      });
    });

    it("should accept IXmlableObject directly as child", () => {
      const element = new BuilderElement({
        name: "w:pPr",
        children: [stringEnumValObj("w:jc", "center")],
      });

      const tree = new Formatter().format(element);
      expect(tree).to.deep.equal({
        "w:pPr": [
          {
            "w:jc": {
              _attr: {
                "w:val": "center",
              },
            },
          },
        ],
      });
    });
  });
});
