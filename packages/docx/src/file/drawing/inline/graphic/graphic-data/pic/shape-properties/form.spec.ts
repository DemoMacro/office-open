import { Formatter } from "@export/formatter";
import { createTransform2D } from "@office-open/core/drawingml";
import { describe, expect, it } from "vite-plus/test";

describe("Form (createTransform2D)", () => {
  describe("#constructor()", () => {
    it("should create", () => {
      const tree = new Formatter().format(
        createTransform2D({
          x: 0,
          y: 0,
          width: 100,
          height: 100,
        }),
      );

      expect(tree).to.deep.equal({
        "a:xfrm": [
          {
            "a:off": {
              _attr: {
                x: 0,
                y: 0,
              },
            },
          },
          {
            "a:ext": {
              _attr: {
                cx: 100,
                cy: 100,
              },
            },
          },
        ],
      });
    });

    it("should create with flip", () => {
      const tree = new Formatter().format(
        createTransform2D({
          x: 0,
          y: 0,
          width: 100,
          height: 100,
          flipHorizontal: true,
          flipVertical: true,
        }),
      );

      expect(tree).to.deep.equal({
        "a:xfrm": [
          {
            _attr: {
              flipH: true,
              flipV: true,
            },
          },
          {
            "a:off": {
              _attr: {
                x: 0,
                y: 0,
              },
            },
          },
          {
            "a:ext": {
              _attr: {
                cx: 100,
                cy: 100,
              },
            },
          },
        ],
      });
    });
  });
});
