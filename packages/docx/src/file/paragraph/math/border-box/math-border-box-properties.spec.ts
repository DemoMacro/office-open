import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathBorderBoxProperties } from "./math-border-box-properties";

describe("createMathBorderBoxProperties", () => {
  it("should create empty border box properties", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({}));
    expect(tree).to.deep.equal({ "m:borderBoxPr": {} });
  });

  it("should create border box properties with hideTop", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ hideTop: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:hideTop": {} }],
    });
  });

  it("should create border box properties with hideBottom", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ hideBottom: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:hideBot": {} }],
    });
  });

  it("should create border box properties with hideLeft", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ hideLeft: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:hideLeft": {} }],
    });
  });

  it("should create border box properties with hideRight", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ hideRight: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:hideRight": {} }],
    });
  });

  it("should create border box properties with strikeHorizontal", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ strikeHorizontal: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:strikeH": {} }],
    });
  });

  it("should create border box properties with strikeVertical", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ strikeVertical: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:strikeV": {} }],
    });
  });

  it("should create border box properties with strikeDiagonalUp", () => {
    const tree = new Formatter().format(createMathBorderBoxProperties({ strikeDiagonalUp: true }));
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:strikeBLTR": {} }],
    });
  });

  it("should create border box properties with strikeDiagonalDown", () => {
    const tree = new Formatter().format(
      createMathBorderBoxProperties({ strikeDiagonalDown: true }),
    );
    expect(tree).to.deep.equal({
      "m:borderBoxPr": [{ "m:strikeTLBR": {} }],
    });
  });
});
