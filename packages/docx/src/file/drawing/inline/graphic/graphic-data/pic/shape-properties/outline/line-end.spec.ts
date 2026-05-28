import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createLineEnd } from "./line-end";

describe("createLineEnd", () => {
  it("should create a line end with type only", () => {
    const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "arrow" }));
    expect(tree).to.deep.equal({
      "a:tailEnd": { _attr: { type: "arrow" } },
    });
  });

  it("should create a line end with type, width, and length", () => {
    const tree = new Formatter().format(
      createLineEnd("a:headEnd", { type: "triangle", width: "medium", length: "large" }),
    );
    expect(tree).to.deep.equal({
      "a:headEnd": { _attr: { type: "triangle", w: "med", len: "lg" } },
    });
  });

  it("should create a line end with STEALTH type", () => {
    const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "stealth" }));
    expect(tree).to.deep.equal({
      "a:tailEnd": { _attr: { type: "stealth" } },
    });
  });

  it("should create a line end with DIAMOND type", () => {
    const tree = new Formatter().format(createLineEnd("a:headEnd", { type: "diamond" }));
    expect(tree).to.deep.equal({
      "a:headEnd": { _attr: { type: "diamond" } },
    });
  });

  it("should create a line end with OVAL type", () => {
    const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "oval" }));
    expect(tree).to.deep.equal({
      "a:tailEnd": { _attr: { type: "oval" } },
    });
  });
});
