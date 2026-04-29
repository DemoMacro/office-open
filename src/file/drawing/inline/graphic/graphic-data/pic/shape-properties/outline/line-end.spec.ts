import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createLineEnd } from "./line-end";

describe("createLineEnd", () => {
    it("should create a line end with type only", () => {
        const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "ARROW" }));
        expect(tree).to.deep.equal({
            "a:tailEnd": { _attr: { type: "arrow" } },
        });
    });

    it("should create a line end with type, width, and length", () => {
        const tree = new Formatter().format(
            createLineEnd("a:headEnd", { type: "TRIANGLE", width: "MEDIUM", length: "LARGE" }),
        );
        expect(tree).to.deep.equal({
            "a:headEnd": { _attr: { type: "triangle", w: "med", len: "lg" } },
        });
    });

    it("should create a line end with STEALTH type", () => {
        const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "STEALTH" }));
        expect(tree).to.deep.equal({
            "a:tailEnd": { _attr: { type: "stealth" } },
        });
    });

    it("should create a line end with DIAMOND type", () => {
        const tree = new Formatter().format(createLineEnd("a:headEnd", { type: "DIAMOND" }));
        expect(tree).to.deep.equal({
            "a:headEnd": { _attr: { type: "diamond" } },
        });
    });

    it("should create a line end with OVAL type", () => {
        const tree = new Formatter().format(createLineEnd("a:tailEnd", { type: "OVAL" }));
        expect(tree).to.deep.equal({
            "a:tailEnd": { _attr: { type: "oval" } },
        });
    });
});
