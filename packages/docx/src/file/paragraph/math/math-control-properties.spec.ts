import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathControlProperties } from "./math-control-properties";

describe("createMathControlProperties", () => {
    it("should create empty control properties when no options", () => {
        const tree = new Formatter().format(createMathControlProperties());
        expect(tree).to.deep.equal({ "m:ctrlPr": {} });
    });

    it("should create empty control properties when undefined", () => {
        const tree = new Formatter().format(createMathControlProperties(undefined));
        expect(tree).to.deep.equal({ "m:ctrlPr": {} });
    });

    it("should create control properties with insertionReference", () => {
        const tree = new Formatter().format(
            createMathControlProperties({ insertionReference: "123" }),
        );
        expect(tree).to.deep.equal({
            "m:ctrlPr": [{ "w:ins": { _attr: { "w:id": "123" } } }],
        });
    });

    it("should create control properties with deletionReference", () => {
        const tree = new Formatter().format(
            createMathControlProperties({ deletionReference: "456" }),
        );
        expect(tree).to.deep.equal({
            "m:ctrlPr": [{ "w:del": { _attr: { "w:id": "456" } } }],
        });
    });

    it("should create control properties with both references", () => {
        const tree = new Formatter().format(
            createMathControlProperties({
                insertionReference: "100",
                deletionReference: "200",
            }),
        );
        expect(tree).to.deep.equal({
            "m:ctrlPr": [
                { "w:ins": { _attr: { "w:id": "100" } } },
                { "w:del": { _attr: { "w:id": "200" } } },
            ],
        });
    });
});
