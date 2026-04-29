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

    it("should create border box properties with hideBot", () => {
        const tree = new Formatter().format(createMathBorderBoxProperties({ hideBot: true }));
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

    it("should create border box properties with strikeH", () => {
        const tree = new Formatter().format(createMathBorderBoxProperties({ strikeH: true }));
        expect(tree).to.deep.equal({
            "m:borderBoxPr": [{ "m:strikeH": {} }],
        });
    });

    it("should create border box properties with strikeV", () => {
        const tree = new Formatter().format(createMathBorderBoxProperties({ strikeV: true }));
        expect(tree).to.deep.equal({
            "m:borderBoxPr": [{ "m:strikeV": {} }],
        });
    });

    it("should create border box properties with strikeBLTR", () => {
        const tree = new Formatter().format(createMathBorderBoxProperties({ strikeBLTR: true }));
        expect(tree).to.deep.equal({
            "m:borderBoxPr": [{ "m:strikeBLTR": {} }],
        });
    });

    it("should create border box properties with strikeTLBR", () => {
        const tree = new Formatter().format(createMathBorderBoxProperties({ strikeTLBR: true }));
        expect(tree).to.deep.equal({
            "m:borderBoxPr": [{ "m:strikeTLBR": {} }],
        });
    });
});
