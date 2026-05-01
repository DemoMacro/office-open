import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathEqArrProperties } from "./math-eq-arr-properties";

describe("createMathEqArrProperties", () => {
    it("should create empty eq arr properties", () => {
        const tree = new Formatter().format(createMathEqArrProperties({}));
        expect(tree).to.deep.equal({ "m:eqArrPr": {} });
    });

    it("should create eq arr properties with baseJc", () => {
        const tree = new Formatter().format(createMathEqArrProperties({ baseJc: "center" }));
        expect(tree).to.deep.equal({
            "m:eqArrPr": [{ "m:baseJc": { _attr: { "m:val": "center" } } }],
        });
    });

    it("should create eq arr properties with maxDist", () => {
        const tree = new Formatter().format(createMathEqArrProperties({ maxDist: true }));
        expect(tree).to.deep.equal({
            "m:eqArrPr": [{ "m:maxDist": {} }],
        });
    });

    it("should create eq arr properties with objDist", () => {
        const tree = new Formatter().format(createMathEqArrProperties({ objDist: false }));
        expect(tree).to.deep.equal({
            "m:eqArrPr": [{ "m:objDist": { _attr: { "w:val": false } } }],
        });
    });

    it("should create eq arr properties with rSpRule", () => {
        const tree = new Formatter().format(createMathEqArrProperties({ rSpRule: "single" }));
        expect(tree).to.deep.equal({
            "m:eqArrPr": [{ "m:rSpRule": { _attr: { "m:val": "single" } } }],
        });
    });

    it("should create eq arr properties with rSp", () => {
        const tree = new Formatter().format(createMathEqArrProperties({ rSp: 100 }));
        expect(tree).to.deep.equal({
            "m:eqArrPr": [{ "m:rSp": { _attr: { "m:val": 100 } } }],
        });
    });
});
