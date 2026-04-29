import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathMatrixProperties } from "./math-matrix-properties";

describe("createMathMatrixProperties", () => {
    it("should create empty matrix properties", () => {
        const tree = new Formatter().format(createMathMatrixProperties({}));
        expect(tree).to.deep.equal({ "m:mPr": {} });
    });

    it("should create matrix properties with baseJc", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ baseJc: "center" }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:baseJc": { _attr: { "m:val": "center" } } }],
        });
    });

    it("should create matrix properties with plcHide", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ plcHide: true }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:plcHide": {} }],
        });
    });

    it("should create matrix properties with rSpRule", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ rSpRule: "single" }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:rSpRule": { _attr: { "m:val": "single" } } }],
        });
    });

    it("should create matrix properties with cGpRule", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ cGpRule: "double" }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:cGpRule": { _attr: { "m:val": "double" } } }],
        });
    });

    it("should create matrix properties with rSp", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ rSp: 100 }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:rSp": { _attr: { "m:val": 100 } } }],
        });
    });

    it("should create matrix properties with cSp", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ cSp: 200 }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:cSp": { _attr: { "m:val": 200 } } }],
        });
    });

    it("should create matrix properties with cGp", () => {
        const tree = new Formatter().format(createMathMatrixProperties({ cGp: 300 }));
        expect(tree).to.deep.equal({
            "m:mPr": [{ "m:cGp": { _attr: { "m:val": 300 } } }],
        });
    });
});
