import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathPhantProperties } from "./math-phant-properties";

describe("createMathPhantProperties", () => {
    it("should create empty phant properties", () => {
        const tree = new Formatter().format(createMathPhantProperties({}));
        expect(tree).to.deep.equal({ "m:phantPr": {} });
    });

    it("should create phant properties with show", () => {
        const tree = new Formatter().format(createMathPhantProperties({ show: true }));
        expect(tree).to.deep.equal({
            "m:phantPr": [{ "m:show": {} }],
        });
    });

    it("should create phant properties with zeroWid", () => {
        const tree = new Formatter().format(createMathPhantProperties({ zeroWid: true }));
        expect(tree).to.deep.equal({
            "m:phantPr": [{ "m:zeroWid": {} }],
        });
    });

    it("should create phant properties with zeroAsc", () => {
        const tree = new Formatter().format(createMathPhantProperties({ zeroAsc: true }));
        expect(tree).to.deep.equal({
            "m:phantPr": [{ "m:zeroAsc": {} }],
        });
    });

    it("should create phant properties with zeroDesc", () => {
        const tree = new Formatter().format(createMathPhantProperties({ zeroDesc: true }));
        expect(tree).to.deep.equal({
            "m:phantPr": [{ "m:zeroDesc": {} }],
        });
    });

    it("should create phant properties with transp", () => {
        const tree = new Formatter().format(createMathPhantProperties({ transp: true }));
        expect(tree).to.deep.equal({
            "m:phantPr": [{ "m:transp": {} }],
        });
    });
});
