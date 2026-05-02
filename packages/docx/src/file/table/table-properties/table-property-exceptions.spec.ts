import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { ShadingType } from "@file/shading";
import { describe, expect, it } from "vite-plus/test";

import { TablePropertyExceptions } from "./table-property-exceptions";

describe("TablePropertyExceptions", () => {
    it("should create an empty tblPrEx", () => {
        const ex = new TablePropertyExceptions({});
        const tree = new Formatter().format(ex);
        expect(tree).to.deep.equal({ "w:tblPrEx": {} });
    });

    it("should create tblPrEx with width", () => {
        const ex = new TablePropertyExceptions({
            width: { size: 5000, type: "pct" },
        });
        const tree = new Formatter().format(ex);
        expect(tree).to.deep.equal({
            "w:tblPrEx": [{ "w:tblW": { _attr: { "w:type": "pct", "w:w": "5000%" } } }],
        });
    });

    it("should create tblPrEx with borders", () => {
        const ex = new TablePropertyExceptions({
            borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: "FF0000" },
            },
        });
        const tree = new Formatter().format(ex);
        const tblPrEx = tree["w:tblPrEx"];
        const content = Array.isArray(tblPrEx) ? tblPrEx[0] : tblPrEx;
        expect(content).to.haveOwnProperty("w:tblBorders");
    });

    it("should create tblPrEx with shading", () => {
        const ex = new TablePropertyExceptions({
            shading: { fill: "FFFF00", type: ShadingType.CLEAR },
        });
        const tree = new Formatter().format(ex);
        expect(tree).to.deep.equal({
            "w:tblPrEx": [{ "w:shd": { _attr: { "w:fill": "FFFF00", "w:val": "clear" } } }],
        });
    });

    it("should create tblPrEx with alignment", () => {
        const ex = new TablePropertyExceptions({
            alignment: "center",
        });
        const tree = new Formatter().format(ex);
        expect(tree).to.deep.equal({
            "w:tblPrEx": [{ "w:jc": { _attr: { "w:val": "center" } } }],
        });
    });
});
