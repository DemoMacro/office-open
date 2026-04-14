import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { describe, expect, it } from "vite-plus/test";

import { PageBorderDisplay, PageBorderZOrder, PageBorders } from "./page-borders";

describe("PageBorders", () => {
    describe("#constructor()", () => {
        it("should create empty element when no options are passed", () => {
            const properties = new PageBorders();
            expect(() => new Formatter().format(properties)).to.throw();
        });

        it("should create page borders with some configuration", () => {
            const properties = new PageBorders({
                pageBorders: {
                    display: PageBorderDisplay.FIRST_PAGE,
                },
            });
            const tree = new Formatter().format(properties);

            expect(Object.keys(tree)).to.deep.equal(["w:pgBorders"]);
            expect(tree["w:pgBorders"]).to.deep.equal({ _attr: { "w:display": "firstPage" } });
        });

        it("should create page borders with default configuration", () => {
            const properties = new PageBorders({});
            const tree = new Formatter().format(properties);

            expect(Object.keys(tree)).to.deep.equal(["w:pgBorders"]);
            expect(tree).to.deep.equal({
                "w:pgBorders": {
                    _attr: {},
                },
            });
        });

        it("should create page borders with full configuration", () => {
            const properties = new PageBorders({
                pageBorderBottom: {
                    color: "556677",
                    size: 30,
                    style: BorderStyle.SINGLE,
                },
                pageBorderLeft: {
                    color: "889900",
                    size: 40,
                    style: BorderStyle.DOTTED,
                },
                pageBorderRight: {
                    color: "223344",
                    size: 20,
                    style: BorderStyle.DOUBLE,
                },
                pageBorderTop: {
                    color: "001122",
                    size: 10,
                    style: BorderStyle.DOUBLE_WAVE,
                },
                pageBorders: {
                    display: PageBorderDisplay.FIRST_PAGE,
                    zOrder: PageBorderZOrder.BACK,
                },
            });
            const tree = new Formatter().format(properties);

            expect(Object.keys(tree)).to.deep.equal(["w:pgBorders"]);
            expect(tree["w:pgBorders"]).to.be.an.instanceof(Array);
            expect(tree["w:pgBorders"][0]).to.deep.equal({
                _attr: { "w:display": "firstPage", "w:zOrder": "back" },
            });
            expect(tree["w:pgBorders"][1]).to.deep.equal({
                "w:top": {
                    _attr: { "w:color": "001122", "w:sz": 10, "w:val": "doubleWave" },
                },
            });
            expect(tree["w:pgBorders"][2]).to.deep.equal({
                "w:left": {
                    _attr: { "w:color": "889900", "w:sz": 40, "w:val": "dotted" },
                },
            });
            expect(tree["w:pgBorders"][3]).to.deep.equal({
                "w:bottom": {
                    _attr: { "w:color": "556677", "w:sz": 30, "w:val": "single" },
                },
            });
            expect(tree["w:pgBorders"][4]).to.deep.equal({
                "w:right": {
                    _attr: { "w:color": "223344", "w:sz": 20, "w:val": "double" },
                },
            });
        });
    });
});
