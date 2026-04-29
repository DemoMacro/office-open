import { Formatter } from "@export/formatter";
import { BuilderElement } from "@file/xml-components";
import { describe, expect, it } from "vite-plus/test";

import { createTableStyleOverride } from "./table-style-override";

describe("createTableStyleOverride", () => {
    it("should create override with type only", () => {
        const tree = new Formatter().format(createTableStyleOverride({ type: "wholeTable" }));
        expect(tree).to.deep.equal({
            "w:tblStylePr": { _attr: { "w:type": "wholeTable" } },
        });
    });

    it("should create override for first row", () => {
        const tree = new Formatter().format(createTableStyleOverride({ type: "firstRow" }));
        expect(tree).to.deep.equal({
            "w:tblStylePr": { _attr: { "w:type": "firstRow" } },
        });
    });

    it("should create override with paragraph properties", () => {
        const tree = new Formatter().format(
            createTableStyleOverride({
                type: "firstRow",
                paragraphProperties: new BuilderElement({
                    name: "w:pPr",
                    children: [
                        new BuilderElement({
                            name: "w:jc",
                            attributes: { val: { key: "w:val", value: "center" } },
                        }),
                    ],
                }),
            }),
        );
        expect(tree).to.deep.equal({
            "w:tblStylePr": [
                { _attr: { "w:type": "firstRow" } },
                {
                    "w:pPr": [{ "w:jc": { _attr: { "w:val": "center" } } }],
                },
            ],
        });
    });

    it("should create override with multiple property types", () => {
        const tree = new Formatter().format(
            createTableStyleOverride({
                type: "band1Horz",
                runProperties: new BuilderElement({
                    name: "w:rPr",
                    children: [
                        new BuilderElement({
                            name: "w:b",
                        }),
                    ],
                }),
                cellProperties: new BuilderElement({
                    name: "w:tcPr",
                    children: [
                        new BuilderElement({
                            name: "w:shd",
                            attributes: {
                                val: { key: "w:val", value: "clear" },
                                fill: { key: "w:fill", value: "E7E6E6" },
                            },
                        }),
                    ],
                }),
            }),
        );
        expect(tree).to.deep.equal({
            "w:tblStylePr": [
                { _attr: { "w:type": "band1Horz" } },
                {
                    "w:rPr": [{ "w:b": {} }],
                },
                {
                    "w:tcPr": [{ "w:shd": { _attr: { "w:fill": "E7E6E6", "w:val": "clear" } } }],
                },
            ],
        });
    });
});
