import { Formatter } from "@export/formatter";
import { BuilderElement } from "@file/xml-components";
import { describe, expect, it } from "vite-plus/test";

import { StructuredDocumentTagBlock } from "./sdt";

describe("StructuredDocumentTagBlock", () => {
    it("should create a block SDT with properties only", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagBlock({
                properties: { alias: "Block Control" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "Block Control" } } }],
                },
            ],
        });
    });

    it("should create a block SDT with children content", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagBlock({
                properties: { text: {} },
                children: [
                    new BuilderElement({
                        name: "w:p",
                        children: [
                            new BuilderElement({
                                name: "w:r",
                                children: [
                                    new BuilderElement({
                                        name: "w:t",
                                    }),
                                ],
                            }),
                        ],
                    }),
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                { "w:sdtPr": [{ "w:text": {} }] },
                {
                    "w:sdtContent": [
                        {
                            "w:p": [
                                {
                                    "w:r": [{ "w:t": {} }],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a block SDT with empty children array", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagBlock({
                properties: { text: {} },
                children: [],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [{ "w:sdtPr": [{ "w:text": {} }] }],
        });
    });
});
