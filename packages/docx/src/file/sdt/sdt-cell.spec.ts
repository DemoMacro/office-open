import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "../paragraph";
import { TableCell } from "../table";
import { StructuredDocumentTagCell } from "./sdt-cell";

describe("StructuredDocumentTagCell", () => {
    it("should create a cell SDT with properties only", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagCell({
                properties: { alias: "Cell Control" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "Cell Control" } } }],
                },
            ],
        });
    });

    it("should create a cell SDT with children content", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagCell({
                properties: { text: {} },
                children: [
                    new TableCell({
                        children: [new Paragraph("test")],
                    }),
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                { "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] },
                {
                    "w:sdtContent": [
                        {
                            "w:tc": [
                                {
                                    "w:p": [
                                        {
                                            "w:r": [
                                                {
                                                    "w:t": [
                                                        { _attr: { "xml:space": "preserve" } },
                                                        "test",
                                                    ],
                                                },
                                            ],
                                        },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a cell SDT with empty children array", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagCell({
                properties: { text: {} },
                children: [],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [{ "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] }],
        });
    });
});
