import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "../paragraph";
import { TableCell, TableRow } from "../table";
import { StructuredDocumentTagRow } from "./sdt-row";

describe("StructuredDocumentTagRow", () => {
    it("should create a row SDT with properties only", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRow({
                properties: { alias: "Row Control" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "Row Control" } } }],
                },
            ],
        });
    });

    it("should create a row SDT with children content", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRow({
                properties: { text: {} },
                children: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph("test")],
                            }),
                        ],
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
                            "w:tr": [
                                {
                                    "w:tc": [
                                        {
                                            "w:p": [
                                                {
                                                    "w:r": [
                                                        {
                                                            "w:t": [
                                                                {
                                                                    _attr: {
                                                                        "xml:space": "preserve",
                                                                    },
                                                                },
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
                },
            ],
        });
    });

    it("should create a row SDT with empty children array", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRow({
                properties: { text: {} },
                children: [],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [{ "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] }],
        });
    });
});
