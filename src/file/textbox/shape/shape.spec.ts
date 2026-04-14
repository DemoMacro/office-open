import { Formatter } from "@export/formatter";
import { Paragraph } from "@file/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { createShape } from "./shape";

describe("createShape", () => {
    it("should work", () => {
        const tree = new Formatter().format(
            createShape({
                children: [new Paragraph("test-content")],
                id: "test-id",
                style: {
                    width: "10pt",
                },
            }),
        );

        expect(tree).toStrictEqual({
            "v:shape": [
                { _attr: { id: "test-id", style: "width:10pt", type: "#_x0000_t202" } },
                {
                    "v:textbox": [
                        { _attr: { insetmode: "auto", style: "mso-fit-shape-to-text:t;" } },
                        {
                            "w:txbxContent": [
                                {
                                    "w:p": [
                                        {
                                            "w:r": [
                                                {
                                                    "w:t": [
                                                        { _attr: { "xml:space": "preserve" } },
                                                        "test-content",
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

    it("should create default styles", () => {
        const tree = new Formatter().format(
            createShape({
                id: "test-id",
            }),
        );

        expect(tree).toStrictEqual({
            "v:shape": [
                { _attr: { id: "test-id", type: "#_x0000_t202" } },
                {
                    "v:textbox": [
                        { _attr: { insetmode: "auto", style: "mso-fit-shape-to-text:t;" } },
                        {
                            "w:txbxContent": {},
                        },
                    ],
                },
            ],
        });
    });
});
