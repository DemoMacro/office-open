import { Formatter } from "@export/formatter";
import { Paragraph } from "@file/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { createVmlTextbox } from "./vml-texbox";

describe("VmlTextbox", () => {
    it("should work", () => {
        const tree = new Formatter().format(
            createVmlTextbox({
                children: [new Paragraph("test-content")],
                style: "test-style",
            }),
        );

        expect(tree).toStrictEqual({
            "v:textbox": [
                { _attr: { insetmode: "auto", style: "test-style" } },
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
        });
    });

    it("should work with inset", () => {
        const tree = new Formatter().format(
            createVmlTextbox({
                children: [new Paragraph("test-content")],
                inset: {
                    bottom: 0,
                    left: 0,
                    right: 0,
                    top: 0,
                },
                style: "test-style",
            }),
        );

        expect(tree).toStrictEqual({
            "v:textbox": [
                { _attr: { inset: "0, 0, 0, 0", insetmode: "custom", style: "test-style" } },
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
        });
    });
});
