import { Formatter } from "@export/formatter";
import { Paragraph } from "@file/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { Textbox } from "./textbox";

describe("VmlTextbox", () => {
    it("should work", () => {
        const tree = new Formatter().format(
            new Textbox({
                children: [new Paragraph("test-content")],
                style: {
                    width: "10pt",
                },
            }),
        );

        expect(tree).toStrictEqual({
            "w:p": [
                {
                    "w:pict": [
                        {
                            "v:shape": [
                                {
                                    _attr: {
                                        id: expect.any(String),
                                        style: "width:10pt",
                                        type: "#_x0000_t202",
                                    },
                                },
                                {
                                    "v:textbox": [
                                        {
                                            _attr: {
                                                insetmode: "auto",
                                                style: "mso-fit-shape-to-text:t;",
                                            },
                                        },
                                        {
                                            "w:txbxContent": [
                                                {
                                                    "w:p": [
                                                        {
                                                            "w:r": [
                                                                {
                                                                    "w:t": [
                                                                        {
                                                                            _attr: {
                                                                                "xml:space":
                                                                                    "preserve",
                                                                            },
                                                                        },
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
                        },
                    ],
                },
            ],
        });
    });
});
