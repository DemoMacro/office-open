// WPS (WordProcessing Shape) text boxes - modern DrawingML-based alternative to legacy VML text boxes (demo 94)
// Demonstrates: basic text box, styled fill/outline, rotation, floating positioning, and vertical alignment
import * as fs from "fs";

import {
    Document,
    HorizontalPositionRelativeFrom,
    Packer,
    Paragraph,
    TextRun,
    VerticalAnchor,
    VerticalPositionRelativeFrom,
    WpsShapeRun,
} from "docx";

const doc = new Document({
    sections: [
        {
            children: [
                // 1. Basic text box - minimal options
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun(
                                            "This is a basic WPS text box with just width and height.",
                                        ),
                                    ],
                                }),
                            ],
                            transformation: {
                                height: 800_000,
                                width: 4_000_000,
                            },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Styled text box - solid fill and outline
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            bold: true,
                                            color: "1F3864",
                                            text: "Styled text box with a light blue background and dark blue border.",
                                        }),
                                    ],
                                }),
                            ],
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "2E74B5",
                                width: 25_400,
                            },
                            solidFill: {
                                type: "rgb",
                                value: "D6E4F0",
                            },
                            transformation: {
                                height: 800_000,
                                width: 4_000_000,
                            },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Rotated text box
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    children: [new TextRun("This text box is rotated 15 degrees.")],
                                }),
                            ],
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "FF0000",
                                width: 12_700,
                            },
                            transformation: {
                                height: 600_000,
                                rotation: 15,
                                width: 3_000_000,
                            },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ children: [new TextRun("")] }),

                // 4. Floating text box - positioned via offset
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun(
                                            "Floating text box positioned at a specific offset on the page.",
                                        ),
                                    ],
                                }),
                            ],
                            floating: {
                                horizontalPosition: {
                                    offset: 3_500_000,
                                    relative: HorizontalPositionRelativeFrom.PAGE,
                                },
                                verticalPosition: {
                                    offset: 100_000,
                                    relative: VerticalPositionRelativeFrom.PARAGRAPH,
                                },
                            },
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "70AD47",
                                width: 19_050,
                            },
                            transformation: {
                                height: 700_000,
                                width: 3_500_000,
                            },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ children: [new TextRun("")] }),

                // 5. Centered vertical alignment with custom margins
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            bodyProperties: {
                                margins: {
                                    bottom: 72_000,
                                    left: 144_000,
                                    right: 144_000,
                                    top: 72_000,
                                },
                                verticalAnchor: VerticalAnchor.CENTER,
                            },
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun(
                                            "Vertically centered text with custom margins.",
                                        ),
                                    ],
                                }),
                            ],
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "BF8F00",
                                width: 19_050,
                            },
                            solidFill: {
                                type: "rgb",
                                value: "FFF2CC",
                            },
                            transformation: {
                                height: 1_200_000,
                                width: 4_000_000,
                            },
                            type: "wps",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
