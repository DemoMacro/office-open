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
} from "@office-open/docx";

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
                                height: 80,
                                width: 400,
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
                                color: { value: "2E74B5" },
                                type: "solidFill",
                                width: 25_400,
                            },
                            solidFill: { value: "D6E4F0" },
                            transformation: {
                                height: 80,
                                width: 400,
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
                                color: { value: "FF0000" },
                                type: "solidFill",
                                width: 12_700,
                            },
                            transformation: {
                                height: 60,
                                rotation: 15,
                                width: 300,
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
                                color: { value: "70AD47" },
                                type: "solidFill",
                                width: 19_050,
                            },
                            transformation: {
                                height: 70,
                                width: 350,
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
                                color: { value: "BF8F00" },
                                type: "solidFill",
                                width: 19_050,
                            },
                            solidFill: { value: "FFF2CC" },
                            transformation: {
                                height: 120,
                                width: 400,
                            },
                            type: "wps",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
