import * as fs from "fs";

import { Presentation, Shape, Packer, SolidFill, Paragraph, Run } from "../src";

const pres = new Presentation({
    title: "My Presentation",
    creator: "Demo",
    slides: [
        // Slide 1: Basic shapes
        {
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 400,
                    height: 200,
                    text: "Hello World",
                    geometry: "rect",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 200,
                    y: 350,
                    width: 500,
                    height: 100,
                    text: "Second shape",
                }),
            ],
        },
        // Slide 2: Full width
        {
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 860,
                    height: 440,
                    text: "Slide 2 - Full Width",
                    geometry: "rect",
                }),
            ],
        },
        // Slide 3: Vertical text
        {
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 600,
                    height: 50,
                    text: "Vertical Text",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 120,
                    height: 300,
                    textVertical: "vert",
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "Vertical Text (top to bottom)", fontSize: 14 }),
                            ],
                        }),
                    ],
                    outline: { color: "4472C4", width: 1 },
                }),
                new Shape({
                    x: 200,
                    y: 120,
                    width: 120,
                    height: 300,
                    textVertical: "vert270",
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "Rotated 270 (bottom to top)", fontSize: 14 }),
                            ],
                        }),
                    ],
                    outline: { color: "ED7D31", width: 1 },
                }),
                new Shape({
                    x: 350,
                    y: 120,
                    width: 120,
                    height: 300,
                    textVertical: "horz",
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [new Run({ text: "Horizontal (default)", fontSize: 14 })],
                        }),
                    ],
                    outline: { color: "70AD47", width: 1 },
                }),
            ],
        },
        // Slide 4: Text anchor & auto-fit
        {
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 600,
                    height: 50,
                    text: "Text Anchor & Auto-Fit",
                    fill: new SolidFill("4472C4"),
                }),
                // Top anchor
                new Shape({
                    x: 50,
                    y: 120,
                    width: 200,
                    height: 200,
                    textAnchor: "t",
                    text: "Top anchored text",
                    outline: { color: "999999", width: 1 },
                }),
                // Center anchor
                new Shape({
                    x: 280,
                    y: 120,
                    width: 200,
                    height: 200,
                    textAnchor: "ctr",
                    text: "Center anchored text",
                    outline: { color: "999999", width: 1 },
                }),
                // Bottom anchor
                new Shape({
                    x: 510,
                    y: 120,
                    width: 200,
                    height: 200,
                    textAnchor: "b",
                    text: "Bottom anchored text",
                    outline: { color: "999999", width: 1 },
                }),
                // Auto-fit normal
                new Shape({
                    x: 50,
                    y: 350,
                    width: 250,
                    height: 80,
                    textAutoFit: "normal",
                    text: "This is a very long text that should auto-fit to shrink within the shape bounds",
                    outline: { color: "4472C4", width: 1 },
                }),
                // Auto-fit shape
                new Shape({
                    x: 330,
                    y: 350,
                    width: 250,
                    height: 80,
                    textAutoFit: "shape",
                    text: "Shape auto-fit text",
                    outline: { color: "ED7D31", width: 1 },
                }),
            ],
        },
        // Slide 5: Text margins & columns
        {
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 600,
                    height: 50,
                    text: "Text Margins & Columns",
                    fill: new SolidFill("4472C4"),
                }),
                // Default margins
                new Shape({
                    x: 50,
                    y: 120,
                    width: 350,
                    height: 150,
                    outline: { color: "999999", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Default margins (no extra padding)",
                                    fontSize: 12,
                                }),
                            ],
                        }),
                    ],
                }),
                // Wide margins
                new Shape({
                    x: 430,
                    y: 120,
                    width: 350,
                    height: 150,
                    textMargins: { top: 100000, bottom: 100000, left: 200000, right: 200000 },
                    outline: { color: "ED7D31", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Wide margins (extra padding all around)",
                                    fontSize: 12,
                                }),
                            ],
                        }),
                    ],
                }),
                // 2 columns
                new Shape({
                    x: 50,
                    y: 300,
                    width: 730,
                    height: 150,
                    textColumns: 2,
                    textColumnSpacing: 12,
                    outline: { color: "70AD47", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "This is column 1 text. The shape is divided into 2 columns with spacing between them.",
                                    fontSize: 12,
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
