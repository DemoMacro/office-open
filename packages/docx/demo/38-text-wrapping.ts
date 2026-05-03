// Example of how to "wrap" text around an image
// Demonstrates: SQUARE, TIGHT, THROUGH, TOP_AND_BOTTOM, NONE (in front of text), NONE (behind text)

import * as fs from "fs";

// Import { Document, Packer, Paragraph } from "@office-open/docx";
import {
    Document,
    ImageRun,
    Packer,
    Paragraph,
    TextWrappingSide,
    TextWrappingType,
} from "@office-open/docx";

const imageData = fs.readFileSync("./demo/images/pizza.gif");

const textParagraphs = [
    new Paragraph(
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque vehicula nec nulla vitae efficitur. Ut interdum mauris eu ipsum rhoncus, nec pharetra velit placerat. Sed vehicula libero ac urna molestie, id pharetra est pellentesque. Praesent iaculis vehicula fringilla. Duis pretium gravida orci eu vestibulum. Mauris tincidunt ipsum dolor, ut ornare dolor pellentesque id. Integer in nulla gravida, lacinia ante non, commodo ex. Vivamus vulputate nisl id lectus finibus vulputate. Ut et nisl mi. Cras fermentum augue arcu, ac accumsan elit euismod id. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Sed ac posuere nisi. Pellentesque tincidunt vehicula bibendum. Phasellus eleifend viverra nisl.",
    ),
    new Paragraph(
        "Proin ac purus faucibus, porttitor magna ut, cursus nisl. Vivamus ante purus, porta accumsan nibh eget, eleifend dignissim odio. Integer sed dictum est, aliquam lacinia justo. Donec ultrices auctor venenatis. Etiam interdum et elit nec elementum. Pellentesque nec viverra mauris. Etiam suscipit leo nec velit fringilla mattis. Pellentesque justo lacus, sodales eu condimentum in, dapibus finibus lacus. Morbi vitae nibh sit amet sem molestie feugiat. In non porttitor enim.",
    ),
    new Paragraph(
        "Ut eget diam cursus quam accumsan interdum at id ante. Ut mollis mollis arcu, eu scelerisque dui tempus in. Quisque aliquam, augue quis ornare aliquam, ex purus ultrices mauris, ut porta dolor dolor nec justo. Nunc a tempus odio, eu viverra arcu. Suspendisse vitae nibh nec mi pharetra tempus. Mauris ut ullamcorper sapien, et sagittis sapien. Vestibulum in urna metus. In scelerisque, massa id bibendum tempus, quam orci rutrum turpis, a feugiat nisi ligula id metus. Praesent id dictum purus. Proin interdum ipsum nulla.",
    ),
    new Paragraph(
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.",
    ),
];

const doc = new Document({
    sections: [
        // 1. SQUARE — Text wraps around the rectangular bounding box
        {
            properties: {},
            children: [
                new Paragraph("SQUARE wrapping:"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                margins: {
                                    bottom: 201_440,
                                    top: 201_440,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    side: TextWrappingSide.BOTH_SIDES,
                                    type: TextWrappingType.SQUARE,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
        // 2. TIGHT — Text wraps tightly around the image contours
        {
            properties: {},
            children: [
                new Paragraph("TIGHT wrapping:"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                margins: {
                                    bottom: 201_440,
                                    top: 201_440,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    side: TextWrappingSide.BOTH_SIDES,
                                    type: TextWrappingType.TIGHT,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
        // 3. THROUGH — Text wraps through the image contours
        {
            properties: {},
            children: [
                new Paragraph("THROUGH wrapping:"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                margins: {
                                    bottom: 201_440,
                                    top: 201_440,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    side: TextWrappingSide.BOTH_SIDES,
                                    type: TextWrappingType.THROUGH,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
        // 4. TOP_AND_BOTTOM — Text appears only above and below the image
        {
            properties: {},
            children: [
                new Paragraph("TOP_AND_BOTTOM wrapping:"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                margins: {
                                    bottom: 201_440,
                                    top: 201_440,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    type: TextWrappingType.TOP_AND_BOTTOM,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
        // 5. NONE — In front of text (behindDocument: false, default)
        {
            properties: {},
            children: [
                new Paragraph("NONE wrapping (in front of text):"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    type: TextWrappingType.NONE,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
        // 6. NONE — Behind text (behindDocument: true)
        {
            properties: {},
            children: [
                new Paragraph("NONE wrapping (behind text):"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: imageData,
                            floating: {
                                behindDocument: true,
                                horizontalPosition: {
                                    offset: 2_014_400,
                                },
                                verticalPosition: {
                                    offset: 0,
                                    relative: "paragraph",
                                },
                                wrap: {
                                    type: TextWrappingType.NONE,
                                },
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                ...textParagraphs,
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
