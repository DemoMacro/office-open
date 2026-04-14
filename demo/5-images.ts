// Example of how to add images to the document - You can use Buffers, UInt8Arrays or Base64 strings

import * as fs from "fs";

import {
    Document,
    HorizontalPositionAlign,
    HorizontalPositionRelativeFrom,
    ImageRun,
    Packer,
    Paragraph,
    VerticalPositionAlign,
    VerticalPositionRelativeFrom,
    convertMillimetersToTwip,
} from "docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph("Hello World"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            altText: {
                                description: "This is an ultimate image",
                                name: "My Ultimate Image",
                                title: "This is an ultimate title",
                            },
                            data: fs.readFileSync("./demo/images/image1.jpeg"),
                            transformation: {
                                height: 100,
                                width: 100,
                            },
                            type: "jpg",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/dog.png").toString("base64"),
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "FF0000",
                            },
                            transformation: {
                                height: 100,
                                width: 100,
                            },
                            type: "png",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            outline: {
                                solidFillType: "rgb",
                                type: "solidFill",
                                value: "0000FF",
                                width: convertMillimetersToTwip(600),
                            },
                            transformation: {
                                flip: {
                                    vertical: true,
                                },
                                height: 100,
                                width: 100,
                            },
                            type: "jpg",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/parrots.bmp"),
                            transformation: {
                                flip: {
                                    horizontal: true,
                                },
                                height: 150,
                                rotation: 225,
                                width: 150,
                            },
                            type: "bmp",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                flip: {
                                    horizontal: true,
                                    vertical: true,
                                },
                                height: 200,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            floating: {
                                horizontalPosition: {
                                    offset: 1_014_400,
                                },
                                verticalPosition: {
                                    offset: 1_014_400,
                                },
                                zIndex: 10,
                            },
                            transformation: {
                                height: 200,
                                rotation: 45,
                                width: 200,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            floating: {
                                horizontalPosition: {
                                    align: HorizontalPositionAlign.RIGHT,
                                    relative: HorizontalPositionRelativeFrom.PAGE,
                                },
                                verticalPosition: {
                                    align: VerticalPositionAlign.BOTTOM,
                                    relative: VerticalPositionRelativeFrom.PAGE,
                                },
                                zIndex: 5,
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "jpg",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/linux-svg.svg"),
                            fallback: {
                                data: fs.readFileSync("./demo/images/linux-png.png"),
                                type: "png",
                            },
                            transformation: {
                                height: 200,
                                width: 200,
                            },
                            type: "svg",
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
