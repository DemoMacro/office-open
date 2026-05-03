// Text Frame (Text Box) example

import * as fs from "fs";

import {
    AlignmentType,
    BorderStyle,
    Document,
    FrameAnchorType,
    HorizontalPositionAlign,
    Packer,
    Paragraph,
    Tab,
    TextRun,
    VerticalPositionAlign,
} from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    border: {
                        bottom: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        left: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        right: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        top: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                    },
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                        }),
                    ],
                    frame: {
                        anchor: {
                            horizontal: FrameAnchorType.MARGIN,
                            vertical: FrameAnchorType.MARGIN,
                        },
                        height: 1000,
                        position: {
                            x: 1000,
                            y: 3000,
                        },
                        type: "absolute",
                        width: 4000,
                    },
                }),
                new Paragraph({
                    border: {
                        bottom: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        left: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        right: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        top: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                    },
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                        }),
                    ],
                    frame: {
                        alignment: {
                            x: HorizontalPositionAlign.CENTER,
                            y: VerticalPositionAlign.TOP,
                        },
                        anchor: {
                            horizontal: FrameAnchorType.MARGIN,
                            vertical: FrameAnchorType.MARGIN,
                        },
                        height: 1000,
                        type: "alignment",
                        width: 4000,
                    },
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    border: {
                        bottom: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        left: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        right: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        top: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                    },
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                        }),
                    ],
                    frame: {
                        alignment: {
                            x: HorizontalPositionAlign.CENTER,
                            y: VerticalPositionAlign.BOTTOM,
                        },
                        anchor: {
                            horizontal: FrameAnchorType.MARGIN,
                            vertical: FrameAnchorType.MARGIN,
                        },
                        height: 1000,
                        type: "alignment",
                        width: 4000,
                    },
                }),
            ],
            properties: {},
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
