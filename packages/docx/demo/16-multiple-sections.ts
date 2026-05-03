// Multiple sections and headers

import * as fs from "fs";

import {
    Document,
    Footer,
    Header,
    NumberFormat,
    Packer,
    PageNumber,
    PageOrientation,
    Paragraph,
    TextRun,
} from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
        },
        {
            children: [new Paragraph("hello")],
            footers: {
                default: new Footer({
                    children: [new Paragraph("Footer on another page")],
                }),
            },
            headers: {
                default: new Header({
                    children: [new Paragraph("First Default Header on another page")],
                }),
            },

            properties: {
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.DECIMAL,
                        start: 1,
                    },
                },
            },
        },
        {
            children: [new Paragraph("hello in landscape")],
            footers: {
                default: new Footer({
                    children: [new Paragraph("Footer on another page")],
                }),
            },
            headers: {
                default: new Header({
                    children: [new Paragraph("Second Default Header on another page")],
                }),
            },
            properties: {
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.DECIMAL,
                        start: 1,
                    },
                    size: {
                        orientation: PageOrientation.LANDSCAPE,
                    },
                },
            },
        },
        {
            children: [
                new Paragraph(
                    "Page number in the header must be 2, because it continues from the previous section.",
                ),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    children: ["Page number: ", PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },

            properties: {
                page: {
                    size: {
                        orientation: PageOrientation.PORTRAIT,
                    },
                },
            },
        },
        {
            children: [
                new Paragraph(
                    "Page number in the header must be III, because it continues from the previous section, but is defined as upper roman.",
                ),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    children: ["Page number: ", PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            properties: {
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.UPPER_ROMAN,
                    },
                    size: {
                        orientation: PageOrientation.PORTRAIT,
                    },
                },
            },
        },
        {
            children: [
                new Paragraph(
                    "Page number in the header must be 25, because it is defined to start at 25 and to be decimal in this section.",
                ),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    children: ["Page number: ", PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            properties: {
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.DECIMAL,
                        start: 25,
                    },
                    size: {
                        orientation: PageOrientation.PORTRAIT,
                    },
                },
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
