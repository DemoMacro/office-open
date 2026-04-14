// Multiple sections with total number of pages in each section

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    NumberFormat,
    Packer,
    PageBreak,
    PageNumber,
    Paragraph,
    TextRun,
} from "docx";

const header = new Header({
    children: [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun("Header on another page"),
                new TextRun({
                    children: ["Page number: ", PageNumber.CURRENT],
                }),
                new TextRun({
                    children: [" to ", PageNumber.TOTAL_PAGES_IN_SECTION],
                }),
            ],
        }),
    ],
});

const footer = new Footer({
    children: [new Paragraph("Foo Bar corp. ")],
});

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Section 1"),
                        new PageBreak(),
                        new TextRun("Section 1"),
                        new PageBreak(),
                    ],
                }),
            ],
            footers: {
                default: footer,
            },
            headers: {
                default: header,
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
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Section 2"),
                        new PageBreak(),
                        new TextRun("Section 2"),
                        new PageBreak(),
                    ],
                }),
            ],
            footers: {
                default: footer,
            },
            headers: {
                default: header,
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
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
