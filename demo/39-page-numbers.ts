// Example how to display page numbers

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

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun("Hello World 1"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 2"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 3"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 4"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 5"), new PageBreak()],
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun("Foo Bar corp. "),
                                new TextRun({
                                    children: ["Page Number: ", PageNumber.CURRENT],
                                }),
                                new TextRun({
                                    children: [" to ", PageNumber.TOTAL_PAGES],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun("Foo Bar corp. "),
                                new TextRun({
                                    children: ["Page Number ", PageNumber.CURRENT],
                                }),
                                new TextRun({
                                    children: [" to ", PageNumber.TOTAL_PAGES],
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
