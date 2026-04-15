// Page numbers

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    Packer,
    PageBreak,
    PageNumber,
    Paragraph,
    TextRun,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun("First Page"), new PageBreak()],
                }),
                new Paragraph("Second Page"),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun("My Title "),
                                new TextRun({
                                    children: [
                                        "Footer - Page ",
                                        PageNumber.CURRENT,
                                        " of ",
                                        PageNumber.TOTAL_PAGES,
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun("First Page Footer "),
                                new TextRun({
                                    children: ["Page ", PageNumber.CURRENT],
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
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun("My Title "),
                                new TextRun({
                                    children: ["Page ", PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
                first: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun("First Page Header "),
                                new TextRun({
                                    children: ["Page ", PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            properties: {
                titlePage: true,
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
