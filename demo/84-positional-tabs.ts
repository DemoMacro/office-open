// Simple example apply positional tabs to a document

import * as fs from "fs";

import {
    Document,
    Packer,
    Paragraph,
    PositionalTab,
    PositionalTabAlignment,
    PositionalTabLeader,
    PositionalTabRelativeTo,
    TextRun,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Full name"),
                        new TextRun({
                            bold: true,
                            children: [
                                new PositionalTab({
                                    alignment: PositionalTabAlignment.RIGHT,
                                    relativeTo: PositionalTabRelativeTo.MARGIN,
                                    leader: PositionalTabLeader.DOT,
                                }),
                                "John Doe",
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            children: [
                                new PositionalTab({
                                    alignment: PositionalTabAlignment.CENTER,
                                    relativeTo: PositionalTabRelativeTo.INDENT,
                                    leader: PositionalTabLeader.HYPHEN,
                                }),
                                "Foo bar",
                            ],
                        }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
