// Example of using tab stops

import * as fs from "fs";

import {
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    TabStopPosition,
    TabStopType,
    TextRun,
} from "docx";

const columnWidth = TabStopPosition.MAX / 4;
const receiptTabStops = [
    // No need to define first left tab column
    // The right aligned tab column position should point to the end of column
    // I.e. in this case
    // (end position of 1st) + (end position of current)
    // ColumnWidth + columnWidth = columnWidth * 2

    { position: columnWidth * 2, type: TabStopType.RIGHT },
    { position: columnWidth * 3, type: TabStopType.RIGHT },
    { position: TabStopPosition.MAX, type: TabStopType.RIGHT },
];
const twoTabStops = [{ position: TabStopPosition.MAX, type: TabStopType.RIGHT }];

const doc = new Document({
    defaultTabStop: 0,
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun("Receipt 001")],
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "To Bob.\tBy Alice.",
                            bold: true,
                        }),
                    ],
                    tabStops: twoTabStops,
                }),
                new Paragraph({
                    children: [new TextRun("Foo Inc\tBar Inc")],
                    tabStops: twoTabStops,
                }),
                new Paragraph({ text: "" }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Item\tPrice\tQuantity\tSub-total",
                            bold: true,
                        }),
                    ],

                    tabStops: receiptTabStops,
                }),
                new Paragraph({
                    tabStops: receiptTabStops,
                    text: "Item 3\t10\t5\t50",
                }),
                new Paragraph({
                    tabStops: receiptTabStops,
                    text: "Item 3\t10\t5\t50",
                }),
                new Paragraph({
                    tabStops: receiptTabStops,
                    text: "Item 3\t10\t5\t50",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "\t\t\tTotal: 200",
                            bold: true,
                        }),
                    ],
                    tabStops: receiptTabStops,
                }),
            ],
            properties: {},
        },
    ],
});

const stream = Packer.toStream(doc);
stream.pipe(fs.createWriteStream("My Document.docx"));
