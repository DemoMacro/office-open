// Demo: Move Revisions (Track Changes - Move operations)
// Demonstrates MovedFromTextRun, MovedToTextRun, and move range markers.

import * as fs from "fs";

import {
    Document,
    MovedFromTextRun,
    MovedToTextRun,
    MoveFromRangeEnd,
    MoveFromRangeStart,
    MoveToRangeEnd,
    MoveToRangeStart,
    Packer,
    Paragraph,
    TextRun,
} from "@office-open/docx";

const REVISION_AUTHOR = "Firstname Lastname";
const REVISION_DATE = "2026-05-02T10:00:00Z";

const doc = new Document({
    features: {
        trackRevisions: true,
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Move Revisions Demo", bold: true, size: 32 })],
                    spacing: { after: 400 },
                }),

                // 1. Simple move: text moved from one location to another
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "1. Simple Text Move", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("This paragraph had text moved "),
                        new MovedFromTextRun({
                            id: 0,
                            author: REVISION_AUTHOR,
                            date: REVISION_DATE,
                            text: "away from here",
                        }),
                        new TextRun(" to the next paragraph."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("And the text appeared here: "),
                        new MovedToTextRun({
                            id: 1,
                            author: REVISION_AUTHOR,
                            date: REVISION_DATE,
                            text: "away from here",
                        }),
                        new TextRun("."),
                    ],
                }),

                new Paragraph({ text: "" }),

                // 2. Move with formatting preserved
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "2. Move with Formatting", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Source paragraph: "),
                        new MovedFromTextRun({
                            id: 2,
                            author: REVISION_AUTHOR,
                            date: REVISION_DATE,
                            text: "bold moved text",
                            bold: true,
                            color: "FF0000",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Destination paragraph: "),
                        new MovedToTextRun({
                            id: 3,
                            author: REVISION_AUTHOR,
                            date: REVISION_DATE,
                            text: "bold moved text",
                            bold: true,
                            color: "FF0000",
                        }),
                    ],
                }),

                new Paragraph({ text: "" }),

                // 3. Move range markers (standalone bookmarks for move regions)
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "3. Move Range Markers", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("This paragraph contains "),
                        new MoveFromRangeStart(
                            4,
                            "moved-paragraph",
                            REVISION_AUTHOR,
                            REVISION_DATE,
                        ),
                        new TextRun("content that was moved"),
                        new MoveFromRangeEnd(4),
                        new TextRun(" elsewhere."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Destination with "),
                        new MoveToRangeStart(5, "moved-paragraph", REVISION_AUTHOR, REVISION_DATE),
                        new TextRun("the moved content"),
                        new MoveToRangeEnd(5),
                        new TextRun(" placed here."),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
