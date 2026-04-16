// Endnotes

import * as fs from "fs";

import { Document, EndnoteReferenceRun, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    endnotes: {
        1: {
            children: [new Paragraph("This is the first endnote with some detailed explanation.")],
        },
        2: {
            children: [
                new Paragraph("Second endnote"),
                new Paragraph("With multiple paragraphs for more complex content."),
            ],
        },
        3: { children: [new Paragraph("Third endnote referencing important source material.")] },
        4: { children: [new Paragraph("Fourth endnote from a different section.")] },
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            size: 28,
                            text: "Endnotes Demo Document",
                        }),
                    ],
                    spacing: { after: 400 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("This document demonstrates endnotes functionality. "),
                        new TextRun("Here is some text with an endnote reference"),
                        new EndnoteReferenceRun(1),
                        new TextRun(". This allows for detailed citations and references "),
                        new EndnoteReferenceRun(2),
                        new TextRun(" without cluttering the main text flow."),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Endnotes appear at the end of the document, "),
                        new TextRun("unlike footnotes which appear at the bottom of each page"),
                        new EndnoteReferenceRun(3),
                        new TextRun(
                            ". This makes them ideal for academic papers and formal documents.",
                        ),
                    ],
                    spacing: { after: 200 },
                }),
            ],
            properties: {
                endnotePr: {
                    numRestart: "eachSect",
                    pos: "docEnd",
                },
                page: {
                    margin: {
                        bottom: 1440,
                        left: 1440,
                        right: 1440,
                        top: 1440, // 1 inch
                    },
                },
            },
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            size: 24,
                            text: "Second Section",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("This is content from a different section "),
                        new TextRun("with its own endnote reference"),
                        new EndnoteReferenceRun(4),
                        new TextRun(
                            ". Endnotes from all sections appear together at the document end.",
                        ),
                    ],
                }),
            ],
            properties: {
                endnotePr: {
                    numRestart: "continuous",
                    pos: "sectEnd",
                },
            },
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            size: 24,
                            text: "Second Section",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("This is content from a different section "),
                        new TextRun("with its own endnote reference"),
                        new EndnoteReferenceRun(4),
                        new TextRun(
                            ". Endnotes from all sections appear together at the document end.",
                        ),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
