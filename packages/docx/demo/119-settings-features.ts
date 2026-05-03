// Document settings features: view, zoom, write protection, display background shape,
// font embedding, document variables

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    background: {
        color: "C45911",
    },
    displayBackgroundShape: true,
    view: "print",
    zoom: {
        percent: 150,
    },
    embedTrueTypeFonts: true,
    saveSubsetFonts: true,
    docVars: [
        { name: "Title", val: "Settings Demo" },
        { name: "Version", val: "1.0" },
        { name: "Author", val: "Test User" },
    ],
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun("Document Settings Demo")],
                }),
                new Paragraph({
                    children: [
                        new TextRun("This document opens in Print Layout view at 150% zoom."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(
                            "The background color is displayed in print layout because displayBackgroundShape is enabled.",
                        ),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(
                            "TrueType fonts are embedded, and only used subsets are saved.",
                        ),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(
                            "Document variables (Title, Version, Author) are stored in settings.xml.",
                        ),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
