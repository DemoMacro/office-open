// Break with clear attribute: controls how text reflows after a line break

import * as fs from "fs";

import { createBreak, Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Line break with clear attribute:"),
                        new TextRun({
                            children: [
                                "First line",
                                createBreak(),
                                "Second line (default, no clear)",
                                createBreak({ clear: "left" }),
                                "Third line (clear left)",
                                createBreak({ clear: "right" }),
                                "Fourth line (clear right)",
                                createBreak({ clear: "all" }),
                                "Fifth line (clear all)",
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
