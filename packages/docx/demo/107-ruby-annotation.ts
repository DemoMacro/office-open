// Demo: Ruby annotation (CT_Ruby) - East Asian pronunciation guides
import * as fs from "fs";

import { createRuby, Document, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Ruby Annotation Demo", bold: true, size: 32 })],
                    spacing: { after: 400 },
                }),

                // Japanese furigana
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "Japanese Furigana", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Base text with ruby: "),
                        createRuby({
                            text: "かな",
                            base: "漢字",
                            alignment: "CENTER",
                            fontSize: 20,
                            raise: 20,
                            baseFontSize: 40,
                            languageId: "ja-JP",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Chinese pinyin
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "Chinese Pinyin", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Pinyin guide: "),
                        createRuby({
                            text: "hàn zì",
                            base: "汉字",
                            alignment: "CENTER",
                            fontSize: 18,
                            raise: 20,
                            baseFontSize: 40,
                            languageId: "zh-CN",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Alignment options
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "Alignment Options", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Left: "),
                        createRuby({ text: "left", base: "Align", alignment: "LEFT" }),
                        new TextRun("  Center: "),
                        createRuby({ text: "center", base: "Align", alignment: "CENTER" }),
                        new TextRun("  Right: "),
                        createRuby({ text: "right", base: "Align", alignment: "RIGHT" }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
