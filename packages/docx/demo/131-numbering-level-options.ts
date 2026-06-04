// Numbering with level restart, legacy spacing/indent, and abstract numbering options.

import * as fs from "fs";

import {
  AlignmentType,
  Document,
  LevelFormat,
  Packer,
  Paragraph,
  TextRun,
} from "@office-open/docx";

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "decimal-with-restart",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            start: 1,
            // Restart numbering at level 1
            lvlRestart: 0,
            // Legacy compatibility spacing and indent
            legacy: { space: 720, indent: 360 },
            style: {
              paragraph: {
                indent: { left: 720, hanging: 360 },
              },
            },
          },
          {
            level: 1,
            format: LevelFormat.LOWER_LETTER,
            text: "%2)",
            alignment: AlignmentType.LEFT,
            start: 1,
            lvlRestart: 0,
            legacy: { space: 360, indent: 180 },
            style: {
              paragraph: {
                indent: { left: 1440, hanging: 360 },
              },
            },
          },
        ],
      },
    ],
  },

  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Numbering with Level Restart and Legacy Options",
              size: 32,
            }),
          ],
        }),

        new Paragraph({
          numbering: { reference: "decimal-with-restart", level: 0 },
          children: [new TextRun("First item")],
        }),
        new Paragraph({
          numbering: { reference: "decimal-with-restart", level: 1 },
          children: [new TextRun("Sub-item a")],
        }),
        new Paragraph({
          numbering: { reference: "decimal-with-restart", level: 1 },
          children: [new TextRun("Sub-item b")],
        }),
        new Paragraph({
          numbering: { reference: "decimal-with-restart", level: 0 },
          children: [new TextRun("Second item")],
        }),
        new Paragraph({
          numbering: { reference: "decimal-with-restart", level: 1 },
          children: [new TextRun("Sub-item a (restarted)")],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
