// Glossary document: building blocks that appear in Word's Quick Parts gallery.
// Open the document and go to Insert > Quick Parts to see the entries.

import * as fs from "fs";

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  DocPartGallery,
  DocPartBehavior,
  DocPartType,
} from "@office-open/docx";

const doc = new Document({
  glossary: {
    parts: [
      {
        name: "Company Disclaimer",
        gallery: DocPartGallery.AUTO_TEXT,
        category: "Legal",
        behaviors: [DocPartBehavior.PARAGRAPH],
        description: "Standard company disclaimer text",
        guid: "{11111111-2222-3333-4444-555555555555}",
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "CONFIDENTIAL: This document contains proprietary information.",
                italics: true,
                color: "808080",
              }),
            ],
          }),
        ],
      },
      {
        name: "Meeting Header",
        gallery: DocPartGallery.HEADERS,
        category: "Corporate",
        types: [DocPartType.NORMAL],
        guid: "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}",
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "Meeting Notes — ", bold: true }),
              new TextRun({ text: "Date: [Insert Date]" }),
            ],
          }),
        ],
      },
      {
        name: "Custom Footer",
        gallery: DocPartGallery.FOOTERS,
        category: "Corporate",
        behaviors: [DocPartBehavior.CONTENT],
        guid: "{FFFFFFFF-0000-1111-2222-333333333333}",
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "Page ", size: 18 }),
              new TextRun({ text: "1", size: 18 }),
              new TextRun({ text: " of ", size: 18 }),
              new TextRun({ text: "[Total]", size: 18 }),
            ],
          }),
        ],
      },
    ],
  },
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun({ text: "Glossary Document Demo", bold: true, size: 32 })],
        }),
        new Paragraph({
          children: [new TextRun("This document contains Quick Parts building blocks.")],
        }),
        new Paragraph({
          children: [
            new TextRun("Open in Word and go to Insert > Quick Parts to see the glossary entries."),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
