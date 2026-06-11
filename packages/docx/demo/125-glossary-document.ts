// Glossary document: building blocks that appear in Word's Quick Parts gallery.
// Open the document and go to Insert > Quick Parts to see the entries.

import { writeFileSync } from "node:fs";

import { DocPartGallery, DocPartBehavior, DocPartType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
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
          {
            paragraph: {
              children: [
                {
                  text: "CONFIDENTIAL: This document contains proprietary information.",
                  italic: true,
                  color: "808080",
                },
              ],
            },
          },
        ],
      },
      {
        name: "Meeting Header",
        gallery: DocPartGallery.HEADERS,
        category: "Corporate",
        types: [DocPartType.NORMAL],
        guid: "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}",
        children: [
          {
            paragraph: {
              children: [{ text: "Meeting Notes — ", bold: true }, { text: "Date: [Insert Date]" }],
            },
          },
        ],
      },
      {
        name: "Custom Footer",
        gallery: DocPartGallery.FOOTERS,
        category: "Corporate",
        behaviors: [DocPartBehavior.CONTENT],
        guid: "{FFFFFFFF-0000-1111-2222-333333333333}",
        children: [
          {
            paragraph: {
              children: [
                { text: "Page ", size: 9 },
                { text: "1", size: 9 },
                { text: " of ", size: 9 },
                { text: "[Total]", size: 9 },
              ],
            },
          },
        ],
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Glossary Document Demo", bold: true, size: 16 }],
          },
        },
        {
          paragraph: {
            children: ["This document contains Quick Parts building blocks."],
          },
        },
        {
          paragraph: {
            children: ["Open in Word and go to Insert > Quick Parts to see the glossary entries."],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
