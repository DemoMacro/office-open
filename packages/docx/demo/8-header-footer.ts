import { writeFileSync } from "node:fs";
// Add text to header and footer

import { AlignmentType, LevelFormat, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  numbering: {
    config: [
      {
        levels: [
          {
            alignment: AlignmentType.START,
            format: LevelFormat.DECIMAL,
            level: 0,
            style: {
              paragraph: {
                indent: {
                  hanging: "0.18in",
                  left: "0.5in",
                },
              },
            },
            text: "%1.",
          },
        ],
        reference: "footer-numbering",
      },
    ],
  },
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      footers: {
        default: [
          { paragraph: "This footer contains a numbered list:" },
          {
            paragraph: {
              numbering: {
                level: 0,
                reference: "footer-numbering",
              },
              text: "First item in the list",
            },
          },
          {
            paragraph: {
              numbering: {
                level: 0,
                reference: "footer-numbering",
              },
              text: "Second item in the list",
            },
          },
          {
            paragraph: {
              numbering: {
                level: 0,
                reference: "footer-numbering",
              },
              text: "Third item in the list",
            },
          },
        ],
      },
      headers: {
        default: [{ paragraph: "Header text" }],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
