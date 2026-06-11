import { writeFileSync } from "node:fs";
// Numbered lists - Add parent number in sub number

import { AlignmentType, HeadingLevel, LevelFormat, generateDocument } from "@office-open/docx";

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
                indent: { hanging: 260, left: "0.5in" },
              },
            },
            text: "%1",
          },
          {
            alignment: AlignmentType.START,
            format: LevelFormat.DECIMAL,
            level: 1,
            style: {
              paragraph: {
                indent: {
                  hanging: 1.25 * 260,
                  left: 1.25 * 720, // 0.5in = 720 twips
                },
              },
              run: {
                bold: true,
                font: "Times New Roman",
                size: 9,
              },
            },
            text: "%1.%2",
          },
        ],
        reference: "my-number-numbering-reference",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "How to make cake",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-number-numbering-reference",
            },
            text: "Step 1 - Add sugar",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-number-numbering-reference",
            },
            text: "Step 2 - Add wheat",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-number-numbering-reference",
            },
            text: "Step 2a - Stir the wheat in a circle",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-number-numbering-reference",
            },
            text: "Step 3 - Put in oven",
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "How to make cake",
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
