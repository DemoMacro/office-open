// Numbered lists - With complex number text

import { writeFileSync } from "node:fs";

import { LevelFormat, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  numbering: {
    config: [
      {
        levels: [
          {
            format: LevelFormat.DECIMAL,
            level: 0,
            text: "%1",
          },
          {
            format: LevelFormat.DECIMAL,
            level: 1,
            text: "%1.%2",
          },
          {
            format: LevelFormat.DECIMAL,
            level: 2,
            text: "%1.%2.%3",
          },
        ],
        reference: "ref1",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "ref1",
            },
            text: "REF1 - lvl:1",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 2,
              reference: "ref1",
            },
            text: "REF1  - lvl:2",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - lvl:0",
          },
        },
        {
          paragraph: {
            text: "Random text",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 1,
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - inst:1 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 0,
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - inst:0 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 0,
              level: 0,
              reference: "ref1",
            },
            text: "REF1 - inst:0 - lvl:0",
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
