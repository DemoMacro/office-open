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
            start: 10,
            text: "%1",
          },
        ],
        reference: "ref1",
      },
      {
        levels: [
          {
            format: LevelFormat.DECIMAL,
            level: 0,
            text: "%1",
          },
        ],
        reference: "ref2",
      },
    ],
  },
  sections: [
    {
      children: [
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
              instance: 1,
              level: 0,
              reference: "ref2",
            },
            text: "REF2 - inst:0 - lvl:0",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 1,
              level: 0,
              reference: "ref2",
            },
            text: "REF2 - inst:0 - lvl:0",
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
