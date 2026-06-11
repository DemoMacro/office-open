import { readFileSync, writeFileSync } from "node:fs";
// Footnotes

import { AlignmentType, LevelFormat, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  footnotes: {
    1: { children: ["Foo", "Bar"] },
    2: {
      children: [
        "This footnote contains a numbered list:",
        {
          numbering: {
            level: 0,
            reference: "footnote-numbering",
          },
          text: "First item in the list",
        },
        {
          numbering: {
            level: 0,
            reference: "footnote-numbering",
          },
          text: "Second item in the list",
        },
        {
          numbering: {
            level: 0,
            reference: "footnote-numbering",
          },
          text: "Third item in the list",
        },
      ],
    },
    3: {
      children: [
        {
          children: [
            {
              image: {
                data: readFileSync("./demo/images/cat.jpg"),
                transformation: {
                  height: 100,
                  width: 100,
                },
                type: "jpg",
              },
            },
            {
              text: "It's a cat",
            },
          ],
        },
      ],
    },
    4: { children: ["Foo1"] },
    5: { children: ["Test1"] },
    6: { children: ["My amazing reference1"] },
  },
  numbering: {
    config: [
      {
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.START,
            style: {
              paragraph: {
                indent: {
                  left: "0.5in",
                  hanging: "0.18in",
                },
              },
            },
          },
        ],
        reference: "footnote-numbering",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                children: ["Hello"],
              },
              { footnoteReference: 1 },
              {
                children: [" World!"],
              },
              { footnoteReference: 2 },
              {
                children: [" GitHub!"],
              },
            ],
          },
        },
        {
          paragraph: {
            children: ["Hello World", { footnoteReference: 3 }],
          },
        },
      ],
      properties: {
        footnotePr: {
          numRestart: "eachSect",
          pos: "beneathText",
        },
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [
              {
                children: ["Hello"],
              },
              { footnoteReference: 4 },
              {
                children: [" World!"],
              },
              { footnoteReference: 5 },
            ],
          },
        },
        {
          paragraph: {
            children: ["Hello World Again", { footnoteReference: 6 }],
          },
        },
      ],
      properties: {
        footnotePr: {
          numRestart: "continuous",
          pos: "sectEnd",
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
