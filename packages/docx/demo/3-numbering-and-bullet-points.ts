import { writeFileSync } from "node:fs";
// Numbering and bullet points example

import { AlignmentType, HeadingLevel, LevelFormat, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  numbering: {
    config: [
      {
        levels: [
          {
            alignment: AlignmentType.START,
            format: LevelFormat.UPPER_ROMAN,
            level: 0,
            style: {
              paragraph: {
                indent: {
                  hanging: "0.18in",
                  left: "0.5in",
                },
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
                  hanging: "0.68in",
                  left: "1in",
                },
              },
            },
            text: "%2.",
          },
          {
            alignment: AlignmentType.START,
            format: LevelFormat.LOWER_LETTER,
            level: 2,
            style: {
              paragraph: {
                indent: {
                  hanging: "1.18in",
                  left: "1.5in",
                },
              },
            },
            text: "%3)",
          },
          {
            alignment: AlignmentType.START,
            format: LevelFormat.UPPER_LETTER,
            level: 3,
            style: {
              paragraph: {
                indent: { hanging: 2420, left: 2880 },
              },
            },
            text: "%4)",
          },
        ],
        reference: "my-crazy-numbering",
      },
      {
        levels: [
          {
            alignment: AlignmentType.LEFT,
            format: LevelFormat.BULLET,
            level: 0,
            style: {
              paragraph: {
                indent: {
                  hanging: "0.25in",
                  left: "0.5in",
                },
              },
            },
            text: "ὠ",
          },
          {
            alignment: AlignmentType.LEFT,
            format: LevelFormat.BULLET,
            level: 1,
            style: {
              paragraph: {
                indent: {
                  hanging: "0.25in",
                  left: "1in",
                },
              },
            },
            text: "¥",
          },
          {
            alignment: AlignmentType.LEFT,
            format: LevelFormat.BULLET,
            level: 2,
            style: {
              paragraph: {
                indent: { hanging: "0.25in", left: 2160 },
              },
            },
            text: "✿",
          },
          {
            alignment: AlignmentType.LEFT,
            format: LevelFormat.BULLET,
            level: 3,
            style: {
              paragraph: {
                indent: { hanging: "0.25in", left: 2880 },
              },
            },
            text: "♺",
          },
          {
            alignment: AlignmentType.LEFT,
            format: LevelFormat.BULLET,
            level: 4,
            style: {
              paragraph: {
                indent: { hanging: "0.25in", left: 3600 },
              },
            },
            text: "☃",
          },
        ],
        reference: "my-unique-bullet-points",
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
              reference: "my-crazy-numbering",
            },
            text: "Hey you",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-crazy-numbering",
            },
            text: "What's up fam",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-crazy-numbering",
            },
            text: "Hello World 2",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 2,
              reference: "my-crazy-numbering",
            },
            text: "Yeah boi",
          },
        },
        {
          paragraph: {
            bullet: {
              level: 0,
            },
            text: "Hey you",
          },
        },
        {
          paragraph: {
            bullet: {
              level: 1,
            },
            text: "What's up fam",
          },
        },
        {
          paragraph: {
            bullet: {
              level: 2,
            },
            text: "Hello World 2",
          },
        },
        {
          paragraph: {
            bullet: {
              level: 3,
            },
            text: "Yeah boi",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 3,
              reference: "my-crazy-numbering",
            },
            text: "101 MSXFM",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-crazy-numbering",
            },
            text: "back to level 1",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-crazy-numbering",
            },
            text: "back to level 0",
          },
        },

        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "Custom Bullet points",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-unique-bullet-points",
            },
            text: "What's up fam",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-unique-bullet-points",
            },
            text: "Hey you",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-unique-bullet-points",
            },
            text: "What's up fam",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 2,
              reference: "my-unique-bullet-points",
            },
            text: "Hello World 2",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 3,
              reference: "my-unique-bullet-points",
            },
            text: "Yeah boi",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 4,
              reference: "my-unique-bullet-points",
            },
            text: "my Awesome numbering",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 1,
              reference: "my-unique-bullet-points",
            },
            text: "Back to level 1",
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              numbering: {
                level: 0,
                reference: "my-crazy-numbering",
              },
              text: "Hey you",
            },
          },
          {
            paragraph: {
              numbering: {
                level: 1,
                reference: "my-crazy-numbering",
              },
              text: "What's up fam",
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              numbering: {
                level: 0,
                reference: "my-crazy-numbering",
              },
              text: "Hey you",
            },
          },
          {
            paragraph: {
              numbering: {
                level: 1,
                reference: "my-crazy-numbering",
              },
              text: "What's up fam",
            },
          },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
