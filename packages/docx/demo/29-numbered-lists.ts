import { writeFileSync } from "node:fs";
// Numbered lists
// The lists can also be restarted by specifying the instance number

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
        ],
        reference: "my-crazy-reference",
      },
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
            text: "%1",
          },
        ],
        reference: "my-number-numbering-reference",
      },
      {
        levels: [
          {
            alignment: AlignmentType.START,
            format: LevelFormat.DECIMAL_ZERO,
            level: 0,
            style: {
              paragraph: {
                indent: {
                  hanging: "0.18in",
                  left: "0.5in",
                },
              },
            },
            text: "[%1]",
          },
        ],
        reference: "padded-numbering-reference",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            contextualSpacing: true,
            numbering: {
              level: 0,
              reference: "my-crazy-reference",
            },
            spacing: {
              before: 200,
            },
            text: "line with contextual spacing",
          },
        },
        {
          paragraph: {
            contextualSpacing: true,
            numbering: {
              level: 0,
              reference: "my-crazy-reference",
            },
            spacing: {
              before: 200,
            },
            text: "line with contextual spacing",
          },
        },
        {
          paragraph: {
            contextualSpacing: false,
            numbering: {
              level: 0,
              reference: "my-crazy-reference",
            },
            spacing: {
              before: 200,
            },
            text: "line without contextual spacing",
          },
        },
        {
          paragraph: {
            contextualSpacing: false,
            numbering: {
              level: 0,
              reference: "my-crazy-reference",
            },
            spacing: {
              before: 200,
            },
            text: "line without contextual spacing",
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
              level: 0,
              reference: "my-number-numbering-reference",
            },
            text: "Step 3 - Put in oven",
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Next",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 2,
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 2,
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Next",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 3,
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 3,
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              instance: 3,
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Next",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "padded-numbering-reference",
            },
            text: "test",
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
