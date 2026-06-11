import { writeFileSync } from "node:fs";
// Example on how to customize the look at feel using Styles

import {
  AlignmentType,
  HeadingLevel,
  LevelFormat,
  UnderlineType,
  generateDocument,
} from "@office-open/docx";

const buffer = await generateDocument({
  creator: "Clippy",
  description: "A brief example of using docx",
  numbering: {
    config: [
      {
        levels: [
          {
            level: 0,
            format: LevelFormat.LOWER_LETTER,
            text: "%1)",
            alignment: AlignmentType.LEFT,
          },
        ],
        reference: "my-crazy-numbering",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "Test heading1, bold and italicized",
          },
        },
        { paragraph: "Some simple content" },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Test heading2 with double red underline",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-crazy-numbering",
            },
            style: "aside",
            text: "Option1",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-crazy-numbering",
            },
            text: "Option5 -- override 2 to 5",
          },
        },
        {
          paragraph: {
            numbering: {
              level: 0,
              reference: "my-crazy-numbering",
            },
            text: "Option3",
          },
        },
        {
          paragraph: {
            children: [
              {
                font: {
                  name: "Monospace",
                },
                text: "Some monospaced content",
              },
            ],
          },
        },
        {
          paragraph: {
            style: "aside",
            text: "An aside, in light gray italics and indented",
          },
        },
        {
          paragraph: {
            style: "wellSpaced",
            text: "This is normal, but well-spaced text",
          },
        },
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "This is a bold run,",
              },
              " switching to normal ",
              {
                text: "and then underlined ",
                underline: {},
              },
              {
                emphasisMark: {},
                text: "and then emphasis-mark ",
              },
              {
                text: "and back to normal.",
              },
              {
                text: "This text will be invisible!",
                vanish: true,
              },
              {
                specVanish: true,
                text: "This text will be VERY invisible! Word processors cannot override this!",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Strong Style",
              },
              {
                text: " - Very strong.",
              },
            ],
            style: "Strong",
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Underline and Strike",
              },
              {
                text: " Override Underline ",
                underline: {
                  type: UnderlineType.NONE,
                },
              },
              {
                text: "Strike and Underline",
              },
            ],
            style: "strikeUnderline",
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Hello World ",
              },
              {
                style: "strikeUnderlineCharacter",
                text: "Underline and Strike",
              },
              {
                text: " Another Hello World",
              },
              {
                scale: 50,
                text: " Scaled text",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Scaled paragraph",
              },
            ],
          },
        },
      ],
    },
  ],
  styles: {
    characterStyles: [
      {
        basedOn: "Normal",
        id: "strikeUnderlineCharacter",
        name: "Strike Underline",
        quickFormat: true,
        run: {
          strike: true,
          underline: {
            type: UnderlineType.SINGLE,
          },
        },
      },
    ],
    default: {
      document: {
        paragraph: {
          alignment: AlignmentType.RIGHT,
        },
        run: {
          font: "Calibri",
          size: 22,
        },
      },
      heading1: {
        paragraph: {
          spacing: {
            after: 120,
          },
        },
        run: {
          bold: true,
          color: "FF0000",
          italics: true,
          size: 28,
        },
      },
      heading2: {
        paragraph: {
          spacing: {
            after: 120,
            before: 240,
          },
        },
        run: {
          bold: true,
          size: 26,
          underline: {
            color: "FF0000",
            type: UnderlineType.DOUBLE,
          },
        },
      },
      listParagraph: {
        run: {
          color: "#FF0000",
        },
      },
    },
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "aside",
        name: "Aside",
        next: "Normal",
        paragraph: {
          indent: {
            left: "0.5in",
          },
          spacing: {
            line: 276,
          },
        },
        run: {
          color: "999999",
          italics: true,
        },
      },
      {
        basedOn: "Normal",
        id: "wellSpaced",
        name: "Well Spaced",
        paragraph: {
          spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
        },
        quickFormat: true,
      },
      {
        basedOn: "Normal",
        id: "strikeUnderline",
        name: "Strike Underline",
        quickFormat: true,
        run: {
          strike: true,
          underline: {
            type: UnderlineType.SINGLE,
          },
        },
      },
    ],
  },
  title: "Sample Document",
});
writeFileSync("My Document.docx", buffer);
