// East Asian layout - Need to use an East Asian font
import { writeFileSync } from "node:fs";

import { HeadingLevel, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "East Asian Layout",
          },
        },

        // East Asian layout - combined characters with round brackets
        {
          paragraph: {
            children: [
              "Combined characters (round brackets): ",
              {
                eastAsianLayout: {
                  combine: true,
                  combineBrackets: "round",
                  id: 1,
                },
                text: "国民",
              },
            ],
            spacing: { after: 200 },
          },
        },

        // East Asian layout - combined characters with square brackets
        {
          paragraph: {
            children: [
              "Combined characters (square brackets): ",
              {
                eastAsianLayout: {
                  combine: true,
                  combineBrackets: "square",
                  id: 2,
                },
                text: "日本語",
              },
            ],
            spacing: { after: 200 },
          },
        },

        // East Asian layout - vertical text
        {
          paragraph: {
            children: [
              "Vertical text: ",
              {
                eastAsianLayout: {
                  vert: true,
                },
                text: "縦書き",
              },
            ],
            spacing: { after: 200 },
          },
        },
      ],
    },
  ],
  styles: {
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "Normal",
        name: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: "MS Gothic",
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
