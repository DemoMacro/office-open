// Highlighting text

import { mkdirSync, writeFileSync } from "node:fs";

import { AlignmentType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [],
      headers: {
        default: [
          {
            paragraph: {
              alignment: AlignmentType.RIGHT,
              children: [
                {
                  bold: true,
                  color: "FF0000",
                  font: {
                    name: "Garamond",
                  },
                  highlight: "yellow",
                  size: 12,
                  text: "Hello World",
                },
              ],
            },
          },
        ],
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/45-highlighting-text.docx", buffer);
