// Highlighting text

import { writeFileSync } from "node:fs";

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
                  size: 24,
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
writeFileSync("My Document.docx", buffer);
