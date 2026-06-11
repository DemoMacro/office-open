// Shading text

import { writeFileSync } from "node:fs";

import { AlignmentType, ShadingType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                emboss: true,
                text: "Embossed text - hello world",
              },
              {
                imprint: true,
                text: "Imprinted text - hello world",
              },
            ],
          },
        },
      ],
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
                  shading: {
                    color: "00FFFF",
                    fill: "FF0000",
                    type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                  },
                  size: 12,
                  text: "Hello World",
                },
              ],
            },
          },
          {
            paragraph: {
              children: [
                {
                  text: "Hello World for entire paragraph",
                },
              ],
              shading: {
                color: "00FFFF",
                fill: "FF0000",
                type: ShadingType.DIAGONAL_CROSS,
              },
            },
          },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
