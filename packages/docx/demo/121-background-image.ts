// Document background image: VML v:background + v:fill

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  background: {
    color: "E8E8E8",
    image: {
      data: readFileSync("./demo/images/dog.png"),
      type: "png",
    },
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: "This document has a background image.",
                bold: true,
                size: 14,
              },
            ],
          },
        },
        {
          paragraph: {
            children: ["The image is rendered as a full-page background using VML v:fill."],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/121-background-image.docx", buffer);
