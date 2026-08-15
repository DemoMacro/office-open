// Example of making content of section vertically aligned

import { mkdirSync, writeFileSync } from "node:fs";

import { VerticalAlignSection, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
              {
                bold: true,
                children: [{ tab: true }, "Github is the best"],
              },
            ],
          },
        },
      ],
      properties: {
        verticalAlign: VerticalAlignSection.CENTER,
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/48-vertical-align.docx", buffer);
