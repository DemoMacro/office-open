// Example of making content of section vertically aligned

import { writeFileSync } from "node:fs";

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
writeFileSync("My Document.docx", buffer);
