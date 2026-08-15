// Change background colour of whole document

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  background: {
    color: "C45911",
  },
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
      properties: {},
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/56-background-color.docx", buffer);
