// Simple example to add text to a document

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

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
                size: 20,
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
writeFileSync(".temp/1-basic.docx", buffer);
