// Simple example to add text to a document

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const font = readFileSync("./demo/assets/Pacifico.ttf");

const buffer = await generateDocument({
  fonts: [{ characterSet: "00", data: font, name: "Pacifico" }],
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                text: "Foo Bar",
                bold: true,
                size: 20,
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
  styles: {
    default: {
      document: {
        run: {
          font: "Pacifico",
        },
      },
    },
  },
});
writeFileSync("My Document.docx", buffer);
