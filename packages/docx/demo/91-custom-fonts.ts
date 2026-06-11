// Simple example to add text to a document

import { readFileSync, writeFileSync } from "node:fs";

import { CharacterSet, generateDocument } from "@office-open/docx";

const font = readFileSync("./demo/assets/Pacifico.ttf");

const buffer = await generateDocument({
  fonts: [{ characterSet: CharacterSet.ANSI, data: font, name: "Pacifico" }],
  sections: [
    {
      children: [
        {
          paragraph: {
            run: {
              font: "Pacifico",
            },
            children: [
              "Hello World",
              {
                text: "Foo Bar",
                bold: true,
                size: 40,
                font: "Pacifico",
              },
              {
                bold: true,
                font: "Pacifico",
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
writeFileSync("My Document.docx", buffer);
