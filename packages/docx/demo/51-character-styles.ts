// Custom character styles using JavaScript configuration

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                style: "myRedStyle",
                text: "Foo bar",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                style: "strong",
                text: "First Word",
              },
              {
                text: " - Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
              },
            ],
          },
        },
      ],
    },
  ],
  styles: {
    characterStyles: [
      {
        basedOn: "Normal",
        id: "myRedStyle",
        name: "My Wonky Style",
        run: {
          color: "990000",
          italics: true,
        },
      },
      {
        basedOn: "Normal",
        id: "strong",
        name: "Strong",
        run: {
          bold: true,
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
