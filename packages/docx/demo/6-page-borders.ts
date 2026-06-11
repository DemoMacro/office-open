// Example of how to change page borders

import { writeFileSync } from "node:fs";

import { generateDocument, HeadingLevel } from "@office-open/docx";

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
                text: "Foo bar",
              },
              {
                bold: true,
                children: [{ tab: true }, "Github is the best"],
              },
            ],
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "Hello World",
          },
        },
        { paragraph: "Foo bar" },
        { paragraph: "Github is the best" },
      ],
      properties: {
        page: {
          margin: {
            bottom: 0,
            left: 0,
            right: 0,
            top: 0,
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
