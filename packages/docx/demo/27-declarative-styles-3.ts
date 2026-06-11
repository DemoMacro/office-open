import { writeFileSync } from "node:fs";
// Custom styles using JavaScript configuration

import { HeadingLevel, UnderlineType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            style: "myWonkyStyle",
            text: "Hello",
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "World",
          },
        },
      ],
    },
  ],
  styles: {
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "myWonkyStyle",
        name: "My Wonky Style",
        next: "Normal",
        paragraph: {
          indent: {
            left: "0.5in",
          },
          spacing: {
            line: 276,
          },
        },
        run: {
          color: "990000",
          italic: true,
        },
      },
      {
        basedOn: "Normal",
        id: "Heading2",
        name: "Heading 2",
        next: "Normal",
        paragraph: {
          spacing: {
            after: 120,
            before: 240,
          },
        },
        quickFormat: true,
        run: {
          bold: true,
          size: 13,
          underline: {
            color: "FF0000",
            type: UnderlineType.DOUBLE,
          },
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
