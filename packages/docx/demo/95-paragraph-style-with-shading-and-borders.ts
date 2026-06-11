import { writeFileSync } from "node:fs";

import { BorderStyle, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
              },
            ],
            style: "withSingleBlackBordersAndYellowShading",
          },
        },
      ],
    },
  ],
  styles: {
    paragraphStyles: [
      {
        basedOn: "Normal",
        id: "withSingleBlackBordersAndYellowShading",
        name: "Paragraph Style with Black Borders and Yellow Shading",
        paragraph: {
          border: {
            bottom: {
              color: "#000000",
              size: 4,
              style: BorderStyle.SINGLE,
            },
            left: {
              color: "#000000",
              size: 4,
              style: BorderStyle.SINGLE,
            },
            right: {
              color: "#000000",
              size: 4,
              style: BorderStyle.SINGLE,
            },
            top: {
              color: "#000000",
              size: 4,
              style: BorderStyle.SINGLE,
            },
          },
          shading: {
            color: "#fff000",
            type: "solid",
          },
        },
      },
    ],
  },
});
writeFileSync("My Document.docx", buffer);
