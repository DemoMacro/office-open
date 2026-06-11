// Example demonstrating page borders with style, colors and size

import { writeFileSync } from "node:fs";

import {
  BorderStyle,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
  generateDocument,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: `Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.`,
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: `Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.`,
              },
            ],
          },
        },
      ],
      properties: {
        page: {
          borders: {
            bottom: {
              style: BorderStyle.SINGLE,
              size: 2 * 8, //2pt;
              color: "000000",
            },
            left: {
              style: BorderStyle.SINGLE,
              size: 1 * 8, //1pt;
              color: "000000",
            },
            right: {
              style: BorderStyle.SINGLE,
              size: 1 * 8, //1pt;
              color: "FF00AA",
            },
            top: {
              style: BorderStyle.SINGLE,
              size: 1 * 8, //1pt;
              color: "000000",
            },
            display: PageBorderDisplay.ALL_PAGES,
            offsetFrom: PageBorderOffsetFrom.TEXT,
            zOrder: PageBorderZOrder.FRONT,
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
