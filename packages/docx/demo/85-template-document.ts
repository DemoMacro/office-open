// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import { HeadingLevel, TextDirection, VerticalAlignTable, patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template.docx"),
  outputType: "nodebuffer",
  placeholders: {
    footer_text: {
      children: [
        "replaced just as",
        " well",
        { hyperlink: { link: "https://www.bbc.co.uk/news", children: ["BBC News Link"] } },
      ],
      type: "paragraph",
    },
    header_adjective: {
      children: ["Delightful Header"],
      type: "paragraph",
    },
    image_test: {
      children: [
        {
          image: {
            data: readFileSync("./demo/images/image1.jpeg"),
            transformation: { height: 100, width: 100 },
            type: "jpg",
          },
        },
      ],
      type: "paragraph",
    },
    item_1: {
      children: [
        "#657",
        { hyperlink: { link: "https://www.bbc.co.uk/news", children: ["BBC News Link"] } },
      ],
      type: "paragraph",
    },
    name: {
      children: ["Sir. ", "John Doe", "(The Conqueror)"],
      type: "paragraph",
    },
    paragraph_replace: {
      children: [
        { paragraph: "Lorem ipsum paragraph" },
        { paragraph: "Another paragraph" },
        {
          paragraph: {
            children: [
              "This is a ",
              { hyperlink: { link: "https://www.google.co.uk", children: ["Google Link"] } },
              {
                image: {
                  data: readFileSync("./demo/images/dog.png"),
                  transformation: { height: 100, width: 100 },
                  type: "png",
                },
              },
            ],
          },
        },
      ],
      type: "document",
    },
    table: {
      children: [
        {
          table: {
            rows: [
              {
                cells: [
                  {
                    children: [{ paragraph: {} }, { paragraph: {} }],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [{ paragraph: {} }, { paragraph: {} }],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [{ paragraph: { text: "bottom to top" } }, { paragraph: {} }],
                    textDirection: TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT,
                  },
                  {
                    children: [{ paragraph: { text: "top to bottom" } }, { paragraph: {} }],
                    textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_1,
                          text: "Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah",
                        },
                      },
                    ],
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "This text should be in the middle of the cell",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "Text above should be vertical from bottom to top",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "Text above should be vertical from top to bottom",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                ],
              },
            ],
          },
        },
      ],
      type: "document",
    },
    table_heading_1: {
      children: ["Heading wow!"],
      type: "paragraph",
    },
  },
});
writeFileSync("My Document.docx", doc);
