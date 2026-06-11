// The demo on the README.md

import { readFileSync, writeFileSync } from "node:fs";

import { HeadingLevel, VerticalAlignTable, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            heading: HeadingLevel.HEADING_1,
            text: "Hello World",
          },
        },
        {
          table: {
            rows: [
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              image: {
                                data: readFileSync("./demo/images/image1.jpeg"),
                                transformation: {
                                  height: 100,
                                  width: 100,
                                },
                                type: "jpg",
                              },
                            },
                          ],
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_1,
                          text: "Hello",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
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
                          text: "World",
                        },
                      },
                    ],
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              image: {
                                data: readFileSync("./demo/images/image1.jpeg"),
                                transformation: {
                                  height: 100,
                                  width: 100,
                                },
                                type: "jpg",
                              },
                            },
                          ],
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: 100,
                    width: 100,
                  },
                  type: "gif",
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
