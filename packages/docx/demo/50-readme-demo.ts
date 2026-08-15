// The demo on the README.md

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

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
                              picture: {
                                data: readFileSync("./demo/images/image1.jpeg"),
                                transformation: {
                                  height: "2.6cm",
                                  width: "2.6cm",
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
                              picture: {
                                data: readFileSync("./demo/images/image1.jpeg"),
                                transformation: {
                                  height: "2.6cm",
                                  width: "2.6cm",
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
                picture: {
                  data: readFileSync("./demo/images/pizza.gif"),
                  transformation: {
                    height: "2.6cm",
                    width: "2.6cm",
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/50-readme-demo.docx", buffer);
