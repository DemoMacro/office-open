// Add image to table cell

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }, { children: [] }],
              },
              {
                cells: [
                  { children: [] },
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              image: {
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
                  { children: [] },
                  { children: [] },
                ],
              },
              {
                cells: [
                  { children: [] },
                  { children: [] },
                  { children: [{ paragraph: "Hello" }] },
                  { children: [] },
                ],
              },
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }, { children: [] }],
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
