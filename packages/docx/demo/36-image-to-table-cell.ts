// Add image to table cell in a header and body

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const imageData = readFileSync("./demo/images/image1.jpeg") as Uint8Array;

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
                                data: imageData,
                                transformation: { height: "2.6cm", width: "2.6cm" },
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
              {
                cells: [{ children: [] }, { children: [] }],
              },
              {
                cells: [{ children: [] }, { children: [] }],
              },
            ],
          },
        },
      ],
      headers: {
        default: [
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
                                  data: imageData,
                                  transformation: { height: "2.6cm", width: "2.6cm" },
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
                {
                  cells: [{ children: [] }, { children: [] }],
                },
                {
                  cells: [{ children: [] }, { children: [] }],
                },
              ],
            },
          },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
