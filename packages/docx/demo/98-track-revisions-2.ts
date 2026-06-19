// Track Revisions for paragraph properties, section properties, tables
// Docs/usage/change-tracking.md

import { writeFileSync } from "node:fs";

import {
  AlignmentType,
  BorderStyle,
  generateDocument,
  HeightRule,
  VerticalAlignTable,
  WidthType,
} from "@office-open/docx";

const REVISION_DATE = "2020-10-06T09:00:00Z";
const REVISION_AUTHOR = "Firstname Lastname";

const buffer = await generateDocument({
  features: {
    trackRevisions: true,
  },
  sections: [
    {
      children: [
        // Paragraph properties revision
        {
          paragraph: {
            alignment: AlignmentType.CENTER,
            children: [
              {
                text: "Paragraph properties revision",
                bold: true,
              },
            ],
            heading: "Heading1",
            revision: {
              alignment: AlignmentType.LEFT,
              author: REVISION_AUTHOR,
              date: REVISION_DATE,
              id: 10,
            },
          },
        },
        { paragraph: { text: "" } },

        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Table 1: Inserted and Deleted Table Row",
              },
            ],
          },
        },
        {
          table: {
            columnWidths: [2000, 2000],
            layout: "fixed",
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Cell 1" }] },
                  { children: [{ paragraph: "Cell 2" }] },
                ],
              },
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          children: ["Inserted row cell"],
                          run: {
                            insertion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 0,
                            },
                          },
                        },
                      },
                    ],
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          children: ["Inserted row cell"],
                          run: {
                            insertion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 0,
                            },
                          },
                        },
                      },
                    ],
                  },
                ],
                insertion: {
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  id: 0,
                },
              },
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          children: ["Deleted row cell"],
                          run: {
                            deletion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 1,
                            },
                          },
                        },
                      },
                    ],
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          children: ["Deleted row cell"],
                          run: {
                            deletion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 1,
                            },
                          },
                        },
                      },
                    ],
                  },
                ],
                deletion: {
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  id: 1,
                },
              },
            ],
          },
        },
        { paragraph: { text: "" } },

        {
          paragraph: {
            children: [{ bold: true, text: "Table 2: Inserted  Table Column" }],
          },
        },
        {
          table: {
            columnWidths: [2000, 2000],
            layout: "fixed",
            rows: [
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              insertion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 2,
                                children: [{ text: "Inserted cell" }],
                              },
                            },
                          ],
                          run: {
                            insertion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 2,
                            },
                          },
                        },
                      },
                    ],
                    insertion: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 2,
                    },
                  },
                  { children: [{ paragraph: "Cell" }] },
                ],
              },
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              insertion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 2,
                                children: [{ text: "Inserted cell" }],
                              },
                            },
                          ],
                          run: {
                            insertion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 2,
                            },
                          },
                        },
                      },
                    ],
                    insertion: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 2,
                    },
                  },
                  { children: [{ paragraph: "Cell" }] },
                ],
              },
            ],
          },
        },
        { paragraph: { text: "" } },

        {
          paragraph: {
            children: [{ bold: true, text: "Table 3: Deleted Table Column" }],
          },
        },
        {
          table: {
            columnWidths: [2000, 2000],
            layout: "fixed",
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Cell" }] },
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              deletion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 3,
                                children: [{ text: "Deleted cell" }],
                              },
                            },
                          ],
                          run: {
                            deletion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 3,
                            },
                          },
                        },
                      },
                    ],
                    deletion: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 3,
                    },
                  },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Cell" }] },
                  {
                    children: [
                      {
                        paragraph: {
                          children: [
                            {
                              deletion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 3,
                                children: [{ text: "Deleted cell" }],
                              },
                            },
                          ],
                          run: {
                            deletion: {
                              author: REVISION_AUTHOR,
                              date: REVISION_DATE,
                              id: 3,
                            },
                          },
                        },
                      },
                    ],
                    deletion: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 3,
                    },
                  },
                ],
              },
            ],
          },
        },
        { paragraph: { text: "" } },

        {
          paragraph: {
            children: [{ bold: true, text: "Table 3: cell properties revision" }],
          },
        },
        {
          table: {
            columnWidths: [2000, 2000],
            rows: [
              {
                cells: [
                  {
                    width: { size: 4000, type: WidthType.DXA },
                    shading: {
                      color: "#00FF00",
                      fill: "#00FF00",
                    },
                    verticalAlign: VerticalAlignTable.CENTER,
                    revision: {
                      width: { size: 2000, type: WidthType.DXA },
                      id: 4,
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      verticalAlign: VerticalAlignTable.TOP,
                    },
                    children: [{ paragraph: "Cell 1" }],
                  },
                  {
                    width: { size: 2000, type: WidthType.DXA },
                    shading: {
                      color: "#00FF00",
                      fill: "#00FF00",
                    },
                    revision: {
                      width: { size: 2000, type: WidthType.DXA },
                      id: 4,
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                    },
                    children: [{ paragraph: "Cell 2" }],
                  },
                ],
                height: { rule: HeightRule.EXACT, value: 600 },
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "Cell 3" }],
                    revision: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 4,
                      width: { size: 2000, type: WidthType.DXA },
                    },
                    width: { size: 4000, type: WidthType.DXA },
                  },
                  {
                    children: [{ paragraph: "Cell 4" }],
                    revision: {
                      author: REVISION_AUTHOR,
                      date: REVISION_DATE,
                      id: 4,
                      width: { size: 2000, type: WidthType.DXA },
                    },
                    width: { size: 2000, type: WidthType.DXA },
                  },
                ],
              },
            ],
          },
        },
        { paragraph: { text: "" } },

        {
          paragraph: {
            children: [{ bold: true, text: "Table 4: row properties revision" }],
          },
        },
        {
          table: {
            columnWidths: [2000, 2000],
            layout: "fixed",
            rows: [
              {
                cantSplit: true,
                cells: [
                  { children: [{ paragraph: "Cell 1" }] },
                  { children: [{ paragraph: "Cell 2" }] },
                ],
                height: { rule: HeightRule.EXACT, value: 600 },
                revision: {
                  author: REVISION_AUTHOR,
                  cantSplit: false,
                  date: REVISION_DATE,
                  id: 5,
                  tableHeader: false,
                },
                tableHeader: true,
              },
              {
                cells: [
                  { children: [{ paragraph: "Cell 3" }] },
                  { children: [{ paragraph: "Cell 4" }] },
                ],
              },
            ],
          },
        },
        { paragraph: { text: "" } },

        {
          table: {
            borders: {
              bottom: {
                color: "FF0000",
                size: 5,
                style: BorderStyle.DASHED,
              },
              left: {
                color: "FF0000",
                size: 5,
                style: BorderStyle.DASHED,
              },
              right: {
                color: "FF0000",
                size: 5,
                style: BorderStyle.DASHED,
              },
              top: {
                color: "FF0000",
                size: 5,
                style: BorderStyle.DASHED,
              },
            },
            columnWidths: [2000, 2000],
            layout: "fixed",
            revision: {
              author: REVISION_AUTHOR,
              borders: {
                bottom: {
                  color: "00FF00",
                  size: 3,
                  style: BorderStyle.DOTTED,
                },
                left: {
                  color: "00FF00",
                  size: 3,
                  style: BorderStyle.DOTTED,
                },
                right: {
                  color: "00FF00",
                  size: 3,
                  style: BorderStyle.DOTTED,
                },
                top: {
                  color: "00FF00",
                  size: 3,
                  style: BorderStyle.DOTTED,
                },
              },
              date: REVISION_DATE,
              id: 6,
            },
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Cell 1" }] },
                  { children: [{ paragraph: "Cell 2" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Cell 3" }] },
                  { children: [{ paragraph: "Cell 4" }] },
                ],
              },
            ],
          },
        },
        { paragraph: { text: "" } },
      ],
      properties: {},
    },
    // Section properties revision
    {
      children: [
        { paragraph: { text: "Paragraph in different section with revision properties" } },
      ],
      properties: {
        page: {
          margin: {
            bottom: 2440,
            left: 2440,
            right: 2440,
            top: 2440,
          },
          size: {
            height: 11_909,
            width: 16_834,
          },
        },
        revision: {
          author: REVISION_AUTHOR,
          date: REVISION_DATE,
          id: 20,
          page: {
            margin: {
              bottom: 1440,
              left: 1440,
              right: 1440,
              top: 1440,
            },
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
