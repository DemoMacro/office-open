// Exporting the document as a stream

import { createWriteStream } from "node:fs";
import { Readable } from "stream";
import type { ReadableStream as WebReadableStream } from "stream/web";

import { TABLE_BORDERS_NONE, WidthType, generateDocumentStream } from "@office-open/docx";

const table1 = {
  table: {
    columnWidths: [3505, 5505],
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Hello" }],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
      {
        cells: [
          {
            children: [],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [{ paragraph: "World" }],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
    ],
  },
};

const table2 = {
  table: {
    columnWidths: [3505, 5505],
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Foo" }],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
      {
        cells: [
          {
            children: [],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [{ paragraph: "Bar" }],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
    ],
  },
};

const noBorderTable = {
  table: {
    borders: TABLE_BORDERS_NONE,
    rows: [
      {
        cells: [
          {
            children: [table1],
          },
          {
            children: [table2],
          },
        ],
      },
    ],
  },
};

const stream = generateDocumentStream({
  sections: [
    {
      children: [noBorderTable],
      properties: {},
    },
  ],
});
Readable.fromWeb(stream as unknown as WebReadableStream).pipe(
  createWriteStream("My Document.docx"),
);
