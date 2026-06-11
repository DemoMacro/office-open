// Example of using tab stops

import { createWriteStream } from "node:fs";
import { Readable } from "stream";
import type { ReadableStream as WebReadableStream } from "stream/web";

import {
  HeadingLevel,
  TabStopPosition,
  TabStopType,
  generateDocumentStream,
} from "@office-open/docx";

const columnWidth = Math.floor(TabStopPosition.MAX / 4);
const receiptTabStops = [
  // No need to define first left tab column
  // The right aligned tab column position should point to the end of column
  // I.e. in this case
  // (end position of 1st) + (end position of current)
  // ColumnWidth + columnWidth = columnWidth * 2

  { position: columnWidth * 2, type: TabStopType.RIGHT },
  { position: columnWidth * 3, type: TabStopType.RIGHT },
  { position: TabStopPosition.MAX, type: TabStopType.RIGHT },
];
const twoTabStops = [{ position: TabStopPosition.MAX, type: TabStopType.RIGHT }];

const stream = generateDocumentStream({
  defaultTabStop: 0,
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Receipt 001"],
            heading: HeadingLevel.HEADING_1,
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "To Bob.\tBy Alice.",
                bold: true,
              },
            ],
            tabStops: twoTabStops,
          },
        },
        {
          paragraph: {
            children: ["Foo Inc\tBar Inc"],
            tabStops: twoTabStops,
          },
        },
        { paragraph: { text: "" } },
        {
          paragraph: {
            children: [
              {
                text: "Item\tPrice\tQuantity\tSub-total",
                bold: true,
              },
            ],
            tabStops: receiptTabStops,
          },
        },
        {
          paragraph: {
            tabStops: receiptTabStops,
            text: "Item 3\t10\t5\t50",
          },
        },
        {
          paragraph: {
            tabStops: receiptTabStops,
            text: "Item 3\t10\t5\t50",
          },
        },
        {
          paragraph: {
            tabStops: receiptTabStops,
            text: "Item 3\t10\t5\t50",
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "\t\t\tTotal: 200",
                bold: true,
              },
            ],
            tabStops: receiptTabStops,
          },
        },
      ],
      properties: {},
    },
  ],
});
Readable.fromWeb(stream as unknown as WebReadableStream).pipe(
  createWriteStream("My Document.docx"),
);
