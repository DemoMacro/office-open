// Custom XML range markers: insert, delete, and move ranges for
// track changes on custom XML markup.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ bold: true, text: "Custom XML Range Markers", size: 16 }],
          },
        },

        // Insert range
        {
          paragraph: {
            children: [
              "Custom XML insert: ",
              {
                customXmlInsRangeStart: { id: 100, author: "Alice", date: "2026-06-04T10:00:00Z" },
              },
              "newly inserted content",
              { customXmlInsRangeEnd: 100 },
              ".",
            ],
          },
        },

        // Delete range
        {
          paragraph: {
            children: [
              "Custom XML delete: ",
              { customXmlDelRangeStart: { id: 101, author: "Bob", date: "2026-06-04T11:00:00Z" } },
              "deleted content",
              { customXmlDelRangeEnd: 101 },
              ".",
            ],
          },
        },

        // Move range
        {
          paragraph: {
            children: [
              "Move from: ",
              {
                customXmlMoveFromRangeStart: {
                  id: 102,
                  author: "Alice",
                  date: "2026-06-04T12:00:00Z",
                },
              },
              "original location",
              { customXmlMoveFromRangeEnd: 102 },
              " -> moved to: ",
              {
                customXmlMoveToRangeStart: {
                  id: 103,
                  author: "Alice",
                  date: "2026-06-04T12:00:00Z",
                },
              },
              "new location",
              { customXmlMoveToRangeEnd: 103 },
              ".",
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
