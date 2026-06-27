// Demo: Move Revisions (Track Changes - Move operations)
// Demonstrates MovedFromTextRun, MovedToTextRun, and move range markers.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const REVISION_AUTHOR = "Firstname Lastname";
const REVISION_DATE = "2026-05-02T10:00:00Z";

const buffer = await generateDocument({
  features: {
    trackRevisions: true,
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Move Revisions Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // 1. Simple move: text moved from one location to another
        {
          paragraph: {
            children: [{ bold: true, text: "1. Simple Text Move", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "This paragraph had text moved ",
              {
                movedFrom: {
                  id: 0,
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  children: [{ text: "away from here" }],
                },
              },
              " to the next paragraph.",
            ],
          },
        },
        {
          paragraph: {
            children: [
              "And the text appeared here: ",
              {
                movedTo: {
                  id: 1,
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  children: [{ text: "away from here" }],
                },
              },
              ".",
            ],
          },
        },

        { paragraph: "" },

        // 2. Move with formatting preserved
        {
          paragraph: {
            children: [{ bold: true, text: "2. Move with Formatting", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Source paragraph: ",
              {
                movedFrom: {
                  id: 2,
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  children: [{ text: "bold moved text", bold: true, color: "FF0000" }],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Destination paragraph: ",
              {
                movedTo: {
                  id: 3,
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  children: [{ text: "bold moved text", bold: true, color: "FF0000" }],
                },
              },
            ],
          },
        },

        { paragraph: "" },

        // 3. Move range markers (standalone bookmarks for move regions)
        {
          paragraph: {
            children: [{ bold: true, text: "3. Move Range Markers", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "This paragraph contains ",
              {
                moveFromRangeStart: {
                  id: 4,
                  name: "moved-paragraph",
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                },
              },
              "content that was moved",
              { moveFromRangeEnd: { id: 4 } },
              " elsewhere.",
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Destination with ",
              {
                moveToRangeStart: {
                  id: 5,
                  name: "moved-paragraph",
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                },
              },
              "the moved content",
              { moveToRangeEnd: { id: 5 } },
              " placed here.",
            ],
          },
        },

        { paragraph: "" },

        // 4. Move sugar — the library allocates range + run ids and pairs the
        //    markers. `{ moveFrom }` bundles moveFromRangeStart + the moved run +
        //    moveFromRangeEnd; `{ moveTo }` does the same for the destination.
        {
          paragraph: {
            children: [{ bold: true, text: "4. Move Sugar", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "This paragraph had text moved ",
              {
                moveFrom: {
                  name: "sugared-move",
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  wrap: [{ text: "away from here", italic: true }],
                },
              },
              " to the next paragraph.",
            ],
          },
        },
        {
          paragraph: {
            children: [
              "And the text appeared here: ",
              {
                moveTo: {
                  name: "sugared-move",
                  author: REVISION_AUTHOR,
                  date: REVISION_DATE,
                  wrap: [{ text: "away from here", italic: true }],
                },
              },
              ".",
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
