// Simple example to add comments to a document

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  comments: {
    children: [
      {
        author: "Ray Chen",
        children: [
          {
            children: [
              {
                text: "some initial text content",
              },
            ],
          },
          {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  transformation: {
                    height: "2.6cm",
                    width: "2.6cm",
                  },
                  type: "jpg",
                },
              },
              {
                text: "comment text content",
              },
              { break: 1, text: "" },
              {
                bold: true,
                text: "More text here",
              },
            ],
          },
        ],
        date: new Date(),
        id: 0,
      },
      {
        author: "Bob Ross",
        children: [
          {
            children: [
              {
                text: "Some initial text content",
              },
            ],
          },
          {
            children: [
              {
                text: "comment text content",
              },
            ],
          },
        ],
        date: new Date(),
        id: 1,
      },
      {
        author: "John Doe",
        children: [
          {
            children: [
              {
                text: "Hello World",
              },
            ],
          },
        ],
        date: new Date(),
        id: 2,
      },
      {
        author: "Beatriz",
        children: [
          {
            children: [
              {
                text: "Another reply",
              },
            ],
          },
        ],
        date: new Date(),
        id: 3,
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              { commentRangeStart: { id: 0 } },
              {
                bold: true,
                text: "Foo Bar",
              },
              { commentRangeEnd: { id: 0 } },
              {
                bold: true,
                children: [{ commentReference: 0 }],
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              { commentRangeStart: { id: 1 } },
              { commentRangeStart: { id: 2 } },
              { commentRangeStart: { id: 3 } },
              {
                bold: true,
                text: "Some text which need commenting",
              },
              { commentRangeEnd: { id: 1 } },
              {
                bold: true,
                children: [{ commentReference: 1 }],
              },
              { commentRangeEnd: { id: 2 } },
              {
                bold: true,
                children: [{ commentReference: 2 }],
              },
              { commentRangeEnd: { id: 3 } },
              {
                bold: true,
                children: [{ commentReference: 3 }],
              },
            ],
          },
        },
        {
          // `{ comment }` sugar — the library allocates the id, pairs the range
          // markers + reference, and registers the comment entry. No manual id.
          paragraph: {
            children: [
              "Before comment, ",
              {
                comment: {
                  author: "Sugar Author",
                  initials: "SA",
                  date: new Date(),
                  children: ["Added via the { comment } sugar — no manual id."],
                  wrap: [{ text: "sugar-wrapped text", bold: true }],
                },
              },
              " after comment.",
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
