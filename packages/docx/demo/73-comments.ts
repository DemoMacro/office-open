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
                    height: 100,
                    width: 100,
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
              { commentRangeStart: 0 },
              {
                bold: true,
                text: "Foo Bar",
              },
              { commentRangeEnd: 0 },
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
              { commentRangeStart: 1 },
              { commentRangeStart: 2 },
              { commentRangeStart: 3 },
              {
                bold: true,
                text: "Some text which need commenting",
              },
              { commentRangeEnd: 1 },
              {
                bold: true,
                children: [{ commentReference: 1 }],
              },
              { commentRangeEnd: 2 },
              {
                bold: true,
                children: [{ commentReference: 2 }],
              },
              { commentRangeEnd: 3 },
              {
                bold: true,
                children: [{ commentReference: 3 }],
              },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
