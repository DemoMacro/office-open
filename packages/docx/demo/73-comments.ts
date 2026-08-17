// Simple example to add comments to a document

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  comments: [
    {
      author: "Ray Chen",
      children: [
        {
          // w14:paraId of the first paragraph keys the commentsExtended entry
          paraId: "1A2B3C01",
          children: [
            {
              text: "some initial text content",
            },
          ],
        },
        {
          children: [
            {
              picture: {
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
      date: new Date().toISOString(),
      id: 0,
    },
    {
      author: "Bob Ross",
      children: [
        {
          paraId: "1A2B3C02",
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
      date: new Date().toISOString(),
      id: 1,
    },
    {
      author: "John Doe",
      children: [
        {
          paraId: "1A2B3C03",
          children: [
            {
              text: "Hello World",
            },
          ],
        },
      ],
      date: new Date().toISOString(),
      id: 2,
    },
    {
      author: "Beatriz",
      children: [
        {
          paraId: "1A2B3C04",
          children: [
            {
              text: "Another reply",
            },
          ],
        },
      ],
      date: new Date().toISOString(),
      id: 3,
    },
  ],
  // word/people.xml — the author registry Word 2013+ writes alongside comments;
  // each entry pairs with comments by exact author-string equality.
  people: [
    { author: "Ray Chen", contact: "ray.chen@example.com" },
    { author: "Bob Ross" },
    { author: "John Doe", contact: "john.doe@example.com" },
    { author: "Beatriz" },
    { author: "Sugar Author" },
  ],
  // word/commentsExtended.xml — resolved state and reply threading, keyed by
  // the w14:paraId of each comment's first paragraph.
  commentsExtended: [
    { paraId: "1A2B3C01", done: false },
    { paraId: "1A2B3C02", done: true },
    { paraId: "1A2B3C03" },
    { paraId: "1A2B3C04", paraIdParent: "1A2B3C02", done: true },
  ],
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
                  date: new Date().toISOString(),
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/73-comments.docx", buffer);
