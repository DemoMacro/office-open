// Endnotes

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  endnotes: [
    {
      id: 1,
      children: ["This is the first endnote with some detailed explanation."],
    },
    {
      id: 2,
      children: ["Second endnote", "With multiple paragraphs for more complex content."],
    },
    {
      id: 3,
      children: ["Third endnote referencing important source material."],
    },
    {
      id: 4,
      children: ["Fourth endnote from a different section."],
    },
  ],
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Endnotes Demo Document", bold: true, size: 14 }],
            spacing: { after: 400 },
          },
        },
        {
          paragraph: {
            children: [
              "This document demonstrates endnotes functionality. ",
              "Here is some text with an endnote reference",
              { endnoteReference: 1 },
              ". This allows for detailed citations and references ",
              { endnoteReference: 2 },
              " without cluttering the main text flow.",
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "Endnotes appear at the end of the document, ",
              "unlike footnotes which appear at the bottom of each page",
              { endnoteReference: 3 },
              ". This makes them ideal for academic papers and formal documents.",
            ],
            spacing: { after: 200 },
          },
        },
      ],
      properties: {
        endnoteProperties: {
          numRestart: "eachSect",
          pos: "docEnd",
        },
        pageMargin: {
          bottom: "2.5cm",
          left: "2.5cm",
          right: "2.5cm",
          top: "2.5cm",
        },
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Second Section", bold: true, size: 12 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "This is content from a different section ",
              "with its own endnote reference",
              { endnoteReference: 4 },
              ". Endnotes from all sections appear together at the document end.",
            ],
          },
        },
      ],
      properties: {
        endnoteProperties: {
          numRestart: "continuous",
          pos: "sectEnd",
        },
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Third Section", bold: true, size: 12 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              "This section has no endnote references of its own, ",
              "but demonstrates that sections can have different endnote properties.",
            ],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/97-endnotes.docx", buffer);
