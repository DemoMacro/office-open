// Endnotes

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  endnotes: {
    1: { children: ["This is the first endnote with some detailed explanation."] },
    2: {
      children: ["Second endnote", "With multiple paragraphs for more complex content."],
    },
    3: { children: ["Third endnote referencing important source material."] },
    4: { children: ["Fourth endnote from a different section."] },
  },
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
        endnotePr: {
          numRestart: "eachSect",
          pos: "docEnd",
        },
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
        endnotePr: {
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
writeFileSync("My Document.docx", buffer);
