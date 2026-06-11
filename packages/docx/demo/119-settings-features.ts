// Document settings features: view, zoom, write protection, display background shape,
// font embedding, document variables

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  background: {
    color: "C45911",
  },
  displayBackgroundShape: true,
  view: "print",
  zoom: {
    percent: 150,
  },
  embedTrueTypeFonts: true,
  saveSubsetFonts: true,
  docVars: [
    { name: "Title", val: "Settings Demo" },
    { name: "Version", val: "1.0" },
    { name: "Author", val: "Test User" },
  ],
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Document Settings Demo"],
          },
        },
        {
          paragraph: {
            children: ["This document opens in Print Layout view at 150% zoom."],
          },
        },
        {
          paragraph: {
            children: [
              "The background color is displayed in print layout because displayBackgroundShape is enabled.",
            ],
          },
        },
        {
          paragraph: {
            children: ["TrueType fonts are embedded, and only used subsets are saved."],
          },
        },
        {
          paragraph: {
            children: ["Document variables (Title, Version, Author) are stored in settings.xml."],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
