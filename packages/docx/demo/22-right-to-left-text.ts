// This demo shows right to left for special languages

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            bidirectional: true,
            children: [
              {
                rightToLeft: true,
                text: "שלום עולם",
              },
            ],
          },
        },
        {
          paragraph: {
            bidirectional: true,
            children: [
              {
                bold: true,
                rightToLeft: true,
                text: "שלום עולם",
              },
            ],
          },
        },
        {
          paragraph: {
            bidirectional: true,
            children: [
              {
                italic: true,
                rightToLeft: true,
                text: "שלום עולם",
              },
            ],
          },
        },
        // Bidirectional override: embed LTR text inside RTL paragraph
        {
          paragraph: {
            bidirectional: true,
            children: [
              {
                rightToLeft: true,
                text: "مرحبا ",
              },
              {
                dir: {
                  val: "ltr",
                  children: ["Hello World"],
                },
              },
              {
                rightToLeft: true,
                text: " مرحبا",
              },
            ],
          },
        },
        // BDO: strong bidirectional override
        {
          paragraph: {
            bidirectional: true,
            children: [
              {
                rightToLeft: true,
                text: "نص عربي ",
              },
              {
                bdo: {
                  val: "ltr",
                  children: ["Forced LTR: 123"],
                },
              },
            ],
          },
        },
        {
          table: {
            rows: [
              {
                cells: [{ children: [{ paragraph: "שלום עולם" }] }, { children: [] }],
              },
              {
                cells: [{ children: [] }, { children: [{ paragraph: "שלום עולם" }] }],
              },
            ],
            visuallyRightToLeft: true,
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
