// Move + offset header and footer

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  evenAndOddHeaderAndFooters: true,
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Hello World 1", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 2", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 3", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 4", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 5", { pageBreak: true }],
          },
        },
      ],
      footers: {
        default: [{ paragraph: { text: "Odd Footer text" } }],
        even: [{ paragraph: { text: "Even Cool Footer text" } }],
      },
      headers: {
        default: [
          { paragraph: { text: "Odd Header text" } },
          { paragraph: { text: "Odd - Some more header text" } },
        ],
        even: [
          { paragraph: { text: "Even header text" } },
          { paragraph: { text: "Even - Some more header text" } },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
