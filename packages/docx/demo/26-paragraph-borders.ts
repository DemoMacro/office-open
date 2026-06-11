// Creates two paragraphs, one with a border and one without

import { writeFileSync } from "node:fs";

import { BorderStyle, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        { paragraph: "No border!" },
        {
          paragraph: {
            border: {
              bottom: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
              top: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
            },
            text: "I have borders on my top and bottom sides!",
          },
        },
        {
          paragraph: {
            border: {
              top: {
                color: "auto",
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
              },
            },
            text: "",
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "This will ",
              },
              {
                border: {
                  color: "auto",
                  size: 6,
                  space: 1,
                  style: BorderStyle.SINGLE,
                },
                text: "have a border.",
              },
              {
                text: " This will not.",
              },
            ],
          },
        },
        // Bar border: vertical line on the left side (common in legal documents)
        {
          paragraph: {
            border: {
              bar: {
                color: "FF0000",
                size: 6,
                space: 4,
                style: BorderStyle.SINGLE,
              },
            },
            children: ["This paragraph has a red bar border on the left side."],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
