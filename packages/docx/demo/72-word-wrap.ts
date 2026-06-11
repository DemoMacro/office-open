// Example on how to preserve word wrap text. Works with all languages.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "我今天遛狗去公园",
              {
                text: "456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345",
              },
            ],
            wordWrap: true,
          },
        },
        {
          paragraph: {
            children: [
              "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua",
              {
                text: "456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345",
              },
            ],
            wordWrap: true,
          },
        },
        {
          paragraph: {
            children: [
              "我今天遛狗去公园",
              {
                text: "456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345",
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua",
              {
                text: "456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345456435234523456435564745673456345",
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
