import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation, patchPresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Step 1: Create a template PPTX with a placeholder and two slides
const options: PresentationOptions = {
  title: "Patch Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: { paragraphs: [{ children: [{ text: "Hello {{name}}!" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: { paragraphs: [{ children: [{ text: "Original second slide" }] }] },
          },
        },
      ],
    },
  ],
};

const templateBuffer = await generatePresentation(options);

// Step 2: Patch — replace {{name}}, replace the second slide, append a new one.
//          Appended slides inherit the template's first slide layout.
const patchedBuffer = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  commands: [
    {
      type: "replace",
      at: { field: "slides", route: [{ index: 0 }], text: { placeholder: "name" } },
      value: "World",
    },
    {
      type: "replace",
      at: { field: "slides", route: [{ index: 1 }] },
      value: {
        children: [
          {
            shape: {
              x: "2.1cm",
              y: "3.2cm",
              width: "19.1cm",
              height: "2.1cm",
              textBody: { paragraphs: [{ children: [{ text: "Replaced second slide" }] }] },
            },
          },
        ],
      },
    },
    {
      type: "insert",
      at: { field: "slides" },
      position: "end",
      value: [
        {
          children: [
            {
              shape: {
                x: "2.1cm",
                y: "3.2cm",
                width: "19.1cm",
                height: "2.1cm",
                textBody: { paragraphs: [{ children: [{ text: "Appended third slide" }] }] },
              },
            },
          ],
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/23-patch.pptx", patchedBuffer);
