import { writeFileSync } from "node:fs";

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
            x: 80,
            y: 120,
            width: 720,
            height: 80,
            textBody: { children: [{ children: [{ text: "Hello {{name}}!" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: 80,
            y: 120,
            width: 720,
            height: 80,
            textBody: { children: [{ children: [{ text: "Original second slide" }] }] },
          },
        },
      ],
    },
  ],
};

const templateBuffer = await generatePresentation(options);

// Step 2: Patch — replace {{name}}, replace the second slide, append a new one.
//          slides.replace keys are 0-based indices in sldIdLst order;
//          slides.append adds slides after the last existing one. Appended
//          slides inherit the template's first slide layout.
const patchedBuffer = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  placeholders: {
    name: [{ text: "World", bold: true, size: 24 }],
  },
  slides: {
    replace: {
      1: {
        children: [
          {
            shape: {
              x: 80,
              y: 120,
              width: 720,
              height: 80,
              textBody: { children: [{ children: [{ text: "Replaced second slide" }] }] },
            },
          },
        ],
      },
    },
    append: [
      {
        children: [
          {
            shape: {
              x: 80,
              y: 120,
              width: 720,
              height: 80,
              textBody: { children: [{ children: [{ text: "Appended third slide" }] }] },
            },
          },
        ],
      },
    ],
  },
});

writeFileSync("My Presentation.pptx", patchedBuffer);
