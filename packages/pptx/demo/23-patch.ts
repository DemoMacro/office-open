import { writeFileSync } from "node:fs";

import type { PresentationOptions } from "@file/file";
import { generate, patchPresentation, PatchType, TextRun } from "@office-open/pptx";

// Step 1: Create a template PPTX with placeholders
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
            textBody: {
              children: [
                {
                  children: [{ text: "Hello {{name}}!" }],
                },
              ],
            },
          },
        },
      ],
    },
  ],
};

const templateBuffer = await generate(options);

// Step 2: Patch the template — replace {{name}} with bold "World"
const patchedBuffer = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  patches: {
    name: {
      type: PatchType.PARAGRAPH,
      children: [new TextRun({ text: "World", bold: true, fontSize: 24 })],
    },
  },
});

writeFileSync("My Presentation.pptx", patchedBuffer);
