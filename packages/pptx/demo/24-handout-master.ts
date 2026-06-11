import { writeFileSync } from "node:fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  includeHandoutMaster: true,
  handoutMasterOptions: {
    headerFooter: {
      date: true,
      header: true,
      footer: true,
      slideNumber: true,
    },
  },
  slides: [
    {
      children: [
        {
          shape: {
            textBody: { text: "Slide 1 with Handout Master" },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            textBody: { text: "Slide 2" },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
writeFileSync("My Presentation.pptx", buffer);
