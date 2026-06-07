import { writeFileSync } from "node:fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

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

const buffer = await generate(options);
writeFileSync("My Presentation.pptx", buffer);
