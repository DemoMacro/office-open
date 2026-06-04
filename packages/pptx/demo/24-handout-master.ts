import { writeFileSync } from "node:fs";

import { Packer, Presentation } from "@office-open/pptx";

const pres = new Presentation({
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
});

const buffer = await Packer.toBuffer(pres);
writeFileSync("My Presentation.pptx", buffer);
