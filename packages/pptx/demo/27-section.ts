// Slide sections: group slides by name into p14:section entries (presentation.xml).
// Slides sharing a section name form one p14:section; slides without a section
// stay ungrouped (absent from p14:sectionLst).

import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Slide Sections Demo",
  creator: "Demo",
  slides: [
    // Introduction section — two slides share one section
    {
      section: "Introduction",
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 600,
            height: 200,
            textBody: { text: "Introduction - Slide 1" },
          },
        },
      ],
    },
    {
      section: "Introduction",
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 600,
            height: 200,
            textBody: { text: "Introduction - Slide 2" },
          },
        },
      ],
    },
    // Content section
    {
      section: "Content",
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 600,
            height: 200,
            textBody: { text: "Content - Slide 3" },
          },
        },
      ],
    },
    // No section — this slide stays ungrouped (not referenced by p14:sectionLst)
    {
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 600,
            height: 200,
            textBody: { text: "Ungrouped Slide 4" },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
