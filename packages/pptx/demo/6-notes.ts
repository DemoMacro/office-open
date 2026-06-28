import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Notes Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Slide 1 - Introduction" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Welcome to the presentation!" },
            fill: "D9E2F3",
          },
        },
      ],
      notes:
        "These are the speaker notes for the introduction slide. Remember to greet the audience.",
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Slide 2 - Key Points" },
            fill: "ED7D31",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Point 1: Architecture\nPoint 2: Implementation\nPoint 3: Testing" },
            fill: "FBE5D6",
          },
        },
      ],
      notes:
        "Key talking points:\n- Architecture follows SOLID principles\n- Implementation uses OOXML spec\n- Testing covers all chart types",
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Slide 3 - No Notes" },
            fill: "70AD47",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
