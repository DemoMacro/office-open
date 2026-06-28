import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Comment Demo",
  creator: "Demo",
  slides: [
    // Slide 1: Two comments from different authors
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Slide with Comments" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Content to review" },
            fill: "ED7D31",
          },
        },
      ],
      comments: [
        {
          author: "Alice Wang",
          text: "The title looks great!",
          x: "5.3cm",
          y: "1.3cm",
          date: "2026-05-15T10:00:00Z",
        },
        {
          author: "Bob Li",
          text: "Consider changing the content color.",
          x: "7.9cm",
          y: "4.8cm",
          initials: "BL",
          date: "2026-05-15T11:00:00Z",
        },
      ],
    },

    // Slide 2: Comment from existing author (tests author reuse)
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Another Slide" },
            fill: "70AD47",
          },
        },
      ],
      comments: [
        {
          author: "Alice Wang",
          text: "Please review this slide carefully.",
          x: "5.3cm",
          y: "1.3cm",
          date: "2026-05-15T12:00:00Z",
        },
      ],
    },

    // Slide 3: No comments
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "No Comments Here" },
            fill: "FFC000",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
